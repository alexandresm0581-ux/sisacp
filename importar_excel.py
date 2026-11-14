"""
importar_excel.py - Copia os dados do seu Excel (.xlsm) para o banco SQLite.

Uso:
    python importar_excel.py "01 - Controle de Peso_v3.xlsm"
    python importar_excel.py "01 - Controle de Peso_v3.xlsm" --recriar

ATENÇÃO: este script é para a MIGRAÇÃO INICIAL. Depois que você passar a
registrar dados direto no sistema web, NÃO rode de novo com --recriar,
senão o banco é apagado e refeito só com o que está no Excel.

O que ele faz, em ordem:
  1. Cria o banco vazio (tabelas do schema.sql)
  2. Lê as tabelas do Excel com openpyxl (valores já calculados pelo Excel)
  3. Converte tipos (datas -> 'AAAA-MM-DD', horas -> 'HH:MM', duração -> minutos)
  4. Grava tudo numa única transação (ou entra tudo, ou não entra nada)
"""
import argparse
import sys
import warnings
from datetime import date, datetime, time, timedelta
from pathlib import Path

from openpyxl import load_workbook

import database as db

warnings.filterwarnings("ignore")  # openpyxl avisa sobre gráficos/formas; não afeta os dados

ORDEM_REFEICOES = ["Café da Manhã", "Almoço", "Lanche", "Jantar", "Ceia"]


# ------------------------------------------------------------------ utilitários
def tabela_excel(wb, aba: str, nome_tabela: str) -> list[dict]:
    """Lê uma 'Tabela' do Excel (Inserir > Tabela) como lista de dicionários."""
    ws = wb[aba]
    linhas = [[c.value for c in linha] for linha in ws[ws.tables[nome_tabela].ref]]
    cabecalho = linhas[0]
    return [dict(zip(cabecalho, linha)) for linha in linhas[1:]]


def data_iso(v) -> str | None:
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return None


def hora_hhmm(v) -> str | None:
    if isinstance(v, datetime):
        v = v.time()
    if isinstance(v, time):
        return v.strftime("%H:%M")
    return None


def minutos(v) -> int | None:
    """Duração do Excel (hora, ex.: 01:40:00) -> minutos (100)."""
    if isinstance(v, datetime):
        v = v.time()
    if isinstance(v, time):
        return v.hour * 60 + v.minute
    if isinstance(v, timedelta):
        return int(v.total_seconds() // 60)
    if isinstance(v, (int, float)):  # fração de dia
        return round(v * 24 * 60)
    return None


def numero(v, padrao=0.0) -> float:
    return float(v) if isinstance(v, (int, float)) else padrao


# ------------------------------------------------------------------ importação
def importar(caminho: Path) -> None:
    print(f"Lendo {caminho.name} (pode levar alguns segundos)...")
    wb = load_workbook(caminho, data_only=True, keep_vba=False)
    resumo: dict[str, int] = {}
    avisos: list[str] = []

    con = db.conectar()
    try:
        with con:  # transação única
            # ---- PERFIL
            ws = wb["Controle de Peso"]
            con.execute(
                "INSERT INTO perfil (id, nome, data_nascimento, altura_m, peso_meta_kg,"
                " data_inicio_acomp, meta_agua_ml) VALUES (1,?,?,?,?,?,?)",
                (
                    ws["Q2"].value,
                    data_iso(ws["Q3"].value),
                    float(ws["O2"].value),
                    float(ws["M4"].value),
                    data_iso(ws["Q6"].value),
                    int(wb["Controle de Hidratação"]["N2"].value or 3000),
                ),
            )

            # ---- PESAGENS
            n = 0
            for r in tabela_excel(wb, "Controle de Peso", "Tabela8"):
                if r["Data"] and isinstance(r["Peso (kg)"], (int, float)):
                    con.execute("INSERT INTO pesagem (data, peso_kg) VALUES (?,?)",
                                (data_iso(r["Data"]), float(r["Peso (kg)"])))
                    n += 1
            resumo["pesagem"] = n

            # ---- HIDRATAÇÃO
            n = 0
            for r in tabela_excel(wb, "Controle de Hidratação", "Tabela11"):
                ml = r["Ingestão Hídrica (ml)"]
                if r["Data"] and isinstance(ml, (int, float)):
                    con.execute("INSERT INTO hidratacao (data, hora, volume_ml) VALUES (?,?,?)",
                                (data_iso(r["Data"]), hora_hhmm(r["Hora"]) or "00:00", int(ml)))
                    n += 1
            resumo["hidratacao"] = n

            # ---- ATIVIDADES FÍSICAS
            n = 0
            for r in tabela_excel(wb, "Controle Atividades Fisicas", "Tabela7"):
                dur = minutos(r["Duração"])
                if r["Data"] and r["Atividade"] and dur:
                    con.execute(
                        "INSERT INTO atividade_fisica (data, atividade, hora_inicio, duracao_min)"
                        " VALUES (?,?,?,?)",
                        (data_iso(r["Data"]), r["Atividade"], hora_hhmm(r["Horário"]), dur))
                    n += 1
            resumo["atividade_fisica"] = n

            # ---- HORÁRIOS DAS REFEIÇÕES (vêm do Cardápio Semanal)
            cardapio = wb["Cardápio Semanal"]
            horarios = {}
            for linha in cardapio.iter_rows(min_row=5, max_row=60, min_col=2, max_col=3, values_only=True):
                if linha[0] in ORDEM_REFEICOES and hora_hhmm(linha[1]):
                    horarios[linha[0]] = hora_hhmm(linha[1])

            # ---- METAS POR REFEIÇÃO
            n = 0
            for r in tabela_excel(wb, "Apoio", "Tabela6"):
                if r["Refeição"] in ORDEM_REFEICOES:
                    con.execute(
                        "INSERT INTO meta_refeicao (refeicao, ordem, horario, kcal, proteinas_g,"
                        " carboidratos_g, gorduras_g) VALUES (?,?,?,?,?,?,?)",
                        (r["Refeição"], ORDEM_REFEICOES.index(r["Refeição"]) + 1,
                         horarios.get(r["Refeição"]),
                         round(numero(r["Meta Calorias (kcal)"]), 2),
                         round(numero(r["Meta Proteínas (g)"]), 2),
                         round(numero(r["Meta Carboidratos (g)"]), 2),
                         round(numero(r["Meta Gorduras (g)"]), 2)))
                    n += 1
            resumo["meta_refeicao"] = n

            # ---- ITENS DO PLANO (catálogo de porções)
            n = 0
            for r in tabela_excel(wb, "Apoio", "Tabela5"):
                if r["Refeição"] in ORDEM_REFEICOES and r["Cod Item"] is not None:
                    con.execute(
                        "INSERT INTO item_plano (refeicao, cod_item, descricao, peso_g, kcal,"
                        " proteinas_g, carboidratos_g, gorduras_g) VALUES (?,?,?,?,?,?,?,?)",
                        (r["Refeição"], int(r["Cod Item"]), r["Item"],
                         numero(r["Peso (g ou ml)"]), numero(r["Calorias (kcal)"]),
                         numero(r["Proteínas (g)"]), numero(r["Carboidratos (g)"]),
                         numero(r["Gorduras (g)"])))
                    n += 1
            resumo["item_plano"] = n

            # ---- CONSUMO (aba "Controle Alimentação")
            n = 0
            for r in tabela_excel(wb, "Controle Alimentação", "tb_Controle_Alimentacao5"):
                if not (r["Data"] and r["Cod Item"] is not None and r["Refeição"]):
                    continue
                subst = 1 if str(r["Subst (S/N)?"]).strip().upper() == "S" else 0
                if subst:  # substituto: o Excel procurava só pelo código do item
                    achado = con.execute("SELECT id FROM item_plano WHERE cod_item=? LIMIT 1",
                                         (int(r["Cod Item"]),)).fetchone()
                else:      # normal: chave = refeição + código
                    achado = con.execute("SELECT id FROM item_plano WHERE refeicao=? AND cod_item=?",
                                         (r["Refeição"], int(r["Cod Item"]))).fetchone()
                if not achado:
                    avisos.append(f"Item não encontrado: {data_iso(r['Data'])} {r['Refeição']} cod {r['Cod Item']}")
                    continue
                con.execute(
                    "INSERT INTO consumo (data, hora, refeicao, item_plano_id, substituto, porcoes)"
                    " VALUES (?,?,?,?,?,1)",
                    (data_iso(r["Data"]), hora_hhmm(r["Hora"]) or "00:00", r["Refeição"], achado[0], subst))
                n += 1
            resumo["consumo"] = n

            # ---- MEDICAMENTOS
            n = 0
            for r in tabela_excel(wb, "Medicamentos", "Tabela3"):
                if not r["Medicamento"]:
                    continue
                inicio, alteracao = r["Início"], r["Data Alteração"]
                if "Pontual" in (str(inicio), str(alteracao)):
                    uso = "Pontual"
                elif data_iso(alteracao):
                    uso = "Com prazo"
                else:
                    uso = "Contínuo"
                con.execute(
                    "INSERT INTO medicamento (nome, posologia, preco, qtde, uso, data_inicio, data_alteracao)"
                    " VALUES (?,?,?,?,?,?,?)",
                    (r["Medicamento"], r["Posologia"], numero(r["Preço"]), int(numero(r["Qtde."], 1)),
                     uso, data_iso(inicio), data_iso(alteracao)))
                n += 1
            resumo["medicamento"] = n

            # ---- CARDÁPIO (dieta prescrita)
            dieta = str(cardapio["B2"].value or "Dieta")
            n = 0
            for linha in cardapio.iter_rows(min_row=5, max_row=60, min_col=2, max_col=11, values_only=True):
                refeicao, _, ordem, item, _, peso, kcal, prot, carb, gord = linha
                if refeicao in ORDEM_REFEICOES and isinstance(ordem, (int, float)) and item:
                    con.execute(
                        "INSERT INTO cardapio_item (dieta, refeicao, ordem_item, descricao, peso_g,"
                        " kcal, proteinas_g, carboidratos_g, gorduras_g) VALUES (?,?,?,?,?,?,?,?,?)",
                        (dieta, refeicao, int(ordem), item, numero(peso), numero(kcal),
                         numero(prot), numero(carb), numero(gord)))
                    n += 1
            resumo["cardapio_item"] = n
    finally:
        con.close()

    print("\nImportação concluída. Linhas gravadas por tabela:")
    for tabela, qtd in resumo.items():
        print(f"  {tabela:<18} {qtd:>5}")
    if avisos:
        print(f"\n{len(avisos)} aviso(s):")
        for a in avisos:
            print("  -", a)
    print("\nNão importadas de propósito: abas ocultas 'Lista Alimentos' e 'Cont Alimentação'"
          " (versão antiga do controle, de fev/2026), tabelas dinâmicas e gráficos.")


def main() -> None:
    ap = argparse.ArgumentParser(description="Importa o Excel do SISACP para o SQLite.")
    ap.add_argument("arquivo", nargs="?", help="caminho do .xlsm (padrão: o primeiro .xlsm desta pasta)")
    ap.add_argument("--recriar", action="store_true", help="APAGA o banco atual e importa de novo")
    args = ap.parse_args()

    caminho = Path(args.arquivo) if args.arquivo else next(db.PASTA.glob("*.xlsm"), None)
    if not caminho or not caminho.exists():
        sys.exit("Arquivo Excel não encontrado. Informe o caminho: python importar_excel.py \"arquivo.xlsm\"")

    if db.banco_existe():
        if not args.recriar:
            sys.exit("O banco já existe e pode conter dados novos. Para apagar e reimportar, use --recriar.")
        db.ARQUIVO_DB.unlink()

    db.criar_estrutura()
    importar(caminho)


if __name__ == "__main__":
    main()
