"""
calculos.py - As "fórmulas" do sistema (regras de negócio).

No Excel estas regras estavam espalhadas em células (=M7/M5, =M2-M3...).
Aqui cada uma vira uma função com nome claro. A tela (app.py) só chama
estas funções e desenha o resultado; ela não calcula nada.
"""
from datetime import date

import pandas as pd

import database as db

# Faixas de IMC (OMS): (limite superior exclusivo, nome)
FAIXAS_IMC = [
    (18.5, "Abaixo do Peso"),
    (25.0, "Normal"),
    (30.0, "Sobrepeso"),
    (35.0, "Obesidade Grau 1"),
    (40.0, "Obesidade Grau 2"),
    (float("inf"), "Obesidade Grau 3"),
]


# ------------------------------------------------------------------ perfil
def perfil() -> dict:
    return db.consulta("SELECT * FROM perfil WHERE id = 1").iloc[0].to_dict()


def idade_texto(nascimento: date, hoje: date) -> str:
    """Idade em anos e meses de calendário (ex.: '51 Anos, 11 Meses')."""
    meses = (hoje.year - nascimento.year) * 12 + (hoje.month - nascimento.month)
    if hoje.day < nascimento.day:
        meses -= 1
    return f"{meses // 12} Anos, {meses % 12} Meses"


def classificar_imc(imc: float) -> str:
    for limite, nome in FAIXAS_IMC:
        if imc < limite:
            return nome
    return FAIXAS_IMC[-1][1]


def imc(peso_kg: float, altura_m: float) -> float:
    return peso_kg / (altura_m ** 2)


# ------------------------------------------------------------------ peso
def evolucao_peso() -> pd.DataFrame:
    return db.consulta("SELECT data, peso_kg, imc FROM vw_pesagem ORDER BY data")


def resumo_peso(hoje: date) -> dict | None:
    """Todos os números do bloco 'Dados de Referência' + rosca de redução."""
    p = perfil()
    pes = evolucao_peso()
    if pes.empty:
        return None
    inicio = date.fromisoformat(p["data_inicio_acomp"])
    ultima_data = date.fromisoformat(pes["data"].iloc[-1])

    inicial = float(pes["peso_kg"].iloc[0])
    atual = float(pes["peso_kg"].iloc[-1])
    meta = float(p["peso_meta_kg"])

    reducao_necessaria = inicial - meta        # Excel: Meta Perda (kg)
    reduzido = inicial - atual                 # Excel: Reduzido (kg)
    dias_ate_ultima = (ultima_data - inicio).days
    dias_periodo = (hoje - inicio).days

    return {
        "nome": p["nome"],
        "idade": idade_texto(date.fromisoformat(p["data_nascimento"]), hoje),
        "altura_m": p["altura_m"],
        "peso_inicial": inicial,
        "peso_atual": atual,
        "peso_meta": meta,
        "reducao_necessaria": reducao_necessaria,
        "reduzido": reduzido,
        "a_perder": atual - meta,
        "pct_reduzido": reduzido / reducao_necessaria if reducao_necessaria else 0.0,
        "dias_periodo": dias_periodo,
        "meses_periodo": dias_periodo / 30,
        "media_kg_dia": reduzido / dias_ate_ultima if dias_ate_ultima > 0 else 0.0,
        "imc_inicial": imc(inicial, p["altura_m"]),
        "imc_atual": imc(atual, p["altura_m"]),
        "imc_meta": imc(meta, p["altura_m"]),
    }


# ------------------------------------------------------------------ alimentação
def metas_refeicao() -> pd.DataFrame:
    return db.consulta("SELECT * FROM meta_refeicao ORDER BY ordem")


def meta_do_dia() -> dict:
    m = metas_refeicao()
    return {"kcal": float(m["kcal"].sum()), "proteinas": float(m["proteinas_g"].sum()),
            "carboidratos": float(m["carboidratos_g"].sum()), "gorduras": float(m["gorduras_g"].sum())}


def consumo_do_dia(dia: date) -> dict:
    df = db.consulta(
        "SELECT COALESCE(SUM(kcal),0) kcal, COALESCE(SUM(proteinas_g),0) proteinas,"
        " COALESCE(SUM(carboidratos_g),0) carboidratos, COALESCE(SUM(gorduras_g),0) gorduras"
        " FROM vw_consumo WHERE data = ?", (dia.isoformat(),))
    return df.iloc[0].to_dict()


def calorias_ate_ultima_refeicao(dia: date) -> tuple[float, float]:
    """(consumido, meta) considerando só as refeições até a última registrada no dia."""
    df = db.consulta(
        "SELECT COALESCE(SUM(m.kcal),0) AS meta FROM meta_refeicao m WHERE m.ordem <= "
        "(SELECT COALESCE(MAX(r.ordem),0) FROM consumo c JOIN meta_refeicao r "
        " ON r.refeicao = c.refeicao WHERE c.data = ?)", (dia.isoformat(),))
    return consumo_do_dia(dia)["kcal"], float(df["meta"].iloc[0])


def consumo_por_refeicao(dia: date) -> pd.DataFrame:
    """Consumido x meta de cada refeição no dia (aba 'Evolução do dia')."""
    return db.consulta(
        "SELECT m.refeicao, m.ordem, m.kcal AS meta_kcal, COALESCE(SUM(v.kcal),0) AS consumo_kcal,"
        " m.proteinas_g AS meta_prot, COALESCE(SUM(v.proteinas_g),0) AS consumo_prot"
        " FROM meta_refeicao m LEFT JOIN vw_consumo v ON v.refeicao = m.refeicao AND v.data = ?"
        " GROUP BY m.refeicao ORDER BY m.ordem", (dia.isoformat(),))


# ------------------------------------------------------------------ hidratação
def agua_do_dia(dia: date) -> tuple[int, int]:
    df = db.consulta("SELECT COALESCE(SUM(volume_ml),0) AS ml FROM hidratacao WHERE data = ?",
                     (dia.isoformat(),))
    return int(df["ml"].iloc[0]), int(perfil()["meta_agua_ml"])


# ------------------------------------------------------------------ atividades
def atividades_por_mes() -> pd.DataFrame:
    return db.consulta(
        "SELECT strftime('%Y-%m', data) AS mes, atividade, COUNT(*) AS qtde,"
        " SUM(duracao_min) AS minutos FROM atividade_fisica GROUP BY mes, atividade ORDER BY mes")


# ------------------------------------------------------------------ utilidades
def ultimo_dia_com_registro() -> date | None:
    df = db.consulta("SELECT MAX(d) AS d FROM (SELECT MAX(data) d FROM consumo"
                     " UNION SELECT MAX(data) FROM hidratacao)")
    v = df["d"].iloc[0]
    return date.fromisoformat(v) if v else None


def br(valor: float, casas: int = 1) -> str:
    """Formata no padrão brasileiro: 1.234,5"""
    s = f"{valor:,.{casas}f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")
