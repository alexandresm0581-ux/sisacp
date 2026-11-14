"""
app.py - Interface web do SISACP (Streamlit).

Como o Streamlit funciona (importante para entender este arquivo):
    * Cada vez que você clica em algo, o script INTEIRO roda de novo, de cima
    a baixo. Não existe "evento de clique" como no VBA.
    * st.button(...) devolve True apenas na execução em que foi clicado.
    * Os widgets (st.date_input, st.selectbox...) devolvem o valor atual.

Para rodar:  streamlit run app.py
"""
import math
from datetime import date, datetime
from zoneinfo import ZoneInfo

import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

import calculos as calc
import database as db

br = calc.br
FUSO_BR = ZoneInfo("America/Fortaleza")

def agora_br():
    """Data/hora de Patos-PB, mesmo quando o app está na nuvem (que usa UTC)."""
    return datetime.now(FUSO_BR)

DIAS_SEMANA = ["segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
               "sexta-feira", "sábado", "domingo"]
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]
MES_ABREV = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
COR_ATIVIDADE = {"Ciclismo": "#C86BC8", "Musculação": "#FFC000", "Vôlei": "#92D050"}

# Logo: balança de banheiro (corpo verde, mostrador branco com ponteiro e pés)
LOGO_SVG = (
    '<svg width="42" height="42" viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">'
    '<rect x="5" y="5" width="54" height="54" rx="13" fill="#2E9E5B"/>'
    '<path d="M15 35a17 17 0 0 1 34 0z" fill="#FFFFFF"/>'
    '<path d="M20 31l2 1.2M32 21v2.6M44 31l-2 1.2" stroke="#9AA5A0" stroke-width="1.8" stroke-linecap="round"/>'
    '<line x1="32" y1="35" x2="40" y2="25" stroke="#D9382B" stroke-width="3" stroke-linecap="round"/>'
    '<circle cx="32" cy="35" r="3.2" fill="#2B2B2B"/>'
    '<rect x="17" y="43" width="12" height="9" rx="4" fill="#FFFFFF" opacity="0.9"/>'
    '<rect x="35" y="43" width="12" height="9" rx="4" fill="#FFFFFF" opacity="0.9"/>'
    '</svg>')

st.set_page_config(page_title="SISACP", page_icon="📉", layout="wide")

if not db.banco_existe():
    st.title("SISACP - Primeira vez aqui")
    st.info("Seu banco ainda não existe na nuvem. Envie seu Excel que eu crio ele pra você.")
    up = st.file_uploader("Envie seu Excel (.xlsm)", type=["xlsm", "xlsx"])
    if up is not None:
        with open("01 - Controle de Peso_v3.xlsm", "wb") as f:
            f.write(up.getbuffer())
        st.success("Excel recebido! Importando...")
        import subprocess, sys
        r = subprocess.run([sys.executable, "importar_excel.py", "01 - Controle de Peso_v3.xlsm"],
                           capture_output=True, text=True)
        if r.returncode == 0:
            st.success("Pronto! Banco criado.")
            st.rerun()
        else:
            st.error("Erro ao importar:")
            st.code(r.stdout + "\n" + r.stderr)
    st.stop()

# =====================================================================
# Funções de gráfico (cada uma recebe números e devolve uma figura)
# =====================================================================
def fig_nutrientes(consumo: dict, meta: dict) -> go.Figure:
    nomes = ["Calorias", "Proteínas", "Carboidratos", "Gorduras"]
    chaves = ["kcal", "proteinas", "carboidratos", "gorduras"]
    unidades = ["kcal", "g", "g", "g"]
    cores = ["#00B050", "#C00000", "#1F5FA8", "#ED7D31"]
    pct = [consumo[k] / meta[k] * 100 if meta[k] else 0 for k in chaves]
    # rótulo em duas linhas: percentual + valor absoluto (consumido / meta)
    rotulos = [f"<b>{br(p)}%</b><br>{br(consumo[k], 0 if u == 'kcal' else 1)} / "
               f"{br(meta[k], 0 if u == 'kcal' else 1)} {u}"
               for p, k, u in zip(pct, chaves, unidades)]
    fig = go.Figure(go.Bar(x=nomes, y=pct, marker_color=cores, text=rotulos, textposition="outside"))
    fig.add_hline(y=100, line_dash="dash", line_color="#888")
    fig.update_yaxes(range=[0, max(125, max(pct) + 25)], ticksuffix="%")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
    return fig

def fig_meta_consumo(consumo: float, meta: float, unid: str, cor_c: str, cor_m: str) -> go.Figure:
    fig = go.Figure(go.Bar(
        y=[f"Consumo ({unid})", f"Meta ({unid})"], x=[consumo, meta], orientation="h",
        marker_color=[cor_c, cor_m], text=[br(consumo, 0), br(meta, 0)], textposition="auto"))
    fig.update_yaxes(autorange="reversed")
    fig.update_layout(height=200, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
    return fig

def fig_rosca(r: dict) -> go.Figure:
    fig = go.Figure(go.Pie(values=[max(r["reduzido"], 0), max(r["a_perder"], 0)], hole=0.72,
                           marker_colors=["#FFB066", "#EAEAEA"], textinfo="none", sort=False,
                           hoverinfo="label+value", labels=["Reduzido (kg)", "À perder (kg)"]))
    fig.update_layout(
        height=300, margin=dict(l=10, r=10, t=10, b=10), showlegend=False,
        annotations=[dict(text=f"<b>{br(r['pct_reduzido'] * 100)}%</b><br>"
                               f"{br(r['reduzido'])} / {br(r['reducao_necessaria'])} kg",
                          x=0.5, y=0.5, showarrow=False, font_size=18)])
    return fig

def fig_peso_imc(df, meta_kg: float) -> go.Figure:
    rotulos = [f"{d[8:10]}/{d[5:7]}" for d in df["data"]]
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=rotulos, y=df["peso_kg"], name="Peso (kg)", marker_color="#FFB066",
                text=[br(p) for p in df["peso_kg"]], textposition="outside", secondary_y=False)
    fig.add_scatter(x=rotulos, y=df["imc"], name="IMC (kg/m²)", mode="lines+markers+text",
                    line_color="#4A90D9", text=[br(i) for i in df["imc"]],
                    textposition="top center", secondary_y=True)
    baixo = math.floor((min(meta_kg, df["peso_kg"].min()) - 1) / 10) * 10
    cima = math.ceil(df["peso_kg"].max() / 10) * 10
    fig.update_yaxes(range=[baixo, cima], title_text="kg", secondary_y=False)
    fator = (df["peso_kg"] / df["imc"]).mean()  # = altura², para alinhar os dois eixos
    fig.update_yaxes(range=[baixo / fator - 2, cima / fator - 2], title_text="IMC",
                     secondary_y=True, showgrid=False)  # -2: sobe a linha acima das barras
    fig.update_xaxes(type="category")
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10),
                      legend=dict(orientation="h", y=1.12))
    return fig

def fig_faixa_imc(r: dict) -> go.Figure:
    zonas = [(15, 18.5, "Abaixo do Peso", "#FFE600"), (18.5, 25, "Normal", "#3FBF5A"),
             (25, 30, "Sobrepeso", "#F4A6A0"), (30, 35, "Obesidade G1", "#F08A4B"),
             (35, 40, "Obesidade G2", "#E8604C"), (40, 45, "Obesidade G3", "#E00000")]
    fig = go.Figure()
    for x0, x1, nome, cor in zonas:
        fig.add_shape(type="rect", x0=x0, x1=x1, y0=0, y1=1, fillcolor=cor, line_width=0, layer="below")
        fig.add_annotation(x=(x0 + x1) / 2, y=1.1, text=f"<b>{nome}</b>", showarrow=False,
                           font=dict(size=12, color="#000000"))  # preto: contraste sobre o fundo branco
    marcadores = [("IMC atual", r["imc_atual"], 0.7, "diamond", "#1F5FA8"),
                  ("IMC inicial", r["imc_inicial"], 0.3, "x", "#7A0000"),
                  ("Meta", r["imc_meta"], 0.5, "star", "#FFFFFF")]
    for nome, valor, y, simbolo, cor in marcadores:
        fig.add_trace(go.Scatter(x=[valor], y=[y], mode="markers", name=f"{nome}: {br(valor)}",
                                 marker=dict(symbol=simbolo, size=17, color=cor,
                                             line=dict(color="black", width=1.5))))
    fig.update_xaxes(range=[15, 45], tickvals=[18.5, 25, 30, 35, 40], showgrid=False)
    fig.update_yaxes(visible=False, range=[0, 1.2])  # faixa branca enxuta acima das cores
    fig.update_layout(height=260, margin=dict(l=10, r=10, t=10, b=100),
                      legend=dict(orientation="h", y=-0.3), plot_bgcolor="white")
    return fig

def fig_atividades(df) -> go.Figure:
    fig = go.Figure()
    meses = sorted(df["mes"].unique())
    rotulos = [MES_ABREV[int(m[5:7]) - 1] + "/" + m[2:4] for m in meses]
    for atividade in df["atividade"].unique():
        sub = df[df["atividade"] == atividade].set_index("mes")
        fig.add_bar(x=rotulos, y=[int(sub["qtde"].get(m, 0)) for m in meses], name=atividade,
                    marker_color=COR_ATIVIDADE.get(atividade), text=[int(sub["qtde"].get(m, 0)) or "" for m in meses],
                    textposition="outside")
    fig.update_layout(barmode="group", height=320, margin=dict(l=10, r=10, t=50, b=10),
                      legend=dict(orientation="h", y=1.12), yaxis_title="Sessões")
    return fig

def fig_por_refeicao(df) -> go.Figure:
    fig = go.Figure()
    fig.add_bar(x=df["refeicao"], y=df["consumo_kcal"], name="Consumo (kcal)", marker_color="#00B050",
                text=[br(v, 0) for v in df["consumo_kcal"]], textposition="outside")
    fig.add_bar(x=df["refeicao"], y=df["meta_kcal"], name="Meta (kcal)", marker_color="#C9D9C9",
                text=[br(v, 0) for v in df["meta_kcal"]], textposition="outside")
    topo = max(df["consumo_kcal"].max(), df["meta_kcal"].max())
    fig.update_yaxes(range=[0, topo * 1.15])  # folga para os números acima das colunas
    fig.update_layout(barmode="group", height=320, margin=dict(l=10, r=10, t=50, b=10),
                      legend=dict(orientation="h", y=1.12))
    return fig

# =====================================================================
# PÁGINA: PAINEL (substitui a aba PRINCIPAL)
# =====================================================================
def pagina_painel() -> None:
    agora = agora_br()
    topo, seletor = st.columns([3, 1])
    topo.title("SISACP - Sistema de Acompanhamento e Controle de Peso")
    topo.markdown(f"#### {DIAS_SEMANA[agora.weekday()]}, {agora.day} de {MESES[agora.month - 1]} "
                  f"de {agora.year} — {agora:%H:%M}")
    dia = seletor.date_input("Dia analisado ('Do dia?')", value=agora.date(), format="DD/MM/YYYY")
    ultimo = calc.ultimo_dia_com_registro()
    if ultimo:
        seletor.caption(f"Último dia com alimentação/água registradas: {ultimo:%d/%m/%Y}")

    # ---------------- linha 1: o dia
    consumo, meta = calc.consumo_do_dia(dia), calc.meta_do_dia()
    kcal_c, kcal_m = calc.calorias_ate_ultima_refeicao(dia)
    ml, meta_ml = calc.agua_do_dia(dia)
    if kcal_c == 0 and ml == 0:
        st.info(f"Nenhum registro de alimentação ou água em {dia:%d/%m/%Y}. "
                "Use a página **Registrar** ou escolha outro dia acima.")

    c1, c2, c3 = st.columns([1.2, 1, 1])
    with c1:
        st.subheader("Consumo de nutrientes (do dia)")
        st.plotly_chart(fig_nutrientes(consumo, meta), width="stretch")
    with c2:
        st.subheader("Calorias até a última refeição")
        st.metric("Evolução", f"{br(kcal_c / kcal_m * 100 if kcal_m else 0)}%")
        st.plotly_chart(fig_meta_consumo(kcal_c, kcal_m, "kcal", "#92D050", "#548235"), width="stretch")
    with c3:
        st.subheader("Água (do dia)")
        st.metric("Evolução", f"{br(ml / meta_ml * 100 if meta_ml else 0)}%")
        st.plotly_chart(fig_meta_consumo(ml, meta_ml, "ml", "#7EC8E3", "#2E86C1"), width="stretch")

    st.divider()

    # ---------------- linha 2: peso
    r = calc.resumo_peso(agora.date())
    if r is None:
        st.warning("Ainda não há pesagens registradas.")
    else:
        d1, d2, d3 = st.columns([1, 1, 1.4])
        with d1:
            st.subheader("Dados de referência")
            st.markdown(
                f"**Nome:** {r['nome']}  \n**Idade:** {r['idade']}  \n"
                f"**Altura:** {br(r['altura_m'], 2)} m  \n"
                f"**Peso inicial:** {br(r['peso_inicial'])} kg  \n"
                f"**Peso atual:** {br(r['peso_atual'])} kg  \n"
                f"**Meta de perda:** {br(r['reducao_necessaria'])} kg  \n"
                f"**Meta de peso:** {br(r['peso_meta'])} kg  \n"
                f"**Período de acomp.:** {r['dias_periodo']} dias | {br(r['meses_periodo'], 2)} meses  \n"
                f"**Média de perda:** {br(r['media_kg_dia'], 3)} kg/dia")
        with d2:
            st.subheader("Redução de peso")
            st.plotly_chart(fig_rosca(r), width="stretch")
            st.caption(f"À perder: {br(r['a_perder'])} kg")
        with d3:
            st.subheader("Evolução peso & IMC")
            st.plotly_chart(fig_peso_imc(calc.evolucao_peso(), r["peso_meta"]), width="stretch")

        st.subheader(f"Índice de massa corporal — {calc.classificar_imc(r['imc_atual'])}")
        st.plotly_chart(fig_faixa_imc(r), width="stretch")

    st.divider()

    # ---------------- linha 3: atividades e refeições
    e1, e2 = st.columns(2)
    with e1:
        st.subheader("Atividades físicas realizadas")
        atv = calc.atividades_por_mes()
        if atv.empty:
            st.caption("Sem atividades registradas.")
        else:
            st.plotly_chart(fig_atividades(atv), width="stretch")
    with e2:
        st.subheader(f"Calorias por refeição — {dia:%d/%m/%Y}")
        st.plotly_chart(fig_por_refeicao(calc.consumo_por_refeicao(dia)), width="stretch")

# =====================================================================
# PÁGINA: REGISTRAR (substitui "Coletas dos Dados" + as macros VBA)
# Cada botão faz um INSERT no banco: é o que a macro fazia ao
# "inserir linha na tabela e colar os valores".
# =====================================================================
def lista_com_exclusao(tabela: str, sql: str, chave: str) -> None:
    """Mostra os últimos registros e permite excluir um deles pelo ID.
    (`tabela` vem sempre do código, nunca do usuário, por isso pode ir no SQL.)"""
    df = db.consulta(sql)
    st.dataframe(df, hide_index=True, width="stretch")
    if not df.empty:
        with st.expander("🗑️ Excluir um registro"):
            id_ = st.selectbox("ID do registro", df["id"].tolist(), key=f"sel_{chave}")
            if st.button("Excluir", key=f"del_{chave}"):
                db.executa(f"DELETE FROM {tabela} WHERE id = ?", (int(id_),))
                st.rerun()

def pagina_registrar() -> None:
    st.title("Registrar dados")
    agora = agora_br()
    aba_peso, aba_agua, aba_ativ, aba_alim = st.tabs(
        ["📉 Peso", "💧 Hidratação", "🏃 Atividade física", "🍽️ Alimentação"])

    # ------------------------------------------------ PESO
    with aba_peso:
        ultimo = calc.evolucao_peso()
        padrao = float(ultimo["peso_kg"].iloc[-1]) if not ultimo.empty else 80.0
        c1, c2 = st.columns(2)
        data = c1.date_input("Data", value=agora.date(), format="DD/MM/YYYY", key="peso_data")
        peso = c2.number_input("Peso (kg)", 30.0, 400.0, padrao, step=0.1, format="%.1f", key="peso_kg")
        altura = calc.perfil()["altura_m"]
        st.caption(f"IMC com este peso: {br(calc.imc(peso, altura))} — {calc.classificar_imc(calc.imc(peso, altura))}")
        if st.button("Registrar peso", type="primary"):
            db.executa("INSERT INTO pesagem (data, peso_kg) VALUES (?, ?)", (data.isoformat(), peso))
            st.success("Peso registrado.")
        st.markdown("**Últimas pesagens**")
        lista_com_exclusao("pesagem", "SELECT id, data, peso_kg, imc FROM vw_pesagem ORDER BY data DESC LIMIT 15", "peso")

    # ------------------------------------------------ ÁGUA
    with aba_agua:
        st.markdown("**Registro rápido (agora)**")
        botoes = st.columns(4)
        for col, ml in zip(botoes, [200, 300, 500, 1000]):
            if col.button(f"💧 {ml} ml", key=f"rapido_{ml}"):
                db.executa("INSERT INTO hidratacao (data, hora, volume_ml) VALUES (?,?,?)",
                           (agora.date().isoformat(), agora.strftime("%H:%M"), ml))
                st.success(f"{ml} ml registrados.")
        st.markdown("**Registro manual**")
        c1, c2, c3 = st.columns(3)
        data = c1.date_input("Data", value=agora.date(), format="DD/MM/YYYY", key="agua_data")
        hora = c2.time_input("Hora", value=agora.time().replace(second=0, microsecond=0), key="agua_hora")
        volume = c3.number_input("Quantidade (ml)", 50, 5000, 300, step=50, key="agua_ml")
        if st.button("Registrar água", type="primary"):
            db.executa("INSERT INTO hidratacao (data, hora, volume_ml) VALUES (?,?,?)",
                       (data.isoformat(), hora.strftime("%H:%M"), int(volume)))
            st.success("Hidratação registrada.")
        total, meta = calc.agua_do_dia(agora.date())
        st.progress(min(total / meta, 1.0) if meta else 0.0, text=f"Hoje: {br(total, 0)} / {br(meta, 0)} ml")
        lista_com_exclusao("hidratacao", "SELECT id, data, hora, volume_ml FROM hidratacao "
                                         "ORDER BY data DESC, hora DESC LIMIT 15", "agua")

    # ------------------------------------------------ ATIVIDADE
    with aba_ativ:
        tipos = db.consulta("SELECT DISTINCT atividade FROM atividade_fisica ORDER BY 1")["atividade"].tolist()
        tipos = tipos + [t for t in ["Caminhada", "Corrida", "Natação"] if t not in tipos] + ["Outra..."]
        c1, c2, c3, c4 = st.columns(4)
        data = c1.date_input("Data", value=agora.date(), format="DD/MM/YYYY", key="ativ_data")
        tipo = c2.selectbox("Atividade", tipos, key="ativ_tipo")
        hora = c3.time_input("Hora de início", value=agora.time().replace(second=0, microsecond=0), key="ativ_hora")
        dur = c4.number_input("Duração (min)", 5, 600, 60, step=5, key="ativ_dur")
        if tipo == "Outra...":
            tipo = st.text_input("Nome da atividade", key="ativ_outra").strip()
        k1, k2 = st.columns(2)
        kcal = k1.number_input("Calorias gastas (opcional, 0 = não informar)", 0.0, 5000.0, 0.0, step=10.0, key="ativ_kcal")
        km = k2.number_input("Distância em km (opcional, 0 = não informar)", 0.0, 500.0, 0.0, step=0.5, key="ativ_km")
        if st.button("Registrar atividade", type="primary"):
            if not tipo:
                st.error("Informe o nome da atividade.")
            else:
                db.executa("INSERT INTO atividade_fisica (data, atividade, hora_inicio, duracao_min,"
                           " kcal_gastas, distancia_km) VALUES (?,?,?,?,?,?)",
                           (data.isoformat(), tipo, hora.strftime("%H:%M"), int(dur), kcal or None, km or None))
                st.success("Atividade registrada.")
        lista_com_exclusao("atividade_fisica", "SELECT id, data, hora_inicio, atividade, duracao_min, kcal_gastas, "
                                               "distancia_km FROM atividade_fisica ORDER BY data DESC, hora_inicio DESC LIMIT 15", "ativ")

    # ------------------------------------------------ ALIMENTAÇÃO
    # (sem st.form de propósito: o seletor de itens precisa reagir à refeição escolhida)
    with aba_alim:
        metas = calc.metas_refeicao()
        c1, c2, c3 = st.columns(3)
        data = c1.date_input("Data", value=agora.date(), format="DD/MM/YYYY", key="al_data")
        hora = c2.time_input("Hora", value=agora.time().replace(second=0, microsecond=0), key="al_hora")
        refeicao = c3.selectbox("Refeição", metas["refeicao"].tolist(), key="al_ref")
        subst = st.checkbox("Alimento substituto (item de outra refeição)", key="al_subst")

        if subst:
            itens = db.consulta("SELECT * FROM item_plano ORDER BY refeicao, cod_item")
        else:
            itens = db.consulta("SELECT * FROM item_plano WHERE refeicao = ? ORDER BY cod_item", (refeicao,))
        rotulo = {str(int(r.id)): f"{r.cod_item} | {r.descricao} — {br(r.peso_g, 0)} g, {br(r.kcal, 0)} kcal"
                                              + (f"  [{r.refeicao}]" if subst else "") for r in itens.itertuples()}
        escolhidos = st.multiselect("Itens consumidos", list(rotulo), format_func=rotulo.get, key=f"al_itens_{subst}_{refeicao}")
        porcoes = st.number_input("Porções (vale para todos os itens marcados; 0,5 = metade)",
                                  0.25, 10.0, 1.0, step=0.25, key="al_porcoes")
        if escolhidos:
            ids_escolhidos = [int(i) for i in escolhidos]
            sel = itens[itens["id"].isin(ids_escolhidos)]
            st.caption(f"Total selecionado: {br(sel['kcal'].sum() * porcoes, 0)} kcal | "
                       f"P {br(sel['proteinas_g'].sum() * porcoes)} g | "
                       f"C {br(sel['carboidratos_g'].sum() * porcoes)} g | "
                       f"G {br(sel['gorduras_g'].sum() * porcoes)} g")
        if st.button("Registrar refeição", type="primary", disabled=not escolhidos):
            for item_id in escolhidos:
                db.executa("INSERT INTO consumo (data, hora, refeicao, item_plano_id, substituto, porcoes)"
                           " VALUES (?,?,?,?,?,?)",
                           (data.isoformat(), hora.strftime("%H:%M"), refeicao, int(item_id), int(subst), porcoes))
            st.success(f"{len(escolhidos)} item(ns) registrado(s) em {refeicao}.")
        lista_com_exclusao("consumo", "SELECT id, data, hora, refeicao, descricao, ROUND(kcal,0) AS kcal "
                                      "FROM vw_consumo ORDER BY data DESC, hora DESC, id DESC LIMIT 20", "al")

# =====================================================================
# PÁGINA: PLANO ALIMENTAR (Cardápio + itens + metas)
# =====================================================================
def pagina_plano() -> None:
    st.title("Plano alimentar")
    aba_card, aba_itens, aba_metas = st.tabs(["Cardápio prescrito", "Itens para registro", "Metas por refeição"])
    metas = calc.metas_refeicao()

    with aba_card:
        card = db.consulta("SELECT * FROM cardapio_item ORDER BY id")
        if card.empty:
            st.caption("Cardápio vazio.")
        else:
            st.caption(card["dieta"].iloc[0])
            for m in metas.itertuples():
                sub = card[card["refeicao"] == m.refeicao]
                if sub.empty:
                    continue
                with st.expander(f"{m.refeicao} — {m.horario or ''}", expanded=True):
                    st.dataframe(sub[["descricao", "peso_g", "kcal", "proteinas_g", "carboidratos_g", "gorduras_g"]]
                                 .rename(columns={"descricao": "Item", "peso_g": "Peso (g/ml)", "kcal": "kcal",
                                                  "proteinas_g": "Prot (g)", "carboidratos_g": "Carb (g)",
                                                  "gorduras_g": "Gord (g)"}), hide_index=True, width="stretch")
                    st.caption(f"Total: {br(sub['kcal'].sum(), 0)} kcal (meta da refeição: {br(m.kcal, 0)} kcal)")

    with aba_itens:
        itens = db.consulta("SELECT cod_item AS cod, refeicao, descricao, peso_g, kcal, proteinas_g, "
                            "carboidratos_g, gorduras_g FROM item_plano ORDER BY refeicao, cod_item")
        st.dataframe(itens, hide_index=True, width="stretch")
        with st.expander("➕ Cadastrar novo item"):
            a, b = st.columns(2)
            ref = a.selectbox("Refeição", metas["refeicao"].tolist(), key="ni_ref")
            desc = b.text_input("Descrição", key="ni_desc")
            n = st.columns(5)
            peso = n[0].number_input("Peso (g/ml)", 0.0, 5000.0, 100.0, key="ni_peso")
            kc = n[1].number_input("kcal", 0.0, 5000.0, 0.0, key="ni_kcal")
            pr = n[2].number_input("Proteínas (g)", 0.0, 500.0, 0.0, key="ni_p")
            ca = n[3].number_input("Carboidratos (g)", 0.0, 1000.0, 0.0, key="ni_c")
            go_ = n[4].number_input("Gorduras (g)", 0.0, 500.0, 0.0, key="ni_g")
            if st.button("Salvar item"):
                if not desc.strip():
                    st.error("Informe a descrição.")
                else:
                    prox = int(db.consulta("SELECT COALESCE(MAX(cod_item),0)+1 AS n FROM item_plano")["n"].iloc[0])
                    db.executa("INSERT INTO item_plano (refeicao, cod_item, descricao, peso_g, kcal, proteinas_g,"
                               " carboidratos_g, gorduras_g) VALUES (?,?,?,?,?,?,?,?)",
                               (ref, prox, desc.strip(), peso, kc, pr, ca, go_))
                    st.success(f"Item {prox} cadastrado.")
                    st.rerun()

    with aba_metas:
        st.caption("Edite os valores na tabela e clique em salvar. A meta do dia é a soma das refeições.")
        editado = st.data_editor(
            metas[["refeicao", "horario", "kcal", "proteinas_g", "carboidratos_g", "gorduras_g"]],
            hide_index=True, width="stretch", disabled=["refeicao"], key="editor_metas")
        if st.button("Salvar metas", type="primary"):
            for r in editado.itertuples():
                db.executa("UPDATE meta_refeicao SET horario=?, kcal=?, proteinas_g=?, carboidratos_g=?,"
                           " gorduras_g=? WHERE refeicao=?",
                           (r.horario, r.kcal, r.proteinas_g, r.carboidratos_g, r.gorduras_g, r.refeicao))
            st.success("Metas atualizadas.")
        total = calc.meta_do_dia()
        st.markdown(f"**Meta do dia:** {br(total['kcal'], 0)} kcal | P {br(total['proteinas'])} g | "
                    f"C {br(total['carboidratos'])} g | G {br(total['gorduras'])} g")

# =====================================================================
# PÁGINA: MEDICAMENTOS
# =====================================================================
def pagina_medicamentos() -> None:
    st.title("Medicamentos prescritos")
    df = db.consulta("SELECT id, nome, posologia, uso, data_inicio, data_alteracao, preco, qtde,"
                     " ROUND(preco * qtde, 2) AS custo FROM medicamento ORDER BY nome")
    hoje = agora_br().date()
    df["falta_alterar_dias"] = [
        (date.fromisoformat(d) - hoje).days if isinstance(d, str) else None for d in df["data_alteracao"]]
    st.dataframe(df, hide_index=True, width="stretch")
    st.metric("Custo total (preço × quantidade)", f"R$ {br(df['custo'].sum(), 2)}")
    with st.expander("➕ Cadastrar medicamento"):
        a, b = st.columns(2)
        nome = a.text_input("Medicamento", key="md_nome")
        uso = b.selectbox("Uso", ["Contínuo", "Pontual", "Com prazo"], key="md_uso")
        pos = st.text_area("Posologia", key="md_pos")
        c = st.columns(4)
        preco = c[0].number_input("Preço (R$)", 0.0, 100000.0, 0.0, key="md_preco")
        qtde = c[1].number_input("Qtde.", 1, 100, 1, key="md_qtde")
        ini = c[2].date_input("Início", value=None, format="DD/MM/YYYY", key="md_ini")
        alt = c[3].date_input("Data de alteração", value=None, format="DD/MM/YYYY", key="md_alt")
        if st.button("Salvar medicamento"):
            if not nome.strip():
                st.error("Informe o nome.")
            else:
                db.executa("INSERT INTO medicamento (nome, posologia, preco, qtde, uso, data_inicio, data_alteracao)"
                           " VALUES (?,?,?,?,?,?,?)",
                           (nome.strip(), pos, preco, int(qtde), uso,
                            ini.isoformat() if ini else None, alt.isoformat() if alt else None))
                st.rerun()

# =====================================================================
# PÁGINA: PERFIL
# =====================================================================
def pagina_perfil() -> None:
    st.title("Perfil e metas gerais")
    p = calc.perfil()
    a, b = st.columns(2)
    nome = a.text_input("Nome", p["nome"])
    nasc = b.date_input("Data de nascimento", value=date.fromisoformat(p["data_nascimento"]),
                        min_value=date(1900, 1, 1), format="DD/MM/YYYY")
    c = st.columns(4)
    altura = c[0].number_input("Altura (m)", 1.0, 2.5, float(p["altura_m"]), step=0.01)
    meta = c[1].number_input("Meta de peso (kg)", 30.0, 300.0, float(p["peso_meta_kg"]), step=0.5)
    inicio = c[2].date_input("Início do acompanhamento", value=date.fromisoformat(p["data_inicio_acomp"]),
                             format="DD/MM/YYYY")
    agua = c[3].number_input("Meta de água (ml/dia)", 500, 10000, int(p["meta_agua_ml"]), step=100)
    if st.button("Salvar perfil", type="primary"):
        db.executa("UPDATE perfil SET nome=?, data_nascimento=?, altura_m=?, peso_meta_kg=?,"
                   " data_inicio_acomp=?, meta_agua_ml=? WHERE id=1",
                   (nome, nasc.isoformat(), altura, meta, inicio.isoformat(), int(agua)))
        st.success("Perfil atualizado.")

# =====================================================================
# NAVEGAÇÃO
# =====================================================================
PAGINAS = {
    "📊 Painel": pagina_painel,
    "➕ Registrar": pagina_registrar,
    "🍽️ Plano alimentar": pagina_plano,
    "💊 Medicamentos": pagina_medicamentos,
    "👤 Perfil": pagina_perfil,
}
st.sidebar.markdown(
    f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:0.6rem">{LOGO_SVG}'
    '<span style="font-size:1.9rem;font-weight:700">SISACP</span></div>',
    unsafe_allow_html=True)
escolha = st.sidebar.radio("Navegação", list(PAGINAS), label_visibility="collapsed")
PAGINAS[escolha]()
