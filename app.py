import streamlit as st
import pandas as pd
import altair as alt
import os

st.set_page_config(page_title="Excesso de Velocidade", layout="wide")
st.title("Painel – Excesso de Velocidade")

ARQUIVO_EXCEL = "Excel_Motor_Dashboard.xlsm"

MESES_ORDEM = [
    "janeiro","fevereiro","março","abril","maio","junho",
    "julho","agosto","setembro","outubro","novembro","dezembro"
]
MESES_ROTULO = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

SEMANAS_ORDEM = ["1 Semana","2 Semana","3 Semana","4 Semana","5 Semana"]
TIPOS_ORDEM = ["Asfalto", "Terra", "Palmeira"]

CORES_TIPO = {
    "Asfalto": "#808080",
    "Terra":   "#8B4513",
    "Palmeira":"#2E8B57"
}

def norm(s: str) -> str:
    return str(s).strip()

@st.cache_data
def listar_abas_norm():
    xls = pd.ExcelFile(ARQUIVO_EXCEL)
    orig = xls.sheet_names
    norm_map = {norm(n): n for n in orig}
    return orig, norm_map

@st.cache_data
def carregar_aba(nome_aba_real: str) -> pd.DataFrame:
    df = pd.read_excel(ARQUIVO_EXCEL, sheet_name=nome_aba_real)
    df.columns = [str(c).strip() for c in df.columns]

    obrig = {"Data_Inicio", "Mes", "Semana", "Motorista", "Tipo", "Quantidade", "Gestor"}
    faltando = sorted(list(obrig - set(df.columns)))
    if faltando:
        raise KeyError(f"Colunas faltando na aba {nome_aba_real}: {', '.join(faltando)}")

    df["Data_Inicio"] = pd.to_datetime(df["Data_Inicio"], errors="coerce", dayfirst=True)
    df = df.dropna(subset=["Data_Inicio"]).copy()

    df["Mes"] = df["Mes"].astype(str).str.strip().str.lower()
    df["Ano"] = df["Data_Inicio"].dt.year
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce").fillna(0)

    df["Semana"] = df["Semana"].astype(str).str.strip()
    df["Tipo"] = df["Tipo"].astype(str).str.strip()
    df["Gestor"] = df["Gestor"].astype(str).str.strip()
    df["Motorista"] = df["Motorista"].astype(str).str.strip()

    return df

@st.cache_data
def carregar_parametros_gersup():
    df_gs = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Parametros_GerSup")
    df_gs.columns = [str(c).strip() for c in df_gs.columns]

    obrig = {"Nome Modificado", "Nome Real"}
    faltando = obrig - set(df_gs.columns)
    if faltando:
        raise KeyError(f"Colunas faltando em Parametros_GerSup: {', '.join(sorted(list(faltando)))}")

    df_gs["Nome Modificado"] = df_gs["Nome Modificado"].astype(str).str.strip()
    df_gs["Nome Real"] = df_gs["Nome Real"].astype(str).str.strip()

    mapa = dict(zip(df_gs["Nome Modificado"], df_gs["Nome Real"]))
    lista_gestores = sorted([g for g in df_gs["Nome Real"].unique().tolist() if str(g).strip() != ""])

    return mapa, lista_gestores

def aplicar_estilo_tabela(df_show: pd.DataFrame):
    header_bg = "#0b2a4a"
    header_fg = "#ffffff"
    styler = df_show.style
    styler = styler.set_table_styles([
        {"selector": "thead th",
         "props": [("background-color", header_bg),
                   ("color", header_fg),
                   ("font-weight", "700"),
                   ("font-size", "16px"),
                   ("border", f"1px solid {header_bg}"),
                   ("text-align", "left")]},
        {"selector": "tbody td",
         "props": [("font-size", "15px"),
                   ("color", "#000000")]}
    ]).set_properties(**{"white-space": "nowrap"})
    return styler

def padronizar_tipo(df_in: pd.DataFrame) -> pd.DataFrame:
    df_out = df_in.copy()
    df_out["Tipo"] = (
        df_out["Tipo"].astype(str).str.strip()
        .replace({"asfalto":"Asfalto","terra":"Terra","palmeira":"Palmeira"})
    )
    return df_out

def meses_do_ano_zero():
    return pd.DataFrame({"Mes": MESES_ORDEM, "Rotulo": MESES_ROTULO, "Total": [0]*12})

def tipos_zero():
    return pd.DataFrame({"Tipo": TIPOS_ORDEM, "Total": [0]*len(TIPOS_ORDEM)})

if not os.path.exists(ARQUIVO_EXCEL):
    st.error(f"Arquivo não encontrado: {ARQUIVO_EXCEL}")
    st.info("Coloque o arquivo na mesma pasta do app ou ajuste ARQUIVO_EXCEL com o caminho completo.")
    st.stop()

abas_orig, abas_map = listar_abas_norm()

st.sidebar.title("Painel de Controle")
tela = st.sidebar.radio(
    "Menu",
    ["Visão Geral", "Motoristas (Top 10)", "Tipos + Semanas"],
    index=0
)

ano = st.sidebar.selectbox("Ano", [2025, 2026], index=0)

candidatas = [f"Exc_{ano}", f"Exc_Velocidade_{ano}", f"Exc_Velora_{ano}"]
aba_real = None
for cand in candidatas:
    if cand in abas_map:
        aba_real = abas_map[cand]
        break

if not aba_real:
    st.error("Não encontrei a aba do ano selecionado.")
    st.write("Abas encontradas no arquivo:", abas_orig)
    st.stop()

df = carregar_aba(aba_real)

mapa_gestor, gestores_oficiais = carregar_parametros_gersup()

df["Gestor"] = (
    df["Gestor"].astype(str).str.strip()
    .map(mapa_gestor)
    .fillna(df["Gestor"].astype(str).str.strip())
)

gestores = ["Todos"] + gestores_oficiais
gestor_sel = st.sidebar.selectbox("Gestor", gestores, index=0)

df_base = df.copy()
if gestor_sel != "Todos":
    df_base = df_base[df_base["Gestor"] == gestor_sel].copy()

df_ano = df_base[df_base["Ano"] == ano].copy()

if tela != "Visão Geral":
    meses_disp = [m for m in MESES_ORDEM if m in set(df_ano["Mes"].unique())]
    if not meses_disp:
        meses_disp = MESES_ORDEM[:]  # mantém a lista para permitir ver 0
    mes = st.sidebar.radio("Mês", meses_disp, index=min(len(meses_disp)-1, 0))
    df_mes = df_ano[df_ano["Mes"] == mes].copy()
else:
    mes = None
    df_mes = pd.DataFrame()

st.markdown(f"<div style='font-size:34px; font-weight:700; color:#000'>{(mes + '/' if mes else '')}{ano}</div>", unsafe_allow_html=True)
if gestor_sel != "Todos":
    st.markdown(f"<div style='font-size:18px; font-weight:700; color:#000'>Gestor: {gestor_sel}</div>", unsafe_allow_html=True)

if tela == "Visão Geral":
    st.markdown("<div style='font-size:26px; font-weight:700; color:#000; margin-top:10px;'>Visão Geral – Total anual por gestor</div>", unsafe_allow_html=True)

    total_ano = int(df_ano["Quantidade"].sum())

    col1, col2 = st.columns([2, 3])
    with col1:
        st.markdown("<div style='font-size:18px; font-weight:700; color:#000; margin-top:10px;'>Total anual</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:60px; font-weight:800; color:#000; line-height:1.0;'>{total_ano}</div>", unsafe_allow_html=True)

    with col2:
        df_tipo = padronizar_tipo(df_ano)
        tot_tipo_ano = (
            df_tipo.groupby("Tipo", as_index=False)["Quantidade"].sum()
            .rename(columns={"Quantidade": "Total"})
        )
        base_tipo = tipos_zero()
        tot_tipo_ano = (
            base_tipo.merge(tot_tipo_ano, on="Tipo", how="left", suffixes=("", "_y"))
        )
        tot_tipo_ano["Total"] = tot_tipo_ano["Total_y"].fillna(tot_tipo_ano["Total"]).astype(int)
        tot_tipo_ano = tot_tipo_ano[["Tipo", "Total"]]

        st.markdown("<div style='font-size:18px; font-weight:700; color:#000; margin-top:10px;'>Total por tipo (ano)</div>", unsafe_allow_html=True)

        chart_tipo = (
            alt.Chart(tot_tipo_ano)
            .mark_bar()
            .encode(
                x=alt.X("Tipo:N", sort=TIPOS_ORDEM, title="", axis=alt.Axis(labelAngle=0, labelFontSize=14, labelFontWeight="bold")),
                y=alt.Y("Total:Q", title="Total"),
                color=alt.Color(
                    "Tipo:N",
                    sort=TIPOS_ORDEM,
                    scale=alt.Scale(domain=TIPOS_ORDEM, range=[CORES_TIPO[t] for t in TIPOS_ORDEM]),
                    legend=alt.Legend(title="Tipo")
                ),
                tooltip=["Tipo", "Total"]
            )
        )
        text_tipo = (
            alt.Chart(tot_tipo_ano)
            .mark_text(dy=-8, fontSize=14, fontWeight="bold", color="black")
            .encode(
                x=alt.X("Tipo:N", sort=TIPOS_ORDEM),
                y="Total:Q",
                text="Total:Q"
            )
        )
        st.altair_chart(chart_tipo + text_tipo, use_container_width=True)

    st.markdown("<div style='font-size:18px; font-weight:700; color:#000; margin-top:14px;'>Total por mês (ano)</div>", unsafe_allow_html=True)

    tot_mes_ano = (
        df_ano.groupby("Mes", as_index=False)["Quantidade"].sum()
        .rename(columns={"Quantidade": "Total"})
    )

    base_mes = meses_do_ano_zero()
    tot_mes_ano = base_mes.merge(tot_mes_ano, left_on="Mes", right_on="Mes", how="left", suffixes=("", "_y"))
    tot_mes_ano["Total"] = tot_mes_ano["Total_y"].fillna(tot_mes_ano["Total"]).astype(int)
    tot_mes_ano = tot_mes_ano[["Mes", "Rotulo", "Total"]]

    bar_mes = (
        alt.Chart(tot_mes_ano)
        .mark_bar()
        .encode(
            x=alt.X("Rotulo:N", sort=MESES_ROTULO, title="", axis=alt.Axis(labelAngle=0, labelFontSize=14, labelFontWeight="bold")),
            y=alt.Y("Total:Q", title="Total"),
            tooltip=["Mes", "Total"]
        )
        .properties(height=290)
    )
    text_mes = (
        alt.Chart(tot_mes_ano)
        .mark_text(dy=-8, fontSize=14, fontWeight="bold", color="black")
        .encode(
            x=alt.X("Rotulo:N", sort=MESES_ROTULO),
            y="Total:Q",
            text="Total:Q"
        )
    )
    st.altair_chart(bar_mes + text_mes, use_container_width=True)

    st.markdown("<div style='font-size:18px; font-weight:700; color:#000; margin-top:14px;'>Top 10 motoristas (ano)</div>", unsafe_allow_html=True)

    ranking_ano = (
        df_ano.groupby("Motorista", as_index=False)["Quantidade"].sum()
        .rename(columns={"Quantidade": "Total"})
        .sort_values("Total", ascending=False)
        .head(10)
        .reset_index(drop=True)
    )

    if ranking_ano.empty:
        st.dataframe(aplicar_estilo_tabela(pd.DataFrame({"Motorista": [], "Total": []})), use_container_width=True, hide_index=True)
    else:
        st.dataframe(aplicar_estilo_tabela(ranking_ano), use_container_width=True, hide_index=True)

elif tela == "Motoristas (Top 10)":
    total_mes = int(df_mes["Quantidade"].sum())
    st.markdown("<div style='font-size:22px; font-weight:700; color:#000; margin-top:10px;'>Total Mês</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:52px; font-weight:800; color:#000; line-height:1.0;'>{total_mes}</div>", unsafe_allow_html=True)

    if df_mes.empty:
        st.warning("Sem dados para este gestor/mês/ano.")
        st.stop()

    ranking = (
        df_mes.groupby("Motorista", as_index=False)["Quantidade"]
        .sum()
        .rename(columns={"Quantidade": "Total"})
        .sort_values("Total", ascending=False)
        .head(10)
        .reset_index(drop=True)
    )

    st.markdown("<div style='font-size:22px; font-weight:700; color:#000; margin-top:18px;'>Top 10 motoristas (mês)</div>", unsafe_allow_html=True)
    st.dataframe(aplicar_estilo_tabela(ranking), use_container_width=True, hide_index=True)

else:
    total_mes = int(df_mes["Quantidade"].sum())
    st.markdown("<div style='font-size:22px; font-weight:700; color:#000; margin-top:10px;'>Total Mês</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:52px; font-weight:800; color:#000; line-height:1.0;'>{total_mes}</div>", unsafe_allow_html=True)

    if df_mes.empty:
        st.warning("Sem dados para este gestor/mês/ano.")
        st.stop()

    df_mes = padronizar_tipo(df_mes)

    tot_sem = (
        df_mes.groupby("Semana", as_index=False)["Quantidade"].sum()
        .rename(columns={"Quantidade": "Total"})
    )
    tot_sem = (
        tot_sem.set_index("Semana")
        .reindex(SEMANAS_ORDEM, fill_value=0)
        .reset_index()
    )

    st.markdown("<div style='font-size:22px; font-weight:700; color:#000; margin-top:18px;'>Total por semana (mês)</div>", unsafe_allow_html=True)

    bar_sem = (
        alt.Chart(tot_sem)
        .mark_bar()
        .encode(
            x=alt.X("Semana:N", sort=SEMANAS_ORDEM, title="", axis=alt.Axis(labelAngle=0, labelFontSize=14, labelFontWeight="bold")),
            y=alt.Y("Total:Q", title="Total"),
            tooltip=["Semana", "Total"]
        )
    )
    text_sem = (
        alt.Chart(tot_sem)
        .mark_text(dy=-8, fontSize=14, fontWeight="bold", color="black")
        .encode(
            x=alt.X("Semana:N", sort=SEMANAS_ORDEM),
            y="Total:Q",
            text="Total:Q"
        )
    )
    st.altair_chart(bar_sem + text_sem, use_container_width=True)

    tot_tipo = (
        df_mes.groupby("Tipo", as_index=False)["Quantidade"].sum()
        .rename(columns={"Quantidade": "Total"})
    )
    tot_tipo = (
        tot_tipo.set_index("Tipo")
        .reindex(TIPOS_ORDEM, fill_value=0)
        .reset_index()
    )

    st.markdown("<div style='font-size:22px; font-weight:700; color:#000; margin-top:18px;'>Total por tipo (mês)</div>", unsafe_allow_html=True)

    chart_tipo = (
        alt.Chart(tot_tipo)
        .mark_bar()
        .encode(
            x=alt.X("Tipo:N", sort=TIPOS_ORDEM, title="", axis=alt.Axis(labelAngle=0, labelFontSize=14, labelFontWeight="bold")),
            y=alt.Y("Total:Q", title="Total"),
            color=alt.Color(
                "Tipo:N",
                sort=TIPOS_ORDEM,
                scale=alt.Scale(domain=TIPOS_ORDEM, range=[CORES_TIPO[t] for t in TIPOS_ORDEM]),
                legend=alt.Legend(title="Tipo")
            ),
            tooltip=["Tipo", "Total"]
        )
    )
    text_tipo = (
        alt.Chart(tot_tipo)
        .mark_text(dy=-8, fontSize=14, fontWeight="bold", color="black")
        .encode(
            x=alt.X("Tipo:N", sort=TIPOS_ORDEM),
            y="Total:Q",
            text="Total:Q"
        )
    )
    st.altair_chart(chart_tipo + text_tipo, use_container_width=True)

    ranking_mes = (
        df_mes.groupby("Motorista", as_index=False)["Quantidade"]
        .sum()
        .rename(columns={"Quantidade": "Total"})
        .sort_values("Total", ascending=False)
        .head(10)
        .reset_index(drop=True)
    )

    st.markdown("<div style='font-size:22px; font-weight:700; color:#000; margin-top:18px;'>Top 10 motoristas (mês)</div>", unsafe_allow_html=True)
    st.dataframe(aplicar_estilo_tabela(ranking_mes), use_container_width=True, hide_index=True)
