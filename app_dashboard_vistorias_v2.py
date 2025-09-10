
import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata
from datetime import datetime

st.set_page_config(page_title="Dashboard Vistorias CRO1", layout="wide")

st.title("📋 Dashboard de Acompanhamento de Vistorias (CRO1) — v2")
st.caption("Correção aplicada para colunas com tipos mistos (números + textos).")

# ---------- Upload / Load ----------
arquivo = st.file_uploader("Envie o Excel (.xlsx). Se não enviar, o app tenta ler o arquivo 'Acomp. de Vistorias CRO1 - 2025.xlsx' da pasta.", type=["xlsx"])

@st.cache_data
def carregar_excel(file_like):
    if file_like is None:
        try:
            xls = pd.ExcelFile("Acomp. de Vistorias CRO1 - 2025.xlsx")
        except Exception as e:
            st.error("Não encontrei 'Acomp. de Vistorias CRO1 - 2025.xlsx' na pasta. Envie sua planilha acima.")
            st.stop()
    else:
        xls = pd.ExcelFile(file_like)

    preferidas = ["ACOMPANHAMENTO VISTORIAS", "Acompanhamento Vistorias"]
    alvo = None
    for sh in xls.sheet_names:
        if sh in preferidas:
            alvo = sh
            break
    if alvo is None:
        alvo = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=alvo)
    return df, alvo, xls.sheet_names

def norm(s):
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower().strip()

df_raw, aba, todas_abas = carregar_excel(arquivo)
st.write(f"**Aba carregada:** `{aba}`  •  Abas no arquivo: {todas_abas}")

# ---------- Mapear colunas ----------
def achar(col_alvo, candidatos):
    alvo_n = norm(col_alvo)
    for c in candidatos:
        if norm(c) == alvo_n:
            return c
    for c in candidatos:
        if alvo_n in norm(c):
            return c
    return None

cols = list(df_raw.columns)

c_obj = achar("OBJETO DE VISTORIA", cols) or achar("OBJETO", cols)
c_om = achar("OM APOIADA", cols)
c_dir = achar("Diretoria Responsavel", cols)
c_urg = achar("Classificacao de Urgencia", cols)
c_sit = achar("Situacao", cols)
c_data_solic = achar("DATA DA SOLICITACAO", cols)
c_data_vist = achar("DATA DA VISTORIA", cols)
c_dias_total = achar("QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO", cols)
c_dias_exec = achar("QUANTIDADE DE DIAS PARA EXECUCAO", cols)
c_status = achar("STATUS - ATUALIZACAO SEMANAL", cols)

df = df_raw.copy()

# Conversões de tipos
for c in [c_data_solic, c_data_vist]:
    if c in df.columns:
        df[c] = pd.to_datetime(df[c], errors="coerce")
for c in [c_dias_total, c_dias_exec]:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="coerce")

# Helper: opções seguras para multiselect (tudo como string e ordenado)
def opts(series):
    try:
        return sorted(series.dropna().astype(str).unique().tolist())
    except Exception:
        return sorted(list({str(x) for x in series.dropna().tolist()}))

# ---------- Filtros ----------
st.sidebar.header("Filtros")
col_data_base = c_data_solic if c_data_solic in df.columns else c_data_vist
if col_data_base and df[col_data_base].notna().any():
    min_dt = pd.to_datetime(df[col_data_base].min())
    max_dt = pd.to_datetime(df[col_data_base].max())
    periodo = st.sidebar.date_input("Período (pela data da solicitação)", value=(min_dt.date(), max_dt.date()))
else:
    periodo = None

dir_sel = st.sidebar.multiselect("Diretoria Responsável", opts(df[c_dir]) if c_dir in df.columns else [])
sit_sel = st.sidebar.multiselect("Situação", opts(df[c_sit]) if c_sit in df.columns else [])
urg_sel = st.sidebar.multiselect("Classificação de Urgência", opts(df[c_urg]) if c_urg in df.columns else [])
om_sel = st.sidebar.multiselect("OM Apoiadora", opts(df[c_om]) if c_om in df.columns else [])

df_f = df.copy()
if periodo and col_data_base:
    ini, fim = periodo if isinstance(periodo, (list, tuple)) else (periodo, periodo)
    df_f = df_f[(df_f[col_data_base] >= pd.to_datetime(ini)) & (df_f[col_data_base] <= pd.to_datetime(fim))]
if dir_sel and c_dir in df.columns:
    df_f = df_f[df_f[c_dir].astype(str).isin(dir_sel)]
if sit_sel and c_sit in df.columns:
    df_f = df_f[df_f[c_sit].astype(str).isin(sit_sel)]
if urg_sel and c_urg in df.columns:
    df_f = df_f[df_f[c_urg].astype(str).isin(urg_sel)]
if om_sel and c_om in df.columns:
    df_f = df_f[df_f[c_om].astype(str).isin(om_sel)]

# ---------- KPIs ----------
col1, col2, col3, col4 = st.columns(4)
total_vist = len(df_f)
finalizadas = None
if c_sit in df_f.columns:
    finalizadas = df_f[c_sit].astype(str).str.upper().str.contains("FINALIZADA").sum()
pct_final = (finalizadas / total_vist * 100) if (finalizadas is not None and total_vist > 0) else 0

prazo_medio = df_f[c_dias_total].mean() if c_dias_total in df_f.columns else None
exec_medio = df_f[c_dias_exec].mean() if c_dias_exec in df_f.columns else None

with col1:
    st.metric("Total de Vistorias", f"{total_vist:,}".replace(",", "."))
with col2:
    st.metric("Finalizadas (%)", f"{pct_final:,.1f}%")
with col3:
    st.metric("Prazo médio total (dias)", f"{prazo_medio:,.1f}" if prazo_medio is not None else "—")
with col4:
    st.metric("Prazo médio execução (dias)", f"{exec_medio:,.1f}" if exec_medio is not None else "—")

st.divider()

# ---------- Gráficos ----------
if col_data_base and df_f[col_data_base].notna().any():
    tmp = (df_f
           .groupby(pd.Grouper(key=col_data_base, freq="MS"))
           .size()
           .reset_index(name="Vistorias"))
    fig1 = px.line(tmp, x=col_data_base, y="Vistorias", markers=True, title="Evolução Mensal de Vistorias")
    st.plotly_chart(fig1, use_container_width=True)

if c_dir in df_f.columns:
    tmp2 = df_f.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False)
    fig2 = px.bar(tmp2, x=c_dir, y="size", title="Vistorias por Diretoria Responsável")
    st.plotly_chart(fig2, use_container_width=True)

if c_sit in df_f.columns:
    tmp3 = df_f.groupby(c_sit, as_index=False).size()
    fig3 = px.pie(tmp3, names=c_sit, values="size", hole=0.4, title="Distribuição por Situação")
    st.plotly_chart(fig3, use_container_width=True)

if c_urg in df_f.columns:
    tmp4 = df_f.groupby(c_urg, as_index=False).size().sort_values("size", ascending=False)
    fig4 = px.bar(tmp4, x=c_urg, y="size", title="Vistorias por Classificação de Urgência")
    st.plotly_chart(fig4, use_container_width=True)

st.subheader("Detalhamento (mais recentes)")
ord_col = col_data_base if col_data_base else (c_data_vist if c_data_vist in df_f.columns else None)
if ord_col:
    df_show = df_f.sort_values(ord_col, ascending=False).head(50)
else:
    df_show = df_f.head(50)
st.dataframe(df_show, use_container_width=True)
