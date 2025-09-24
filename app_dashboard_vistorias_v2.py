# app.py
# -*- coding: utf-8 -*-
# CRO1 — Editor + Dashboards (Google Sheets)

import warnings
warnings.filterwarnings(
    "ignore",
    message=".*outside the limits for dates.*",
    category=UserWarning,
    module="openpyxl",
)
warnings.filterwarnings(
    "ignore",
    message=".*Data Validation extension is not supported and will be removed.*",
    category=UserWarning,
    module="openpyxl",
)

import unicodedata
from datetime import datetime
from pathlib import Path

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials
from streamlit_option_menu import option_menu

# =========================================================
# CONFIG GERAL
# =========================================================
st.set_page_config(page_title="CRO1 — Editor & Dashboards (Sheets)", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# =========================================================
# CONEXÃO GOOGLE SHEETS
# =========================================================
def has_gsheets() -> bool:
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and "spreadsheet_url" in st.secrets["gsheets"]
        and bool(st.secrets["gsheets"]["spreadsheet_url"])
    )

@st.cache_resource(show_spinner=False)
def _gs_client():
    """Cliente gspread autenticado via service account do secrets.toml"""
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def _book():
    """Spreadsheet (arquivo) aberto pela URL do secrets.toml"""
    return _gs_client().open_by_url(st.secrets["gsheets"]["spreadsheet_url"])

def _ensure_ws_with_header(sheet, title: str, header: list[str]):

    # ===== Leituras/gravações tolerantes a cabeçalho “bagunçado” =====


def _make_unique_headers(raw_headers):
    """Gera nomes únicos: vazio -> col_1; duplicados -> nome_2, nome_3, ..."""
    out, seen = [], {}
    for j, h in enumerate(raw_headers, start=1):
        h = (h or "").strip()
        if not h:
            h = f"col_{j}"
        base = h
        if base in seen:
            seen[base] += 1
            h = f"{base}_{seen[base]}"
        else:
            seen[base] = 1
        out.append(h)
    return out

def read_ws_loose(ws, header_row=None) -> pd.DataFrame:
    """
    Lê a worksheet tolerando cabeçalho repetido/mesclado/vazio.
    - Se header_row não for dado, usa a primeira linha com algum conteúdo.
    - Garante nomes únicos nas colunas.
    """
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    # descobre a linha do cabeçalho
    if header_row is None:
        hdr_idx = next(
            (i for i, row in enumerate(values) if any(str(c).strip() for c in row)),
            0
        )
    else:
        hdr_idx = max(0, int(header_row) - 1)

    headers = _make_unique_headers(values[hdr_idx])
    body = values[hdr_idx + 1 :]

    # remove linhas finais 100% vazias (opcional)
    while body and not any(str(c).strip() for c in body[-1]):
        body.pop()

    df = pd.DataFrame(body, columns=headers).replace("", pd.NA)
    return df

def write_ws_over(ws, df: pd.DataFrame):
    """Sobrescreve a aba a partir de A1 com o DataFrame mostrado na tela."""
    ws.clear()
    headers = list(df.columns)
    rows = df.fillna("").astype(str).values.tolist()
    ws.update("A1", [headers] + rows, value_input_option="USER_ENTERED")
    """Garante que a worksheet exista e tenha o cabeçalho informado."""
    try:
        ws = sheet.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sheet.add_worksheet(title=title, rows=2000, cols=max(10, len(header)))
        ws.update("1:1", [header])
        return ws
    head = ws.row_values(1)
    if head != header:
        ws.update("1:1", [header])
    return ws

@st.cache_data(ttl=60, show_spinner=False)
def read_tab_df(tab_name: str) -> pd.DataFrame:
    """Lê uma aba do Sheets como DataFrame (infere header da linha 1)."""
    ws = _book().worksheet(tab_name)
    data = read_ws_loose(ws)
    df = pd.DataFrame(data)
    # normaliza datas
    for c in df.columns:
        if "DATA" in c.upper():
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

def overwrite_tab_from_df(tab_name: str, df: pd.DataFrame, keep_header=True):
    """
    Sobrescreve a aba com o DataFrame.
    - Se keep_header=True, usa df.columns como cabeçalho na linha 1.
    """
    sh = _book()
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=max(2000, len(df) + 10), cols=max(10, len(df.columns)))
    else:
        # limpa toda a aba
        ws.clear()

    if keep_header:
        values = [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist()
    else:
        values = df.fillna("").astype(str).values.tolist()

    # garante ter colunas/linhas suficientes
    ws.update("A1", values, value_input_option="USER_ENTERED")
    # invalida cache de leitura
    read_tab_df.clear()

# =========================================================
# HELPERS (normalização e mapeamento de colunas)
# =========================================================
def _norm(s: str) -> str:
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower().strip()

def col_or_none(df: pd.DataFrame, opts: list[str]) -> str | None:
    cols = list(df.columns)
    # exata
    for o in opts:
        for c in cols:
            if _norm(c) == _norm(o):
                return c
    # contém
    for o in opts:
        alvo = _norm(o)
        for c in cols:
            if alvo in _norm(c):
                return c
    return None

def card_title(txt: str):
    st.markdown(
        "<div style='padding:8px 12px;border-radius:10px;background:#f6f6f9;"
        "border:1px solid #e5e7eb;font-weight:700;font-size:20px;'>"
        f"{txt}</div>",
        unsafe_allow_html=True
    )

# =========================================================
# SIDEBAR (STATUS + MENU)
# =========================================================
with st.sidebar:
    st.write("🔌 Google Sheets:", "ON ✅" if has_gsheets() else "OFF ❌")
    if not has_gsheets():
        st.error("Secrets não detectado. Configure `.streamlit/secrets.toml`.")

    if has_gsheets() and st.button("🔄 Limpar cache e recarregar"):
        read_tab_df.clear()
        st.success("Cache limpo.")
        st.rerun()

    MENU = option_menu(
        "CRO1 — Sistema",
        ["🗂️ Editor da Planilha", "📊 Dashboards"],
        icons=["table", "bar-chart"],
        default_index=0,
        menu_icon="grid"
    )

# =========================================================
# 1) EDITOR DA PLANILHA (visualizar/editar/salvar)
# =========================================================
if MENU == "🗂️ Editor da Planilha":
    st.title("🗂️ Editor da Planilha (Google Sheets)")
    if not has_gsheets():
        st.stop()

    sh = _book()
    tabs = [ws.title for ws in sh.worksheets()]
    st.success("Google Sheets conectado ✅")
    st.caption(f"Planilha: {st.secrets['gsheets']['spreadsheet_url']}")

    # escolha de aba
    tab_name = st.selectbox("Escolha a aba para visualizar/editar:", tabs, index=0)
    btn_reload = st.button("↻ Recarregar aba selecionada")

    if btn_reload:
        read_tab_df.clear()

    # carrega DF da aba
    try:
        df_tab = read_tab_df(tab_name)
    except Exception as e:
        st.error(f"Falha ao ler a aba **{tab_name}**: {e}")
        st.stop()

    st.caption(f"Linhas: {len(df_tab)} • Colunas: {list(df_tab.columns)}")

    # Editor interativo
    edited_df = st.data_editor(
        df_tab,
        use_container_width=True,
        num_rows="dynamic",
        key=f"editor_{tab_name}",
        height=520,
    )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 Salvar alterações na aba"):
            try:
                # tentativa de converter colunas com "DATA" para ISO antes de salvar
                _df_out = edited_df.copy()
                for c in _df_out.columns:
                    if "DATA" in c.upper():
                        _df_out[c] = pd.to_datetime(_df_out[c], errors="coerce")
                        _df_out[c] = _df_out[c].dt.strftime("%Y-%m-%d")

                overwrite_tab_from_df(tab_name, _df_out, keep_header=True)
                st.success(f"Alterações salvas em **{tab_name}**.")
            except Exception as e:
                st.error(f"Falha ao salvar: {e}")

    with col2:
        csv = edited_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Baixar CSV desta aba", csv, file_name=f"{tab_name}.csv", mime="text/csv")

# =========================================================
# 2) DASHBOARDS (usa diretamente a aba escolhida)
# =========================================================
if MENU == "📊 Dashboards":
    st.title("📊 Resumos (Dashboards)")

    if not has_gsheets():
        st.warning("Ative o Google Sheets para carregar dashboards.")
        st.stop()

    sh = _book()
    tabs = [ws.title for ws in sh.worksheets()]

    # escolha da aba base para o dashboard (ex.: ACOMPANHAMENTO VISTORIAS)
    base_tab = st.selectbox(
        "Escolha a aba (fonte dos gráficos/KPIs):",
        tabs,
        index=0,
        key="dashboard_tab",
    )

    try:
        df = read_tab_df(base_tab)
    except Exception as e:
        st.error(f"Falha ao ler a aba **{base_tab}**: {e}")
        st.stop()

    if df.empty:
        st.warning("A aba está vazia.")
        st.stop()
def _make_unique_headers(raw_headers):
    """Gera nomes únicos: vazio -> col_1, duplicado -> nome_2, nome_3, ..."""
    out, seen = [], {}
    for j, h in enumerate(raw_headers, start=1):
        h = (h or "").strip()
        if not h:
            h = f"col_{j}"
        base = h
        if base in seen:
            seen[base] += 1
            h = f"{base}_{seen[base]}"
        else:
            seen[base] = 1
        out.append(h)
    return out

def read_ws_loose(ws, header_row=None) -> pd.DataFrame:
    """
    Lê a worksheet tolerando cabeçalho repetido/mesclado/vazio.
    - Se header_row não for dado, usa a primeira linha que tenha algum conteúdo.
    - Garante nomes únicos para as colunas.
    """
    values = ws.get_all_values()  # lista de listas
    if not values:
        return pd.DataFrame()

    # acha a linha do cabeçalho
    if header_row is None:
        hdr_idx = next(
            (i for i, row in enumerate(values) if any(str(c).strip() for c in row)),
            0
        )
    else:
        hdr_idx = max(0, int(header_row) - 1)

    headers = _make_unique_headers(values[hdr_idx])
    body = values[hdr_idx + 1 :]
    # corta linhas completamente vazias no fim (opcional)
    while body and not any(str(c).strip() for c in body[-1]):
        body.pop()

    df = pd.DataFrame(body, columns=headers)
    # troca strings vazias por NA (opcional)
    df = df.replace("", pd.NA)
    return df
    # --------- Mapeamento tolerante de colunas ---------
    c_obj = col_or_none(df, ["OBJETO DE VISTORIA", "OBJETO"])
    c_om  = col_or_none(df, ["OM APOIADA", "OM APOIADORA", "OM"])
    c_dir = col_or_none(df, ["Diretoria Responsavel", "Diretoria Responsável", "Diretoria"])
    c_urg = col_or_none(df, ["Classificacao da Urgencia", "Classificação da Urgência", "Urgencia"])
    c_sit = col_or_none(df, ["Situacao", "Situação"])
    c_data_solic = col_or_none(df, ["DATA DA SOLICITACAO", "DATA DA SOLICITAÇÃO"])
    c_data_vist  = col_or_none(df, ["DATA DA VISTORIA"])
    c_dias_total = col_or_none(df, ["QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO"])
    c_dias_exec  = col_or_none(df, ["QUANTIDADE DE DIAS PARA EXECUCAO", "QUANTIDADE DE DIAS PARA EXECUÇÃO"])
    c_status     = col_or_none(df, ["STATUS - ATUALIZACAO SEMANAL", "Status"])

    # --------- Tipos ---------
    for c in [c_data_solic, c_data_vist]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in [c_dias_total, c_dias_exec]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    st.caption(f"Aba: **{base_tab}** • Linhas: {len(df)}")

    # --------- Filtros ---------
    st.sidebar.subheader("Filtros (Dashboards)")
    col_data_base = c_data_solic if c_data_solic in df.columns else c_data_vist
    if col_data_base and df[col_data_base].notna().any():
        min_dt = pd.to_datetime(df[col_data_base].min()).date()
        max_dt = pd.to_datetime(df[col_data_base].max()).date()
        periodo = st.sidebar.date_input("Período (pela data da solicitação)", value=(min_dt, max_dt))
    else:
        periodo = None

    def opts(series):
        try:
            return sorted(series.dropna().astype(str).unique().tolist())
        except Exception:
            return sorted(list({str(x) for x in series.dropna().tolist()}))

    dir_sel = st.sidebar.multiselect("Diretoria Responsável", opts(df[c_dir]) if c_dir in df.columns else [])
    sit_sel = st.sidebar.multiselect("Situação", opts(df[c_sit]) if c_sit in df.columns else [])
    urg_sel = st.sidebar.multiselect("Classificação de Urgência", opts(df[c_urg]) if c_urg in df.columns else [])
    om_sel  = st.sidebar.multiselect("OM Apoiadora", opts(df[c_om]) if c_om in df.columns else [])
    sla_dias = st.sidebar.number_input("SLA (dias para 'dentro do prazo')", 1, 365, value=30)

    df_f = df.copy()
    if periodo and col_data_base:
        ini, fim = periodo
        df_f = df_f[(df_f[col_data_base] >= pd.to_datetime(ini)) & (df_f[col_data_base] <= pd.to_datetime(fim))]
    if dir_sel and c_dir in df.columns:
        df_f = df_f[df_f[c_dir].astype(str).isin(dir_sel)]
    if sit_sel and c_sit in df.columns:
        df_f = df_f[df_f[c_sit].astype(str).isin(sit_sel)]
    if urg_sel and c_urg in df.columns:
        df_f = df_f[df_f[c_urg].astype(str).isin(urg_sel)]
    if om_sel and c_om in df.columns:
        df_f = df_f[df_f[c_om].astype(str).isin(om_sel)]

    # --------- KPIs ---------
    colk1, colk2, colk3, colk4, colk5 = st.columns(5)
    total_vist = len(df_f)
    finalizadas = df_f[c_sit].astype(str).str.upper().str.contains("FINALIZADA").sum() if c_sit in df_f.columns else None
    pct_final = (finalizadas / total_vist * 100) if (finalizadas is not None and total_vist > 0) else 0
    prazo_medio_total = df_f[c_dias_total].mean() if c_dias_total in df_f.columns else None
    prazo_medio_exec  = df_f[c_dias_exec].mean() if c_dias_exec   in df_f.columns else None
    pct_sla = None
    if c_dias_total in df_f.columns and total_vist > 0:
        dentro_sla = (df_f[c_dias_total] <= sla_dias).sum()
        pct_sla = dentro_sla / total_vist * 100

    with colk1: st.metric("Total de Vistorias", f"{total_vist:,}".replace(",", "."))
    with colk2: st.metric("Finalizadas (%)", f"{pct_final:,.1f}%")
    with colk3: st.metric("Prazo médio total (dias)", f"{prazo_medio_total:,.1f}" if prazo_medio_total is not None else "—")
    with colk4: st.metric("Prazo médio execução (dias)", f"{prazo_medio_exec:,.1f}" if prazo_medio_exec is not None else "—")
    with colk5: st.metric(f"% dentro do SLA (≤{sla_dias}d)", f"{pct_sla:,.1f}%" if pct_sla is not None else "—")

    st.divider()

    # --------- Gráficos ---------
    if col_data_base and df_f[col_data_base].notna().any():
        tmp = (df_f.groupby(pd.Grouper(key=col_data_base, freq="MS"))
               .size().reset_index(name="Vistorias"))
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

    if c_dir in df_f.columns and c_dias_total in df_f.columns:
        base = df_f.dropna(subset=[c_dir, c_dias_total]).copy()
        base["Dentro SLA"] = base[c_dias_total] <= sla_dias
        tmp_sla = (base.groupby(c_dir)["Dentro SLA"].mean()*100).reset_index(name="pct_sla")
        fig_sla = px.bar(tmp_sla.sort_values("pct_sla"), x="pct_sla", y=c_dir, orientation="h",
                         title=f"% Dentro do SLA (≤{sla_dias}d) por Diretoria",
                         labels={"pct_sla": "% dentro do SLA"})
        st.plotly_chart(fig_sla, use_container_width=True)

    if col_data_base and c_sit in df_f.columns and df_f[col_data_base].notna().any():
        aux = df_f.copy()
        aux["Mes"] = aux[col_data_base].dt.to_period("M").dt.to_timestamp()
        piv = (aux.groupby(["Mes", c_sit]).size().reset_index(name="Qtd")
               .pivot(index="Mes", columns=c_sit, values="Qtd").fillna(0))
        fig_hm = px.imshow(piv.T, aspect="auto",
                           labels=dict(x="Mês", y="Situação", color="Qtd"),
                           title="Heatmap — Mês x Situação")
        st.plotly_chart(fig_hm, use_container_width=True)

    card_title("Detalhamento (mais recentes)")
    ord_col = col_data_base if col_data_base else (c_data_vist if c_data_vist in df_f.columns else None)
    df_show = df_f.sort_values(ord_col, ascending=False).head(80) if ord_col else df_f.head(80)
    st.dataframe(df_show, use_container_width=True)
