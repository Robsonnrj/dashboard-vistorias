# -*- coding: utf-8 -*-
import pandas as pd
import plotly.express as px
import streamlit as st
from core.data_loader import read_df


# ------------------------------- helpers UI ------------------------------- #
def _kpi_block(label: str, value: str, sub: str):
    st.markdown(
        f"""
        <div style="border:1px solid #e5e7eb;border-radius:12px;padding:12px 16px;background:#fff">
          <div style="color:#6b7280;font-size:.85rem">{label}</div>
          <div style="font-size:1.8rem;font-weight:800">{value}</div>
          <div style="color:#6b7280;font-size:.8rem">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _optional_col(df: pd.DataFrame, candidates):
    """Tenta achar uma coluna por nomes equivalentes/parecidos."""
    if df is None or df.empty:
        return None
    # busca exata
    for c in candidates:
        if c in df.columns:
            return c
    # busca por "contém"
    up = [x.upper() for x in candidates]
    for c in df.columns:
        cc = c.upper()
        if any(u in cc for u in up):
            return c
    return None


# --------------------------------- page ---------------------------------- #
def page():
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    # lê a aba base escolhida na sidebar (mapeada no app principal)
    tab_base = st.session_state["tabs_map"]["solicitacoes"]
    df = read_df(tab_base)

    if df.empty:
        st.info("Sem dados ainda.")
        return

    # remove colunas duplicadas se houver (segurança extra)
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()]

    # mapeamento "tolerante" de colunas
    c_data = _optional_col(df, ["data_limite", "DATA DA SOLICITACAO", "DATA"])
    c_sit = _optional_col(df, ["status_atual", "Situação", "Situacao", "STATUS"])
    c_dir = _optional_col(df, ["diretoria", "Diretoria Responsável", "Diretoria"])
    c_om = _optional_col(df, ["om_solicitante", "OM APOIADA", "OM"])

    # ------------------------------ filtros ------------------------------ #
    colF1, colF2 = st.columns(2)
    with colF1:
        dir_opts = sorted(df[c_dir].dropna().astype(str).unique().tolist()) if c_dir else []
        dir_sel = st.multiselect("Diretoria", dir_opts)
    with colF2:
        sit_opts = sorted(df[c_sit].dropna().astype(str).unique().tolist()) if c_sit else []
        sit_sel = st.multiselect("Status", sit_opts)

    dff = df.copy()
    if dir_sel and c_dir:
        dff = dff[dff[c_dir].astype(str).isin(dir_sel)]
    if sit_sel and c_sit:
        dff = dff[dff[c_sit].astype(str).isin(sit_sel)]

    # ------------------------------- KPIs -------------------------------- #
    colK1, colK2, colK3, colK4 = st.columns(4)
    total = len(dff)
    pend = (dff[c_sit].astype(str).str.upper() == "SOLICITADA").sum() if c_sit else 0
    agend = (dff[c_sit].astype(str).str.upper() == "AGENDADA").sum() if c_sit else 0
    final = (dff[c_sit].astype(str).str.upper() == "FINALIZADA").sum() if c_sit else 0
    with colK1:
        _kpi_block("Solicitações", f"{total:,}".replace(",", "."), "Total")
    with colK2:
        _kpi_block("Pendentes", f"{pend:,}".replace(",", "."), "Status SOLICITADA")
    with colK3:
        _kpi_block("Agendadas", f"{agend:,}".replace(",", "."), "Status AGENDADA")
    with colK4:
        _kpi_block("Finalizadas", f"{final:,}".replace(",", "."), "Status FINALIZADA")

    st.divider()

    # ------------------------------ gráficos ---------------------------- #
    cols = st.columns(2)

    if c_dir and c_dir in dff.columns:
        with cols[0]:
            tmp = dff.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False)
            fig = px.bar(tmp, x=c_dir, y="size", title="Vistorias por Diretoria")
            st.plotly_chart(fig, use_container_width=True)

    if c_sit and c_sit in dff.columns:
        with cols[1]:
            tmp = dff.groupby(c_sit, as_index=False).size()
            fig = px.pie(tmp, names=c_sit, values="size", hole=0.45, title="Distribuição por Status")
            st.plotly_chart(fig, use_container_width=True)

    # --------------------------- últimos registros ----------------------- #
    st.subheader("Últimos registros")

    # segurança extra: se por algum motivo veio DataFrame ao fatiar, pegue a 1ª coluna
    if c_data and c_data in dff.columns:
        # re-garantir unicidade de cabeçalhos
        if dff.columns.duplicated().any():
            dff = dff.loc[:, ~dff.columns.duplicated()]

        series = dff[c_data]
        if isinstance(series, pd.DataFrame):
            series = series.iloc[:, 0]
        dff[c_data] = pd.to_datetime(series, errors="coerce")
        dff_show = dff.sort_values(c_data, ascending=False).head(50)
    else:
        dff_show = dff.head(50)

    st.dataframe(dff_show, use_container_width=True, height=360)
