# -*- coding: utf-8 -*-
import pandas as pd
import plotly.express as px
import streamlit as st
from core.data_loader import read_df

def _kpi_block(label: str, value: str, sub: str):
    st.markdown(
        f"""
        <div style="border:1px solid #e5e7eb;border-radius:12px;padding:12px 16px;background:#fff">
          <div style="color:#6b7280;font-size:.85rem">{label}</div>
          <div style="font-size:1.8rem;font-weight:800">{value}</div>
          <div style="color:#6b7280;font-size:.8rem">{sub}</div>
        </div>
        """, unsafe_allow_html=True
    )

def _optional_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    up = {c.upper(): c for c in df.columns}
    for wanted in candidates:
        for col_up, col in up.items():
            if wanted.upper() in col_up:
                return col
    return None

def page():
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    # aba base escolhida na sidebar (mapeada em st.session_state['tabs_map'])
    tab_base = st.session_state["tabs_map"]["solicitacoes"]
    df = read_df(tab_base)      # <<< AQUI estava o bug: use df, não df_oms
    if df.empty:
        st.info("Sem dados ainda.")
        return

    c_data = _optional_col(df, ["data_limite","DATA DA SOLICITACAO","DATA"])
    c_sit  = _optional_col(df, ["status_atual","Situação","Situacao","STATUS"])
    c_dir  = _optional_col(df, ["diretoria","Diretoria Responsavel","Diretoria"])
    c_om   = _optional_col(df, ["om_solicitante","OM APOIADA","OM"])

    # Filtros
    colF1, colF2 = st.columns(2)
    with colF1:
        dir_sel = st.multiselect("Diretoria", sorted(df[c_dir].dropna().unique()) if c_dir else [])
    with colF2:
        sit_sel = st.multiselect("Status", sorted(df[c_sit].dropna().unique()) if c_sit else [])

    dff = df.copy()
    if dir_sel and c_dir:
        dff = dff[dff[c_dir].isin(dir_sel)]
    if sit_sel and c_sit:
        dff = dff[dff[c_sit].isin(sit_sel)]

    # KPIs
    colK1, colK2, colK3, colK4 = st.columns(4)
    total = len(dff)
    pend  = (dff[c_sit].astype(str).str.upper() == "SOLICITADA").sum() if c_sit else 0
    agend = (dff[c_sit].astype(str).str.upper() == "AGENDADA").sum() if c_sit else 0
    final = (dff[c_sit].astype(str).str.upper() == "FINALIZADA").sum() if c_sit else 0
    with colK1: _kpi_block("Solicitações", f"{total:,}".replace(",", "."), "Total")
    with colK2: _kpi_block("Pendentes", f"{pend:,}".replace(",", "."), "Status SOLICITADA")
    with colK3: _kpi_block("Agendadas", f"{agend:,}".replace(",", "."), "Status AGENDADA")
    with colK4: _kpi_block("Finalizadas", f"{final:,}".replace(",", "."), "Status FINALIZADA")

    st.divider()
    cols = st.columns(2)

    if c_dir and c_dir in dff.columns:
        with cols[0]:
            tmp = dff.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False)
            st.plotly_chart(px.bar(tmp, x=c_dir, y="size", title="Vistorias por Diretoria"),
                            use_container_width=True)

    if c_sit and c_sit in dff.columns:
        with cols[1]:
            tmp = dff.groupby(c_sit, as_index=False).size()
            st.plotly_chart(px.pie(tmp, names=c_sit, values="size", hole=.45,
                                   title="Distribuição por Status"),
                            use_container_width=True)

    st.subheader("Últimos registros")
    if c_data and c_data in dff.columns:
    dff[c_data] = pd.to_datetime(dff[c_data], errors="coerce")
    dff = dff.sort_values(c_data, ascending=False)
    st.dataframe(dff.head(50), use_container_width=True, height=360)
