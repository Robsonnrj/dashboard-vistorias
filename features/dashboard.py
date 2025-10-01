# -*- coding: utf-8 -*-
import pandas as pd
import plotly.express as px
import streamlit as st
from core.data_loader import read_df
from core.config import TAB_SOLICITACOES

def _kpi(label: str, value: str, sub: str):
    st.markdown(
        f"<div style='border:1px solid #e5e7eb;border-radius:12px;padding:12px 16px;background:#fff'>"
        f"<div style='color:#6b7280;font-size:.85rem'>{label}</div>"
        f"<div style='font-size:1.8rem;font-weight:800'>{value}</div>"
        f"<div style='color:#6b7280;font-size:.8rem'>{sub}</div></div>",
        unsafe_allow_html=True
    )

def _opt(df, candidates):
    for c in candidates:
        if c in df.columns: return c
    up = [c.upper() for c in candidates]
    for c in df.columns:
        if any(k in c.upper() for k in up): return c
    return None

def page():
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    df = read_df(TAB_SOLICITACOES)
    if df.empty:
        st.info(f"Sem dados na aba **{TAB_SOLICITACOES}**.")
        return

    c_data = _opt(df, ["data_limite","data_solicitacao","DATA DA SOLICITACAO","DATA"])
    c_sit  = _opt(df, ["status_atual","Situação","Situacao","STATUS"])
    c_dir  = _opt(df, ["diretoria","Diretoria Responsável"])
    c_om   = _opt(df, ["om_solicitante","OM APOIADA","OM"])

    colF1, colF2 = st.columns(2)
    with colF1:
        dir_sel = st.multiselect("Diretoria",
                                 sorted(df[c_dir].dropna().unique().tolist()) if c_dir else [])
    with colF2:
        sit_sel = st.multiselect("Status",
                                 sorted(df[c_sit].dropna().unique().tolist()) if c_sit else [])

    dff = df.copy()
    if dir_sel and c_dir: dff = dff[dff[c_dir].isin(dir_sel)]
    if sit_sel and c_sit: dff = dff[dff[c_sit].isin(sit_sel)]

    colK1, colK2, colK3, colK4 = st.columns(4)
    total = len(dff)
    pend  = (dff[c_sit] == "SOLICITADA").sum() if c_sit else 0
    agend = (dff[c_sit] == "AGENDADA").sum() if c_sit else 0
    final = (dff[c_sit] == "FINALIZADA").sum() if c_sit else 0
    with colK1: _kpi("Solicitações", f"{total:,}".replace(",", "."), "Total")
    with colK2: _kpi("Pendentes", f"{pend:,}".replace(",", "."), "SOLICITADA")
    with colK3: _kpi("Agendadas", f"{agend:,}".replace(",", "."), "AGENDADA")
    with colK4: _kpi("Finalizadas", f"{final:,}".replace(",", "."), "FINALIZADA")

    st.divider()
    cols = st.columns(2)

    if c_dir:
        with cols[0]:
            tmp = dff.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False)
            st.plotly_chart(px.bar(tmp, x=c_dir, y="size", title="Vistorias por Diretoria"),
                            use_container_width=True)

    if c_sit:
        with cols[1]:
            tmp = dff.groupby(c_sit, as_index=False).size()
            st.plotly_chart(px.pie(tmp, names=c_sit, values="size", hole=.45,
                                   title="Distribuição por Status"), use_container_width=True)

    st.subheader("Últimos registros")
    if c_data and c_data in dff.columns:
        dff[c_data] = pd.to_datetime(dff[c_data], errors="coerce")
        dff = dff.sort_values(c_data, ascending=False)
    st.dataframe(dff.head(50), use_container_width=True, height=360)
