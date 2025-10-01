# features/dashboard.py
# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px

from core.data_loader import read_df
from core.config import TAB_SOLICITACOES, TAB_VALIDACAO

def _clean(x): return "" if pd.isna(x) else str(x).strip()

def _load_oms_from_sources():
    """Lê OMs e Diretorias de Validacao_de_Dados (preferência) ou da base."""
    for tab in (TAB_VALIDACAO, TAB_SOLICITACOES):
        try:
            df = read_df(tab)
        except Exception:
            continue
        if df.empty:
            continue
        cols = {c.lower(): c for c in df.columns}
        c_sig = cols.get("om") or cols.get("om apoiada") or cols.get("sigla")
        c_dir = cols.get("diretoria responsável") or cols.get("diretoria")
        if not c_sig or not c_dir:
            continue
        tmp = df[[c_sig, c_dir]].copy()
        tmp.columns = ["OM", "Diretoria"]
        tmp["OM"] = tmp["OM"].map(_clean)
        tmp["Diretoria"] = tmp["Diretoria"].map(_clean)
        tmp = tmp[(tmp["OM"]!="") & (tmp["Diretoria"]!="")].drop_duplicates()
        return tmp
    return pd.DataFrame(columns=["OM","Diretoria"])

def page():
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    try:
        df = read_df(TAB_SOLICITACOES)
    except Exception as e:
        st.error(f"Falha ao ler a base: {e}")
        return
    if df.empty:
        st.info("Sem dados na aba ACOMPANHAMENTO VISTORIAS.")
        return

    # Colunas usuais (mapeamento tolerante)
    cols = {c.lower(): c for c in df.columns}
    c_obj = cols.get("objeto de vistoria")
    c_om  = cols.get("om apoiada") or cols.get("om")
    c_dir = cols.get("diretoria responsável") or cols.get("diretoria")
    c_sit = cols.get("situação") or cols.get("situacao")
    c_urg = cols.get("classificação de urgência") or cols.get("classificacao de urgencia")
    c_dtS = cols.get("data da solicitação") or cols.get("data da solicitacao")

    # Filtros hierárquicos
    st.sidebar.subheader("Filtros")
    ref_oms = _load_oms_from_sources()
    dir_opts = sorted(ref_oms["Diretoria"].dropna().unique().tolist()) if not ref_oms.empty else \
               sorted(df[c_dir].dropna().astype(str).unique().tolist()) if c_dir else []
    dir_sel = st.sidebar.multiselect("Diretoria", dir_opts)

    om_opts = []
    if dir_sel and not ref_oms.empty:
        om_opts = sorted(ref_oms[ref_oms["Diretoria"].isin(dir_sel)]["OM"].unique().tolist())
    elif c_om:
        om_opts = sorted(df[c_om].dropna().astype(str).unique().tolist())
    om_sel = st.sidebar.multiselect("OM Apoiada", om_opts)

    # Período
    if c_dtS and df[c_dtS].notna().any():
        try:
            base_dt = pd.to_datetime(df[c_dtS], errors="coerce")
            min_dt, max_dt = base_dt.min().date(), base_dt.max().date()
            per = st.sidebar.date_input("Período (pela data da solicitação)", value=(min_dt, max_dt))
        except Exception:
            per = None
    else:
        per = None

    # Aplica filtros
    dff = df.copy()
    if dir_sel and c_dir: dff = dff[dff[c_dir].astype(str).isin(dir_sel)]
    if om_sel and c_om:  dff = dff[dff[c_om].astype(str).isin(om_sel)]
    if per and c_dtS:
        ini, fim = per
        base_dt = pd.to_datetime(dff[c_dtS], errors="coerce")
        dff = dff[(base_dt >= pd.to_datetime(ini)) & (base_dt <= pd.to_datetime(fim))]

    # KPIs simples
    st.caption(f"Registros filtrados: {len(dff):,}")
    col1, col2, col3 = st.columns(3)
    with col1: st.metric("Total", f"{len(dff):,}")
    if c_sit:
        pend = (dff[c_sit].astype(str).str.upper().eq("NÃO ATENDIDA") | 
                dff[c_sit].astype(str).str.upper().eq("NAO ATENDIDA")).sum()
        with col2: st.metric("Não atendidas", f"{pend:,}")
    if c_urg:
        urgentes = dff[c_urg].astype(str).str.upper().eq("URGENTE").sum()
        with col3: st.metric("Urgentes", f"{urgentes:,}")

    # Gráficos
    gcols = st.columns(2)
    if c_dir:
        tmp = dff.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False)
        fig = px.bar(tmp, x=c_dir, y="size", title="Vistorias por Diretoria")
        gcols[0].plotly_chart(fig, use_container_width=True)

    if c_sit:
        tmp = dff.groupby(c_sit, as_index=False).size()
        fig = px.pie(tmp, names=c_sit, values="size", hole=.45, title="Distribuição por Situação")
        gcols[1].plotly_chart(fig, use_container_width=True)

    st.subheader("Últimos registros")
    if c_dtS in dff.columns:
        dff["_dt"] = pd.to_datetime(dff[c_dtS], errors="coerce")
        dff = dff.sort_values("_dt", ascending=False).drop(columns=["_dt"])
    st.dataframe(
        dff[[x for x in [c_obj, c_om, c_dir, c_sit, c_urg, c_dtS] if x in dff.columns]].head(50),
        use_container_width=True, height=380
    )
