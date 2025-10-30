# -*- coding: utf-8 -*-
"""
core/audit.py
Página de Auditoria com a mesma "ideia" das demais:
- filtros na sidebar
- KPIs no topo
- gráficos (linha e barras)
- tabela e download
"""

from __future__ import annotations
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

from core.config import TAB_AUDIT
from core.data_loader import read_df


def _nf(x) -> str:
    return "" if pd.isna(x) else str(x).strip()


def page():
    st.header("🔎 Registro de Auditoria — Trilhas de Alterações")

    # ================== carga ==================
    try:
        df_raw = read_df(TAB_AUDIT)
    except Exception as e:
        st.error(f"Não consegui ler a aba de auditoria: {e}")
        return

    if df_raw is None or df_raw.empty:
        st.info("Ainda não há registros de auditoria.")
        return

    df = df_raw.copy()

    # map de colunas (tolerante a variações)
    cols = {c.lower(): c for c in df.columns}
    c_num   = cols.get("numero") or cols.get("número")
    c_ts    = cols.get("ts") or cols.get("timestamp") or cols.get("data")
    c_campo = cols.get("campo") or cols.get("coluna")
    c_antes = cols.get("antes")
    c_depois= cols.get("depois")
    if not c_num or not c_ts or not c_campo:
        st.error("A planilha de auditoria precisa ter, no mínimo, as colunas: numero, ts, campo.")
        return

    # normalizações
    df[c_ts] = pd.to_datetime(df[c_ts], errors="coerce")
    for c in [c_num, c_campo, c_antes, c_depois]:
        if c and c in df.columns:
            df[c] = df[c].astype(str).map(_nf)

    # ================== filtros (sidebar) ==================
    with st.sidebar:
        st.subheader("Filtros — Auditoria")

        # período
        dt_min, dt_max = df[c_ts].min(), df[c_ts].max()
        ini, fim = st.date_input(
            "Período",
            (
                (dt_min.date() if pd.notna(dt_min) else pd.Timestamp.today().date()),
                (dt_max.date() if pd.notna(dt_max) else pd.Timestamp.today().date()),
            ),
        )

        # números (registros) e campos
        ops_num = sorted(df[c_num].dropna().unique())
        ops_cam = sorted(df[c_campo].dropna().unique())

        numeros = st.multiselect("Número", ops_num)
        campos  = st.multiselect("Campo alterado", ops_cam)

        termo = st.text_input("Contém (antes/depois)", placeholder="texto livre")

    # aplica filtros
    mask = pd.Series(True, index=df.index)
    mask &= df[c_ts].between(pd.to_datetime(ini), pd.to_datetime(fim) + pd.Timedelta(days=1))
    if numeros:
        mask &= df[c_num].isin(numeros)
    if campos:
        mask &= df[c_campo].isin(campos)
    if termo:
        termo_cf = termo.casefold()
        bloco = []
        for c in [c_antes, c_depois]:
            if c and c in df.columns:
                bloco.append(df[c].str.casefold().str.contains(termo_cf, na=False))
        mask &= np.column_stack(bloco).any(axis=1) if bloco else mask

    dff = df[mask].copy()

    # ================== KPIs ==================
    st.subheader("Indicadores")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Eventos de auditoria (filtro)", f"{len(dff)}")
    c2.metric("Registros distintos", f"{dff[c_num].nunique()}")
    c3.metric("Campos distintos", f"{dff[c_campo].nunique()}")
    ult7 = dff[dff[c_ts] >= (pd.Timestamp.now() - pd.Timedelta(days=7))]
    c4.metric("Eventos nos últimos 7 dias", f"{len(ult7)}")

    # ================== gráficos ==================
    st.subheader("Visualizações")

    # 1) Série temporal — eventos/dia
    ts = (
        dff.assign(dia=dff[c_ts].dt.floor("D"))
           .groupby("dia", as_index=False)
           .size()
           .sort_values("dia")
    )
    if not ts.empty:
        fig_ts = px.line(ts, x="dia", y="size", markers=True, title="Eventos por dia")
        fig_ts.update_traces(line_shape="spline")
        st.plotly_chart(fig_ts, use_container_width=True)

    # 2) Barras — Top campos alterados
    gr_campos = (
        dff.groupby(c_campo, as_index=False)
           .size()
           .sort_values("size", ascending=False)
           .head(15)
    )
    if not gr_campos.empty:
        fig_cam = px.bar(gr_campos, x="size", y=c_campo, orientation="h", title="Top campos alterados")
        st.plotly_chart(fig_cam, use_container_width=True)

    # 3) Barras — Top números (registros) mais alterados
    gr_nums = (
        dff.groupby(c_num, as_index=False)
           .size()
           .sort_values("size", ascending=False)
           .head(15)
    )
    if not gr_nums.empty:
        fig_num = px.bar(gr_nums, x="size", y=c_num, orientation="h", title="Registros com mais alterações")
        st.plotly_chart(fig_num, use_container_width=True)

    # ================== detalhe por número ==================
    with st.expander("🔍 Ver histórico detalhado por Número"):
        num_sel = st.selectbox("Escolha um número", [""] + ops_num)
        if num_sel:
            h = dff[dff[c_num] == num_sel].sort_values(c_ts, ascending=False)
            st.dataframe(h, use_container_width=True, hide_index=True)

    # ================== tabela + download ==================
    st.subheader("Tabela de Auditoria (filtrada)")
    st.dataframe(dff.sort_values(c_ts, ascending=False), use_container_width=True, hide_index=True)

    # botão de download (CSV)
    @st.cache_data(ttl=60)
    def _to_csv(df_: pd.DataFrame) -> bytes:
        return df_.to_csv(index=False).encode("utf-8")

    st.download_button(
        "⬇️ Baixar CSV (filtro)",
        data=_to_csv(dff),
        file_name="auditoria_filtrada.csv",
        mime="text/csv",
        use_container_width=True,
    )
