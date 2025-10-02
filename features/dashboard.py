# -*- coding: utf-8 -*-
import pandas as pd
import plotly.express as px
import streamlit as st
import altair as alt  # opcional (usado no stacked opcional)
from core.data_loader import read_df
from core.config import TAB_SOLICITACOES


def _pick(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Encontra coluna pela lista de candidatos (match exato ou contém, casefold)."""
    if df.empty:
        return None
    cols = list(df.columns)

    def nf(s: str) -> str: return s.casefold().strip()

    # exato
    for c in candidates:
        for cc in cols:
            if nf(cc) == nf(c):
                return cc
    # contém
    for c in candidates:
        target = nf(c)
        for cc in cols:
            if target in nf(cc):
                return cc
    return None


def page():
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    try:
        df = read_df(TAB_SOLICITACOES)
    except Exception as e:
        st.error(f"Não foi possível ler a aba **{TAB_SOLICITACOES}**: {e}")
        return

    if df.empty:
        st.info("A aba está vazia.")
        return

    # Mapeamento tolerante
    c_obj  = _pick(df, ["OBJETO DE VISTORIA", "OBJETO"])
    c_om   = _pick(df, ["OM APOIADA", "OM"])
    c_dir  = _pick(df, ["Diretoria Responsável", "Diretoria"])
    c_urg  = _pick(df, ["Classificação de Urgência", "Urgência"])
    c_sit  = _pick(df, ["Situação", "Status", "STATUS - ATUALIZAÇÃO SEMANAL"])
    c_dt_s = _pick(df, ["DATA DA SOLICITAÇÃO", "Data", "DATA DA SOLICITAÇÃO_2"])
    c_dt_v = _pick(df, ["DATA DA VISTORIA"])

    # Normalizações de data (robustas)
    for c in [c_dt_s, c_dt_v]:
        if c and c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
            try:
                # remove timezone, se houver
                df[c] = df[c].dt.tz_localize(None)
            except Exception:
                pass

    st.caption("Base: **{0}** • Registros: **{1}**".format(TAB_SOLICITACOES, len(df)))

    # Filtros simples
    colF1, colF2, colF3 = st.columns(3)
    with colF1:
        dir_sel = st.multiselect("Diretoria", sorted(df[c_dir].dropna().astype(str).unique().tolist()) if c_dir else [])
    with colF2:
        sit_sel = st.multiselect("Situação", sorted(df[c_sit].dropna().astype(str).unique().tolist()) if c_sit else [])
    with colF3:
        urg_sel = st.multiselect("Urgência", sorted(df[c_urg].dropna().astype(str).unique().tolist()) if c_urg else [])

    dff = df.copy()
    if c_dir and dir_sel: dff = dff[dff[c_dir].astype(str).isin(dir_sel)]
    if c_sit and sit_sel: dff = dff[dff[c_sit].astype(str).isin(sit_sel)]
    if c_urg and urg_sel: dff = dff[dff[c_urg].astype(str).isin(urg_sel)]

    # KPIs
    colK1, colK2, colK3, colK4 = st.columns(4)
    total = len(dff)
    pend  = dff[c_sit].astype(str).str.casefold().str.contains("não atendida|solicitad").sum() if c_sit else 0
    andam = dff[c_sit].astype(str).str.casefold().str.contains("andament|execu").sum() if c_sit else 0
    fini  = dff[c_sit].astype(str).str.casefold().str.contains("finaliz").sum() if c_sit else 0
    with colK1: st.metric("Total", f"{total:,}".replace(",", "."))
    with colK2: st.metric("Pendentes", f"{pend:,}".replace(",", "."))
    with colK3: st.metric("Em andamento", f"{andam:,}".replace(",", "."))
    with colK4: st.metric("Finalizadas", f"{fini:,}".replace(",", "."))

    st.divider()
    cols = st.columns(2)

    # Gráficos de barras e pizza
    if c_dir:
        with cols[0]:
            tmp = dff.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False)
            st.plotly_chart(px.bar(tmp, x=c_dir, y="size", title="Vistorias por Diretoria",
                                   labels={"size": "Vistorias"}), use_container_width=True)

    if c_sit:
        with cols[1]:
            tmp = dff.groupby(c_sit, as_index=False).size()
            st.plotly_chart(px.pie(tmp, names=c_sit, values="size", hole=.45,
                                   title="Distribuição por Situação",
                                   labels={"size": "Vistorias"}), use_container_width=True)

    # Evolução Mensal (corrigido/robusto)
    if c_dt_s:
        base = dff.dropna(subset=[c_dt_s]).copy()
        if not base.empty:
            # mês normalizado (primeiro dia do mês)
            base["_MES"] = base[c_dt_s].dt.to_period("M").dt.to_timestamp()

            evol = (base.groupby("_MES", as_index=False)
                        .size()
                        .rename(columns={"_MES": "MES", "size": "Vistorias"})
                        .sort_values("MES"))

            # sequência contínua de meses entre min e max
            if not evol.empty:
                idx = pd.period_range(evol["MES"].min(), evol["MES"].max(), freq="M").to_timestamp()
                evol = (evol.set_index("MES")
                             .reindex(idx, fill_value=0)
                             .rename_axis("MES")
                             .reset_index())

            st.plotly_chart(
                px.line(evol, x="MES", y="Vistorias", markers=True,
                        title="Evolução Mensal").update_layout(xaxis_title="DATA DA SOLICITAÇÃO",
                                                              yaxis_title="Vistorias"),
                use_container_width=True,
            )

            # ----- (opcional) evolução mensal por Situação empilhada -----
            if c_sit:
                por_sit = (base.assign(SITUACAO=base[c_sit].astype(str).str.strip())
                               .groupby(["_MES", "SITUACAO"], as_index=False)
                               .size()
                               .rename(columns={"_MES": "MES", "size": "Vistorias"}))

                if not por_sit.empty:
                    meses = pd.period_range(por_sit["MES"].min(), por_sit["MES"].max(), freq="M").to_timestamp()
                    todas_sit = sorted(por_sit["SITUACAO"].unique().tolist())
                    idx = pd.MultiIndex.from_product([meses, todas_sit], names=["MES", "SITUACAO"])
                    por_sit = (por_sit.set_index(["MES", "SITUACAO"])
                                     .reindex(idx, fill_value=0)
                                     .reset_index())

                    # usando altair para área empilhada (poderia ser plotly também)
                    chart_sit = (
                        alt.Chart(por_sit)
                           .mark_area()
                           .encode(
                               x=alt.X("MES:T", title="DATA DA SOLICITAÇÃO", axis=alt.Axis(format="%b %Y")),
                               y=alt.Y("Vistorias:Q", stack="zero", title="Vistorias"),
                               color=alt.Color("SITUACAO:N", title="Situação")
                           )
                           .properties(title="Evolução Mensal por Situação", height=260)
                    )
                    st.altair_chart(chart_sit, use_container_width=True)

    st.subheader("📄 Registros (completos)")
    # Mostra a tabela inteira, sem truncar para 50
    st.dataframe(dff, use_container_width=True, height=500)
