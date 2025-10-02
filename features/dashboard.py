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
    c_dt_conc = _pick(df, ["DATA DE CONCLUSÃO", "DATA FINAL", "CONCLUÍDA EM"])

    # Normalizações de data (robustas)
    for c in [c_dt_s, c_dt_v, c_dt_conc]:
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

    # Calcular dias de conclusão se possível
    if c_dt_s and c_dt_conc:
        dff["DIAS_CONCLUSAO"] = (dff[c_dt_conc] - dff[c_dt_s]).dt.days
        dias_conc = dff[dff[c_sit].astype(str).str.contains("conclu", case=False, na=False)]["DIAS_CONCLUSAO"]
        with st.expander("Tempo Médio de Conclusão (dias)"):
            if not dias_conc.empty:
                st.metric("Dias médios", f"{dias_conc.mean():.2f}")
                st.write(dias_conc.describe())
            else:
                st.write("Sem dados suficientes para calcular.")

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

    
    # Evolução Mensal por Data da Solicitacao
    if c_dt_s:
        base = dff.dropna(subset=[c_dt_s]).copy()
        if not base.empty:
            base["_MES"] = base[c_dt_s].dt.to_period("M")
            month_order = sorted(base["_MES"].unique().tolist())
            base["_MES_STR"] = base["_MES"].astype(str)

            evol = (
                base.groupby("_MES_STR", as_index=False)
                    .size()
                    .rename(columns={"_MES_STR": "MÊS", "size": "Vistorias"})
            )
            full_months = [str(m) for m in month_order]
            evol = (evol.set_index("MÊS")
                        .reindex(full_months, fill_value=0)
                        .rename_axis("MÊS")
                        .reset_index())

            fig = px.line(evol, x="MÊS", y="Vistorias", markers=True, title="Evolução Mensal")
            fig.update_layout(
                xaxis_title="DATA DA SOLICITAÇÃO",
                yaxis_title="Vistorias",
                xaxis=dict(type="category", categoryorder="array", categoryarray=full_months),
            )
            st.plotly_chart(fig, use_container_width=True)

            # Evolução Mensal por Situação (empilhado) opcional
            if c_sit:
                por_sit = (
                    base.assign(Sit=base[c_sit].astype(str).str.strip())
                        .groupby(["_MES_STR", "Sit"], as_index=False)
                        .size()
                        .rename(columns={"_MES_STR": "MÊS", "size": "Vistorias"})
                )

                if not por_sit.empty:
                    todas_sit = sorted(por_sit["Sit"].unique().tolist())
                    full_idx = pd.MultiIndex.from_product([full_months, todas_sit], names=["MÊS", "Sit"])
                    por_sit = (por_sit.set_index(["MÊS", "Sit"])
                                     .reindex(full_idx, fill_value=0)
                                     .reset_index())

                    fig2 = px.area(
                        por_sit, x="MÊS", y="Vistorias", color="Sit",
                        title="Evolução Mensal por Situação"
                    )
                    fig2.update_layout(
                        xaxis_title="DATA DA SOLICITAÇÃO",
                        yaxis_title="Vistorias",
                        xaxis=dict(type="category", categoryorder="array", categoryarray=full_months),
                    )
                    st.plotly_chart(fig2, use_container_width=True)

    # Evolução mensal das vistorias concluídas por data de conclusão
    if c_dt_conc and c_sit:
        concluidas = dff[dff[c_sit].astype(str).str.contains("conclu", case=False, na=False)].copy()
        concluidas = concluidas.dropna(subset=[c_dt_conc])
        if not concluidas.empty:
            concluidas["_MES_CONC"] = concluidas[c_dt_conc].dt.to_period("M")
            month_order_conc = sorted(concluidas["_MES_CONC"].unique().tolist())
            concluidas["_MES_CONC_STR"] = concluidas["_MES_CONC"].astype(str)

            evol_conc = (
                concluidas.groupby("_MES_CONC_STR", as_index=False).size()
                .rename(columns={"_MES_CONC_STR": "MÊS_CONC", "size": "Concluídas"})
            )
            full_months_conc = [str(m) for m in month_order_conc]
            evol_conc = (
                evol_conc.set_index("MÊS_CONC")
                .reindex(full_months_conc, fill_value=0)
                .rename_axis("MÊS_CONC")
                .reset_index()
            )

            fig_conc = px.line(evol_conc, x="MÊS_CONC", y="Concluídas", markers=True, title="Vistorias Concluídas por Mês")
            fig_conc.update_layout(
                xaxis_title="MÊS DE CONCLUSÃO",
                yaxis_title="Vistorias Concluídas",
                xaxis=dict(type="category", categoryorder="array", categoryarray=full_months_conc),
            )
            st.plotly_chart(fig_conc, use_container_width=True)
