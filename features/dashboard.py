# -*- coding: utf-8 -*-
import pandas as pd
import plotly.express as px
import streamlit as st
import altair as alt  # opcional
import unicodedata
from core.data_loader import read_df
from core.config import TAB_SOLICITACOES


# ---------------------------
# Helpers
# ---------------------------
def _pick(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Encontra coluna pela lista de candidatos (match exato ou contém, casefold)."""
    if df.empty:
        return None
    cols = list(df.columns)
    def nf(s: str) -> str: return s.casefold().strip()
    # exato
    for c in candidates:
        for cc in cols:
            if nf(cc) == nf(c): return cc
    # contém
    for c in candidates:
        target = nf(c)
        for cc in cols:
            if target in nf(cc): return cc
    return None

def _norm_txt(s: str) -> str:
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    return s.strip().upper()

_MAP_DISPLAY_SIT = {
    "AGENDADA":"Agendada",
    "CONCLUIDA":"Concluída", "CONCLUIDO":"Concluído",
    "EM ANDAMENTO":"Em andamento",
    "FINALIZADA":"Finalizada", "FINALIZADO":"Finalizado",
}


# ---------------------------
# Página
# ---------------------------
def page():
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    # Base
    try:
        df = read_df(TAB_SOLICITACOES)
    except Exception as e:
        st.error(f"Não foi possível ler a aba **{TAB_SOLICITACOES}**: {e}")
        return
    if df.empty:
        st.info("A aba está vazia.")
        return

    # Mapear colunas
    c_dir   = _pick(df, ["Diretoria Responsável", "Diretoria"])
    c_sit   = _pick(df, ["Situação", "Status", "STATUS - ATUALIZAÇÃO SEMANAL"])
    c_urg   = _pick(df, ["Classificação da Urgência", "Urgência"])
    c_dt_s  = _pick(df, ["DATA DA SOLICITAÇÃO", "Data", "DATA DA SOLICITAÇÃO_2"])
    c_dt_v  = _pick(df, ["DATA DA VISTORIA"])
    c_dt_c  = _pick(df, ["DATA DE CONCLUSÃO", "DATA FINAL", "CONCLUÍDA EM"])
    c_dias_exec = _pick(df, ["QUANTIDADE DE DIAS PARA EXECUÇÃO", "DIAS EXECUCAO"])
    c_dias_total = _pick(df, ["QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO", "DIAS ATENDIMENTO TOTAL"])

    # Datas -> datetime sem timezone
    for c in [c_dt_s, c_dt_v, c_dt_c]:
        if c and c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    # Números
    for c in [c_dias_exec, c_dias_total]:
        if c and c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    st.caption(f"Base: **{TAB_SOLICITACOES}** • Registros: **{len(df)}**")

    # ---------------- Filtros ----------------
    colF1, colF2, colF3 = st.columns(3)
    with colF1:
        dir_sel = st.multiselect("Diretoria",
            sorted(df[c_dir].dropna().astype(str).unique().tolist()) if c_dir else [])
    with colF2:
        sit_sel = st.multiselect("Situação",
            sorted(df[c_sit].dropna().astype(str).unique().tolist()) if c_sit else [])
    with colF3:
        urg_sel = st.multiselect("Urgência",
            sorted(df[c_urg].dropna().astype(str).unique().tolist()) if c_urg else [])

    dff = df.copy()
    if c_dir and dir_sel: dff = dff[dff[c_dir].astype(str).isin(dir_sel)]
    if c_sit and sit_sel: dff = dff[dff[c_sit].astype(str).isin(sit_sel)]
    if c_urg and urg_sel: dff = dff[dff[c_urg].astype(str).isin(urg_sel)]

    # ---------------- KPIs ----------------
    colK1, colK2, colK3, colK4 = st.columns(4)
    total = len(dff)
    if c_sit and c_sit in dff.columns:
        _sit_norm = dff[c_sit].astype(str).map(_norm_txt)
        pend  = _sit_norm.str.contains("NAO ATENDIDA|SOLICITAD", regex=True, na=False).sum()
        andam = _sit_norm.str.contains("ANDAMENT|EXECU",        regex=True, na=False).sum()
        fini  = _sit_norm.str.contains("FINALIZ|CONCLUID",       regex=True, na=False).sum()
    else:
        pend = andam = fini = 0
    with colK1: st.metric("Total", f"{total:,}".replace(",", "."))
    with colK2: st.metric("Pendentes", f"{pend:,}".replace(",", "."))
    with colK3: st.metric("Em andamento", f"{andam:,}".replace(",", "."))
    with colK4: st.metric("Finalizadas", f"{fini:,}".replace(",", "."))

    st.divider()

    # =========================================================
    # NOVOS GRÁFICOS “FUNÇÃO” (evolução temporal de desempenho)
    # Substituem: Evolução Mensal e Evolução por Situação (contagens)
    # =========================================================

    # -------- 1) Desempenho de Execução vs Data da Solicitação --------
    if c_dt_s and c_dias_exec and (c_dt_s in dff.columns) and (c_dias_exec in dff.columns):
        base = dff.dropna(subset=[c_dt_s, c_dias_exec]).copy()
        if not base.empty:
            # Dispersão por ponto (cada vistoria)
            base["_Sit_display"] = (base[c_sit].astype(str).map(_norm_txt)
                                    .map(lambda x: _MAP_DISPLAY_SIT.get(x, x.title()))
                                    if c_sit in base.columns else "—")

            fig_scatter = px.scatter(
                base, x=c_dt_s, y=c_dias_exec, color="_Sit_display",
                title="Dias para Execução × Data da Solicitação (ponto a ponto)",
                labels={c_dt_s:"Data da Solicitação", c_dias_exec:"Dias para Execução", "_Sit_display":"Situação"},
                hover_data=[c_dir] if c_dir in base.columns else None
            )
            st.plotly_chart(fig_scatter, use_container_width=True)

            # Linha mensal (média) + média móvel (3 meses)
            base["_MES"] = base[c_dt_s].dt.to_period("M").astype(str)
            evol = (base.groupby("_MES", as_index=False)[c_dias_exec]
                         .mean(numeric_only=True)
                         .rename(columns={c_dias_exec:"Dias Médios Execução"}))
            # preencher meses ausentes
            pr = pd.period_range(base[c_dt_s].min().to_period("M"), base[c_dt_s].max().to_period("M"), freq="M")
            evol = (evol.set_index("_MES")
                        .reindex(pr.astype(str))
                        .rename_axis("_MES").reset_index())
            evol["Dias Médios Execução"] = evol["Dias Médios Execução"].astype(float)
            # média móvel
            evol["Média Móvel (3m)"] = evol["Dias Médios Execução"].rolling(3, min_periods=1).mean()

            fig_line = px.line(
                evol, x="_MES", y=["Dias Médios Execução","Média Móvel (3m)"],
                markers=True, title="Tendência Mensal — Dias para Execução",
                labels={"_MES":"Mês", "value":"Dias"}
            )
            fig_line.update_layout(xaxis=dict(type="category", categoryorder="array", categoryarray=pr.astype(str).tolist()))
            st.plotly_chart(fig_line, use_container_width=True)

    st.divider()

    # -------- 2) Desempenho de Atendimento Total vs Data de Conclusão --------
    if c_dt_c and c_dias_total and (c_dt_c in dff.columns) and (c_dias_total in dff.columns):
        base2 = dff.dropna(subset=[c_dt_c, c_dias_total]).copy()
        if not base2.empty:
            base2["_Sit_display"] = (base2[c_sit].astype(str).map(_norm_txt)
                                     .map(lambda x: _MAP_DISPLAY_SIT.get(x, x.title()))
                                     if c_sit in base2.columns else "—")

            fig_scatter2 = px.scatter(
                base2, x=c_dt_c, y=c_dias_total, color="_Sit_display",
                title="Dias para Atendimento Total × Data de Conclusão (ponto a ponto)",
                labels={c_dt_c:"Data de Conclusão", c_dias_total:"Dias p/ Atendimento Total", "_Sit_display":"Situação"},
                hover_data=[c_dir] if c_dir in base2.columns else None
            )
            st.plotly_chart(fig_scatter2, use_container_width=True)

            base2["_MES"] = base2[c_dt_c].dt.to_period("M").astype(str)
            evol2 = (base2.groupby("_MES", as_index=False)[c_dias_total]
                           .mean(numeric_only=True)
                           .rename(columns={c_dias_total:"Dias Médios Atendimento"}))
            pr2 = pd.period_range(base2[c_dt_c].min().to_period("M"), base2[c_dt_c].max().to_period("M"), freq="M")
            evol2 = (evol2.set_index("_MES")
                         .reindex(pr2.astype(str))
                         .rename_axis("_MES").reset_index())
            evol2["Dias Médios Atendimento"] = evol2["Dias Médios Atendimento"].astype(float)
            evol2["Média Móvel (3m)"] = evol2["Dias Médios Atendimento"].rolling(3, min_periods=1).mean()

            fig_line2 = px.line(
                evol2, x="_MES", y=["Dias Médios Atendimento","Média Móvel (3m)"],
                markers=True, title="Tendência Mensal — Dias para Atendimento Total",
                labels={"_MES":"Mês", "value":"Dias"}
            )
            fig_line2.update_layout(xaxis=dict(type="category", categoryorder="array", categoryarray=pr2.astype(str).tolist()))
            st.plotly_chart(fig_line2, use_container_width=True)

    # (opcional) se quiser manter os gráficos antigos, coloque-os abaixo de um st.expander().
