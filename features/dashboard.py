# -*- coding: utf-8 -*-
import pandas as pd
import plotly.express as px
import streamlit as st
import altair as alt  # opcional
import unicodedata
from core.data_loader import read_df
from core.config import TAB_SOLICITACOES


# =========================
# Helpers
# =========================
def _pick(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Encontra coluna pela lista de candidatos (match exato ou contém, casefold)."""
    if df.empty:
        return None
    cols = list(df.columns)

    def nf(s: str) -> str:
        return str(s).casefold().strip()

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


def _norm_txt(s: str) -> str:
    """Normaliza texto para comparação (sem acento, upper, trim)."""
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.strip().upper()


_MAP_DISPLAY_SIT = {
    "AGENDADA": "Agendada",
    "CONCLUIDA": "Concluída",
    "CONCLUIDO": "Concluído",
    "EM ANDAMENTO": "Em andamento",
    "FINALIZADA": "Finalizada",
    "FINALIZADO": "Finalizado",
}


# ===== funções p/ séries “função” =====
def _as_date(s):
    """Converte para datetime (naive) e trunca para dia."""
    s = pd.to_datetime(s, errors="coerce")
    try:
        s = s.dt.tz_convert(None)
    except Exception:
        try:
            s = s.dt.tz_localize(None)
        except Exception:
            pass
    return s.dt.floor("D")


def _daily_range(dmin, dmax):
    return pd.date_range(dmin, dmax, freq="D") if pd.notna(dmin) and pd.notna(dmax) else pd.DatetimeIndex([])


def _line_fig(df_line: pd.DataFrame, title: str, xlab: str, ylab: str):
    # df_line precisa ter colunas: DATA, Valor, Média Móvel (7d)
    fig = px.line(
        df_line,
        x="DATA",
        y=["Valor", "Média Móvel (7d)"],
        markers=False,
        title=title,
        labels={"DATA": xlab, "value": ylab, "variable": "Série"},
    )
    fig.update_traces(mode="lines")  # força linhas (sem marcadores)
    fig.update_layout(xaxis=dict(type="date"))
    return fig


# =========================
# Página
# =========================
def page():
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    # Carregar base
    try:
        df = read_df(TAB_SOLICITACOES)
    except Exception as e:
        st.error(f"Não foi possível ler a aba **{TAB_SOLICITACOES}**: {e}")
        return

    if df.empty:
        st.info("A aba está vazia.")
        return

    # Mapear colunas
    c_obj = _pick(df, ["OBJETO DE VISTORIA", "OBJETO"])
    c_om = _pick(df, ["OM APOIADA", "OM"])
    c_dir = _pick(df, ["Diretoria Responsável", "Diretoria"])
    c_urg = _pick(df, ["Classificação da Urgência", "Urgência"])
    c_sit = _pick(df, ["Situação", "Status", "STATUS - ATUALIZAÇÃO SEMANAL"])
    c_dt_s = _pick(df, ["DATA DA SOLICITAÇÃO", "Data", "DATA DA SOLICITAÇÃO_2"])
    c_dt_v = _pick(df, ["DATA DA VISTORIA"])
    c_dt_c = _pick(df, ["DATA DE CONCLUSÃO", "DATA FINAL", "CONCLUÍDA EM"])
    c_dias_exec = _pick(df, ["QUANTIDADE DE DIAS PARA EXECUÇÃO", "DIAS EXECUCAO"])
    c_dias_total = _pick(df, ["QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO", "DIAS ATENDIMENTO TOTAL"])

    # Datas
    for c in [c_dt_s, c_dt_v, c_dt_c]:
        if c and c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # Numéricos
    for c in [c_dias_exec, c_dias_total]:
        if c and c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    st.caption("Base: **{0}** • Registros: **{1}**".format(TAB_SOLICITACOES, len(df)))

    # =========================
    # Filtros
    # =========================
    colF1, colF2, colF3 = st.columns(3)
    with colF1:
        dir_sel = st.multiselect(
            "Diretoria",
            sorted(df[c_dir].dropna().astype(str).unique().tolist()) if c_dir else [],
        )
    with colF2:
        sit_sel = st.multiselect(
            "Situação",
            sorted(df[c_sit].dropna().astype(str).unique().tolist()) if c_sit else [],
        )
    with colF3:
        urg_sel = st.multiselect(
            "Urgência",
            sorted(df[c_urg].dropna().astype(str).unique().tolist()) if c_urg else [],
        )

    dff = df.copy()
    if c_dir and dir_sel:
        dff = dff[dff[c_dir].astype(str).isin(dir_sel)]
    if c_sit and sit_sel:
        dff = dff[dff[c_sit].astype(str).isin(sit_sel)]
    if c_urg and urg_sel:
        dff = dff[dff[c_urg].astype(str).isin(urg_sel)]

    # =========================
    # KPIs
    # =========================
    colK1, colK2, colK3, colK4 = st.columns(4)
    total = len(dff)
    if c_sit and c_sit in dff.columns:
        _sit_norm = dff[c_sit].astype(str).map(_norm_txt)
        pend = _sit_norm.str.contains("NAO ATENDIDA|SOLICITAD", regex=True, na=False).sum()
        andam = _sit_norm.str.contains("ANDAMENT|EXECU", regex=True, na=False).sum()
        fini = _sit_norm.str.contains("FINALIZ|CONCLUID", regex=True, na=False).sum()
    else:
        pend = andam = fini = 0
    with colK1:
        st.metric("Total", f"{total:,}".replace(",", "."))
    with colK2:
        st.metric("Pendentes", f"{pend:,}".replace(",", "."))
    with colK3:
        st.metric("Em andamento", f"{andam:,}".replace(",", "."))
    with colK4:
        st.metric("Finalizadas", f"{fini:,}".replace(",", "."))

    # =========================
    # Gráficos categoriais (mantidos)
    # =========================
    st.divider()
    cols = st.columns(2)

    if c_dir and c_dir in dff.columns:
        with cols[0]:
            tmp = dff.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False)
            st.plotly_chart(
                px.bar(
                    tmp, x=c_dir, y="size", title="Vistorias por Diretoria", labels={"size": "Vistorias"}
                ),
                use_container_width=True,
            )

    if c_sit and c_sit in dff.columns:
        with cols[1]:
            tmp = dff.copy()
            tmp["Sit_display"] = tmp[c_sit].astype(str).map(_norm_txt).map(
                lambda x: _MAP_DISPLAY_SIT.get(x, x.title())
            )
            tmp = tmp.groupby("Sit_display", as_index=False).size()
            st.plotly_chart(
                px.pie(
                    tmp,
                    names="Sit_display",
                    values="size",
                    hole=0.45,
                    title="Distribuição por Situação",
                    labels={"size": "Vistorias"},
                ),
                use_container_width=True,
            )

    # =========================
    # Séries temporais “função”
    # =========================
    with st.expander("📈 Séries temporais (curvas de função)", expanded=True):
        # ---- 1) Dias para Execução (DATA DA SOLICITAÇÃO) ----
        if c_dt_s and c_dias_exec and (c_dt_s in dff.columns) and (c_dias_exec in dff.columns):
            base = dff.dropna(subset=[c_dt_s, c_dias_exec]).copy()
            if not base.empty:
                base["DATA_SOLICITACAO_D"] = _as_date(base[c_dt_s])

                # média diária
                by_day = (
                    base.groupby("DATA_SOLICITACAO_D", as_index=False)[c_dias_exec]
                    .mean(numeric_only=True)
                    .rename(columns={c_dias_exec: "Valor"})
                )
                dr = _daily_range(by_day["DATA_SOLICITACAO_D"].min(), by_day["DATA_SOLICITACAO_D"].max())
                line = (
                    by_day.set_index("DATA_SOLICITACAO_D")
                    .reindex(dr, fill_value=None)
                    .rename_axis("DATA")
                    .reset_index()
                )
                line["Média Móvel (7d)"] = line["Valor"].rolling(7, min_periods=1).mean()

                fig_exec = _line_fig(
                    line,
                    "Dias para Execução — Série Diária",
                    "Data (Solicitação)",
                    "Dias",
                )
                st.plotly_chart(fig_exec, use_container_width=True)

        # ---- 2) Dias para Atendimento Total (DATA DE CONCLUSÃO) ----
        if c_dt_c and c_dias_total and (c_dt_c in dff.columns) and (c_dias_total in dff.columns):
            base2 = dff.dropna(subset=[c_dt_c, c_dias_total]).copy()
            if not base2.empty:
                base2["DATA_CONCLUSAO_D"] = _as_date(base2[c_dt_c])

                by_day2 = (
                    base2.groupby("DATA_CONCLUSAO_D", as_index=False)[c_dias_total]
                    .mean(numeric_only=True)
                    .rename(columns={c_dias_total: "Valor"})
                )
                dr2 = _daily_range(by_day2["DATA_CONCLUSAO_D"].min(), by_day2["DATA_CONCLUSAO_D"].max())
                line2 = (
                    by_day2.set_index("DATA_CONCLUSAO_D")
                    .reindex(dr2, fill_value=None)
                    .rename_axis("DATA")
                    .reset_index()
                )
                line2["Média Móvel (7d)"] = line2["Valor"].rolling(7, min_periods=1).mean()

                fig_total = _line_fig(
                    line2,
                    "Dias para Atendimento Total — Série Diária",
                    "Data (Conclusão)",
                    "Dias",
                )
                st.plotly_chart(fig_total, use_container_width=True)

        # ---- 3) Backlog(t): vistorias abertas ao longo do tempo ----
        if c_dt_s and (c_dt_s in dff.columns):
            b = dff.copy()
            b["DS"] = _as_date(b[c_dt_s])

            ev = b.dropna(subset=["DS"])[["DS"]].assign(delta=1).rename(columns={"DS": "DATA"})
            if c_dt_c and (c_dt_c in dff.columns):
                b["DC"] = _as_date(b[c_dt_c])
                ev2 = b.dropna(subset=["DC"])[["DC"]].assign(delta=-1).rename(columns={"DC": "DATA"})
                ev = pd.concat([ev, ev2], ignore_index=True)

            if not ev.empty:
                ev = ev.groupby("DATA", as_index=False)["delta"].sum()
                drb = _daily_range(ev["DATA"].min(), ev["DATA"].max())
                ser = (
                    ev.set_index("DATA")
                    .reindex(drb, fill_value=0)
                    .rename_axis("DATA")
                    .reset_index()
                )
                ser["Valor"] = ser["delta"].cumsum()
                ser["Média Móvel (7d)"] = ser["Valor"].rolling(7, min_periods=1).mean()

                fig_backlog = _line_fig(
                    ser.drop(columns=["delta"]),
                    "Backlog de Vistorias — Abertas ao Longo do Tempo",
                    "Data",
                    "Vistorias Abertas",
                )
                st.plotly_chart(fig_backlog, use_container_width=True)
