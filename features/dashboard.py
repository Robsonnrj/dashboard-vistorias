# -*- coding: utf-8 -*-
import pandas as pd
import plotly.express as px
import streamlit as st
import altair as alt  # opcional (usado no stacked opcional)
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

    def nf(s: str) -> str:
        return s.casefold().strip()

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
    """Normaliza strings para comparação: sem acento, maiúscula e trim."""
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.strip().upper()


# Exibição “bonita” para rótulos após normalização
_MAP_DISPLAY_SIT = {
    "AGENDADA": "Agendada",
    "CONCLUIDA": "Concluída",
    "CONCLUIDO": "Concluído",
    "EM ANDAMENTO": "Em andamento",
    "FINALIZADA": "Finalizada",
    "FINALIZADO": "Finalizado",
}


# ---------------------------
# Página
# ---------------------------
def page():
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    # Carrega base
    try:
        df = read_df(TAB_SOLICITACOES)
    except Exception as e:
        st.error(f"Não foi possível ler a aba **{TAB_SOLICITACOES}**: {e}")
        return

    if df.empty:
        st.info("A aba está vazia.")
        return

    # Mapeamento tolerante de colunas
    c_obj = _pick(df, ["OBJETO DE VISTORIA", "OBJETO"])
    c_om = _pick(df, ["OM APOIADA", "OM"])
    c_dir = _pick(df, ["Diretoria Responsável", "Diretoria"])
    c_urg = _pick(df, ["Classificação de Urgência", "Urgência"])
    c_sit = _pick(df, ["Situação", "Status", "STATUS - ATUALIZAÇÃO SEMANAL"])
    c_dt_s = _pick(df, ["DATA DA SOLICITAÇÃO", "Data", "DATA DA SOLICITAÇÃO_2"])
    c_dt_v = _pick(df, ["DATA DA VISTORIA"])
    c_dt_conc = _pick(df, ["DATA DE CONCLUSÃO", "DATA FINAL", "CONCLUÍDA EM"])

    # Normalizações de datas (robustas e sem timezone)
    for c in [c_dt_s, c_dt_v, c_dt_conc]:
        if c and c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce", utc=True)
            # remove timezone (UTC) para evitar warnings no plotly/pandas
            try:
                df[c] = df[c].dt.tz_convert(None)
            except Exception:
                # se já estiver tz-naive, ignora
                df[c] = df[c].dt.tz_localize(None)

    st.caption("Base: **{0}** • Registros: **{1}**".format(TAB_SOLICITACOES, len(df)))

    # ============================
    # Filtros
    # ============================
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

    # ============================
    # KPIs
    # ============================
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

    # ============================
    # Gráficos categoriais
    # ============================
    st.divider()
    cols = st.columns(2)

    if c_dir and c_dir in dff.columns:
        with cols[0]:
            tmp = dff.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False)
            st.plotly_chart(
                px.bar(
                    tmp,
                    x=c_dir,
                    y="size",
                    title="Vistorias por Diretoria",
                    labels={"size": "Vistorias"},
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

    # ============================
    # Evolução Mensal por DATA DA SOLICITAÇÃO
    # ============================
    if c_dt_s and c_dt_s in dff.columns:
        base = dff.dropna(subset=[c_dt_s]).copy()
        if not base.empty:
            base["_MES"] = base[c_dt_s].dt.to_period("M")
            mes_min, mes_max = base["_MES"].min(), base["_MES"].max()
            full_periods = pd.period_range(mes_min, mes_max, freq="M")
            base["_MES_STR"] = base["_MES"].astype(str)

            # Total por mês (com meses ausentes = 0)
            evol = (
                base.groupby("_MES_STR", as_index=False)
                .size()
                .rename(columns={"_MES_STR": "MÊS", "size": "Vistorias"})
                .set_index("MÊS")
                .reindex(full_periods.astype(str), fill_value=0)
                .reset_index()
                .rename(columns={"index": "MÊS"})
            )

            fig = px.line(
                evol, x="MÊS", y="Vistorias", markers=True, title="Evolução Mensal"
            )
            fig.update_layout(
                xaxis_title="DATA DA SOLICITAÇÃO",
                yaxis_title="Vistorias",
                xaxis=dict(
                    type="category",
                    categoryorder="array",
                    categoryarray=full_periods.astype(str).tolist(),
                ),
            )
            st.plotly_chart(fig, use_container_width=True)

            # Evolução Mensal por Situação (stack) — mais visível com poucos meses
            if c_sit and c_sit in base.columns:
                base_sit = base.copy()
                base_sit["Sit"] = base_sit[c_sit].astype(str).map(_norm_txt)

                por_sit = (
                    base_sit.groupby(["_MES_STR", "Sit"], as_index=False)
                    .size()
                    .rename(columns={"_MES_STR": "MÊS", "size": "Vistorias"})
                )

                if not por_sit.empty:
                    todas_sit = sorted(por_sit["Sit"].unique().tolist())
                    idx = pd.MultiIndex.from_product(
                        [full_periods.astype(str), todas_sit], names=["MÊS", "Sit"]
                    )
                    por_sit = (
                        por_sit.set_index(["MÊS", "Sit"])
                        .reindex(idx, fill_value=0)
                        .reset_index()
                    )
                    por_sit["Vistorias"] = por_sit["Vistorias"].astype(int)
                    por_sit["Sit_display"] = por_sit["Sit"].map(
                        lambda x: _MAP_DISPLAY_SIT.get(x, x.title())
                    )

                    fig2 = px.bar(
                        por_sit,
                        x="MÊS",
                        y="Vistorias",
                        color="Sit_display",
                        barmode="stack",
                        title="Evolução Mensal por Situação",
                        labels={"Vistorias": "Vistorias", "Sit_display": "Situação"},
                    )
                    fig2.update_layout(
                        xaxis_title="DATA DA SOLICITAÇÃO",
                        xaxis=dict(
                            type="category",
                            categoryorder="array",
                            categoryarray=full_periods.astype(str).tolist(),
                        ),
                    )
                    st.plotly_chart(fig2, use_container_width=True)

    # ============================
    # Evolução de Concluídas por DATA DE CONCLUSÃO
    # ============================
    if c_dt_conc and c_sit and (c_dt_conc in dff.columns) and (c_sit in dff.columns):
        concluidas = dff.copy()
        concluidas["_SIT_N"] = concluidas[c_sit].astype(str).map(_norm_txt)
        concluidas = concluidas[
            concluidas["_SIT_N"].str.contains("CONCLUID|FINALIZ", regex=True, na=False)
        ].dropna(subset=[c_dt_conc])

        if not concluidas.empty:
            concluidas["_MES_CONC"] = concluidas[c_dt_conc].dt.to_period("M")
            conc_min, conc_max = concluidas["_MES_CONC"].min(), concluidas["_MES_CONC"].max()
            full_conc = pd.period_range(conc_min, conc_max, freq="M")
            concluidas["_MES_CONC_STR"] = concluidas["_MES_CONC"].astype(str)

            evol_conc = (
                concluidas.groupby("_MES_CONC_STR", as_index=False)
                .size()
                .rename(columns={"_MES_CONC_STR": "MÊS_CONC", "size": "Concluídas"})
                .set_index("MÊS_CONC")
                .reindex(full_conc.astype(str), fill_value=0)
                .reset_index()
                .rename(columns={"index": "MÊS_CONC"})
            )

            fig_conc = px.line(
                evol_conc,
                x="MÊS_CONC",
                y="Concluídas",
                markers=True,
                title="Vistorias Concluídas por Mês",
            )
            fig_conc.update_layout(
                xaxis_title="MÊS DE CONCLUSÃO",
                yaxis_title="Vistorias Concluídas",
                xaxis=dict(
                    type="category",
                    categoryorder="array",
                    categoryarray=full_conc.astype(str).tolist(),
                ),
            )
            st.plotly_chart(fig_conc, use_container_width=True)
