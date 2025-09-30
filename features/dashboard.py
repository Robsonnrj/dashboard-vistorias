# -*- coding: utf-8 -*-
import pandas as pd
import plotly.express as px
import streamlit as st
from core.data_loader import read_df


# ---------------------- helpers UI ----------------------
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


def _optional_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """tenta achar a coluna por nome exato ou 'contém' (case/acentos-insensitive)."""
    if df is None or df.empty:
        return None

    cols = list(df.columns)

    # 1) exata
    for c in candidates:
        if c in cols:
            return c

    # 2) contém (normalizando)
    def n(s: str) -> str:
        return (
            str(s)
            .strip()
            .casefold()
            .replace("ã", "a")
            .replace("õ", "o")
            .replace("ç", "c")
            .replace("í", "i")
            .replace("á", "a")
            .replace("é", "e")
            .replace("ê", "e")
        )

    cand_norm = [n(x) for x in candidates]
    for c in cols:
        cc = n(c)
        if any(sn in cc for sn in cand_norm):
            return c
    return None


# ---------------------- página ----------------------
def page():
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    # onde está mapeado o nome da aba?
    # espera-se que st.session_state["tabs_map"]["solicitacoes"] exista;
    # se não existir, tentamos alguns nomes comuns como fallback.
    tab_val = None
    if "tabs_map" in st.session_state and "solicitacoes" in st.session_state["tabs_map"]:
        tab_val = st.session_state["tabs_map"]["solicitacoes"]
    else:
        # fallbacks: ajuste para o seu caso se precisar
        for guess in (
            "ACOMPANHAMENTO VISTORIAS",
            "Acompanhamento Vistorias",
            "Solicitações",
            "solicitacoes",
        ):
            try:
                # testa rapidamente se a aba existe
                tmp = read_df(guess)
                if tmp is not None:
                    tab_val = guess
                    break
            except Exception:
                pass

    if not tab_val:
        st.warning("Não foi possível determinar a aba de Solicitações. Configure 'tabs_map.solicitacoes'.")
        return

    # ---- leitura de dados
    df = read_df(tab_val)  # <<< agora 'df' existe
    if df is None or df.empty:
        st.info("Sem dados ainda nessa aba.")
        return

    # ---- mapeamento de colunas (robusto)
    c_data = _optional_col(df, ["data_limite", "DATA DA SOLICITACAO", "DATA DA SOLICITAÇÃO", "DATA"])
    c_sit = _optional_col(df, ["status_atual", "Situação", "Situacao", "STATUS"])
    c_dir = _optional_col(df, ["diretoria", "Diretoria Responsável", "Diretoria Responsavel", "Diretoria"])
    c_om = _optional_col(df, ["om_solicitante", "OM APOIADA", "OM APOIADORA", "OM"])

    # ---- filtros (só mostram opções se as colunas existirem)
    colF1, colF2 = st.columns(2)
    with colF1:
        dir_opts = sorted(df[c_dir].dropna().astype(str).unique().tolist()) if c_dir else []
        dir_sel = st.multiselect("Diretoria", dir_opts)
    with colF2:
        sit_opts = sorted(df[c_sit].dropna().astype(str).unique().tolist()) if c_sit else []
        sit_sel = st.multiselect("Status", sit_opts)

    dff = df.copy()
    if c_dir and dir_sel:
        dff = dff[dff[c_dir].astype(str).isin(dir_sel)]
    if c_sit and sit_sel:
        dff = dff[dff[c_sit].astype(str).isin(sit_sel)]

    # ---- KPIs
    total = len(dff)
    if c_sit:
        sit_up = dff[c_sit].astype(str).str.upper().str.strip()
        pend = sit_up.eq("SOLICITADA").sum()
        agend = sit_up.eq("AGENDADA").sum()
        final = sit_up.eq("FINALIZADA").sum()
    else:
        pend = agend = final = 0

    colK1, colK2, colK3, colK4 = st.columns(4)
    with colK1:
        _kpi_block("Solicitações", f"{total:,}".replace(",", "."), "Total")
    with colK2:
        _kpi_block("Pendentes", f"{pend:,}".replace(",", "."), "Status SOLICITADA")
    with colK3:
        _kpi_block("Agendadas", f"{agend:,}".replace(",", "."), "Status AGENDADA")
    with colK4:
        _kpi_block("Finalizadas", f"{final:,}".replace(",", "."), "Status FINALIZADA")

    st.divider()

    # ---- gráficos
    cols = st.columns(2)

    if c_dir and not dff.empty:
        with cols[0]:
            tmp = dff.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False)
            fig = px.bar(tmp, x=c_dir, y="size", title="Vistorias por Diretoria")
            st.plotly_chart(fig, use_container_width=True)

    if c_sit and not dff.empty:
        with cols[1]:
            tmp = dff.groupby(c_sit, as_index=False).size()
            fig = px.pie(tmp, names=c_sit, values="size", hole=0.45, title="Distribuição por Status")
            st.plotly_chart(fig, use_container_width=True)

    # ---- tabela “Últimos registros”
    st.subheader("Últimos registros")
    df_show = dff.copy()
    if c_data and c_data in df_show.columns:
        df_show[c_data] = pd.to_datetime(df_show[c_data], errors="coerce")
        df_show = df_show.sort_values(c_data, ascending=False)
    st.dataframe(df_show.head(50), use_container_width=True, height=360)
