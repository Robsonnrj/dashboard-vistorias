# -*- coding: utf-8 -*-
"""
core/audit.py
Página de Auditoria — filtros na sidebar, KPIs, gráficos e tabela + download.
Robusto a variações de cabeçalhos e valores ausentes.
"""

from __future__ import annotations

import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

from core.config import TAB_AUDIT
from core.data_loader import read_df


# ----------------------------- helpers ----------------------------- #
def _nf(x) -> str:
    """Normaliza: NaN -> '', demais -> str.strip()."""
    return "" if pd.isna(x) else str(x).strip()


def _first_col(df: pd.DataFrame, *cands: str) -> str | None:
    """
    Retorna a primeira coluna do DF cujo nome (normalizado) bate com os candidatos.
    Tenta match exato (casefold) e depois 'contém'.
    """
    if df is None or df.empty:
        return None

    def norm(s: str) -> str:
        import unicodedata as _ud
        s = _ud.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
        return s.casefold().strip()

    cols = list(df.columns)
    nmap = {norm(c): c for c in cols}

    # exato
    for want in cands:
        if norm(want) in nmap:
            return nmap[norm(want)]

    # contém
    for want in cands:
        tgt = norm(want)
        for c in cols:
            if tgt in norm(c):
                return c
    return None


# ------------------------------- page ------------------------------ #
def page():
    st.header("🔎 Registro de Auditoria — Trilhas de Alterações")

    # ================== carga ==================
    try:
        df_raw = read_df(TAB_AUDIT)
    except Exception as e:
        st.error(f"Não consegui ler a aba de auditoria '{TAB_AUDIT}': {e}")
        return

    if df_raw is None or df_raw.empty:
        st.info("Ainda não há registros de auditoria.")
        return

    df = df_raw.copy()

    # Mapeamento tolerante de colunas
    c_num   = _first_col(df, "numero", "número", "num", "id", "protocolo")
    c_ts    = _first_col(df, "timestamp", "ts", "data", "quando", "em")
    c_campo = _first_col(df, "campo", "coluna", "atributo", "chave")
    c_antes = _first_col(df, "antes", "valor_anterior", "de")
    c_depois= _first_col(df, "depois", "valor_novo", "para")

    # Checagem mínima
    faltam = [n for (n, c) in [("numero", c_num), ("timestamp", c_ts), ("campo", c_campo)] if not c]
    if faltam:
        st.error("A planilha de auditoria precisa ter, no mínimo, as colunas: "
                 + ", ".join(faltam) + ".")
        st.caption("Colunas detectadas: " + ", ".join(df.columns))
        return

    # Normalizações
    df[c_ts] = pd.to_datetime(df[c_ts], errors="coerce")
    for c in [c_num, c_campo, c_antes, c_depois]:
        if c and c in df.columns:
            df[c] = df[c].astype(str).map(_nf)

    # ================== filtros (sidebar) ==================
    with st.sidebar:
        st.subheader("Filtros — Auditoria")

        # Período
        dt_min, dt_max = df[c_ts].min(), df[c_ts].max()
        ini_default = (dt_min.date() if pd.notna(dt_min) else pd.Timestamp.today().date())
        fim_default = (dt_max.date() if pd.notna(dt_max) else pd.Timestamp.today().date())
        periodo = st.date_input("Período", (ini_default, fim_default))
        if isinstance(periodo, (list, tuple)) and len(periodo) == 2:
            ini, fim = periodo
        else:
            ini = fim_default
            fim = fim_default

        # Números e campos
        ops_num = sorted(df[c_num].dropna().unique().tolist())
        ops_cam = sorted(df[c_campo].dropna().unique().tolist())

        numeros = st.multiselect("Número", ops_num)
        campos  = st.multiselect("Campo alterado", ops_cam)

        termo = st.text_input("Contém (antes/depois)", placeholder="texto livre")

    # ================== aplica filtros ==================
    mask = pd.Series(True, index=df.index)

    # período (inclui o dia final completo)
    ini_ts = pd.to_datetime(ini)
    fim_ts = pd.to_datetime(fim) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    mask &= df[c_ts].between(ini_ts, fim_ts)

    if numeros:
        mask &= df[c_num].isin(numeros)
    if campos:
        mask &= df[c_campo].isin(campos)

    if termo:
        termo_cf = termo.casefold()
        # OR entre colunas de texto (antes/depois)
        any_series = pd.Series(False, index=df.index)
        for c in [c_antes, c_depois]:
            if c and c in df.columns:
                any_series |= df[c].astype(str).str.casefold().str.contains(termo_cf, na=False)
        mask &= any_series

    dff = df.loc[mask].copy()

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
    else:
        st.caption("Sem dados suficientes para série temporal.")

    # 2) Barras — Top campos alterados
    gr_campos = (
        dff.groupby(c_campo, as_index=False)
           .size()
           .sort_values("size", ascending=False)
           .head(15)
    )
    if not gr_campos.empty:
        fig_cam = px.bar(gr_campos, x="size", y=c_campo, orientation="h",
                         title="Top campos alterados")
        st.plotly_chart(fig_cam, use_container_width=True)

    # 3) Barras — Top números (registros) mais alterados
    gr_nums = (
        dff.groupby(c_num, as_index=False)
           .size()
           .sort_values("size", ascending=False)
           .head(15)
    )
    if not gr_nums.empty:
        fig_num = px.bar(gr_nums, x="size", y=c_num, orientation="h",
                         title="Registros com mais alterações")
        st.plotly_chart(fig_num, use_container_width=True)

    # ================== detalhe por número ==================
    with st.expander("🔍 Ver histórico detalhado por Número"):
        ops_num_all = [""] + ops_num
        num_sel = st.selectbox("Escolha um número", ops_num_all, index=0)
        if num_sel:
            h = dff[dff[c_num] == num_sel].sort_values(c_ts, ascending=False)
            st.dataframe(h, use_container_width=True, hide_index=True)

    # ================== tabela + download ==================
    st.subheader("Tabela de Auditoria (filtrada)")
    st.dataframe(
        dff.sort_values(c_ts, ascending=False),
        use_container_width=True,
        hide_index=True,
    )

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


# aliases para compatibilidade (se chamarem main/app)
main = page
app = page
