# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from core.data_loader import read_df
from core.audit import atualizar_status, trilha

STATUS = [
    "SOLICITADA",
    "AGENDADA",
    "EM_EXECUCAO",
    "FINALIZADA",
    "RELATORIO_GERADO",
    "INTEGRADA_OBRAS",
]

def _pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Retorna a primeira coluna do DF que 'bate' com a lista de candidatos (tolerante a acentos/caixa)."""
    if df is None or df.empty:
        return None
    def norm(s: str) -> str:
        import unicodedata
        s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
        return s.strip().lower()
    cols = list(df.columns)
    # match exato normalizado
    for cand in candidates:
        for c in cols:
            if norm(c) == norm(cand):
                return c
    # match por "contém"
    for cand in candidates:
        target = norm(cand)
        for c in cols:
            if target in norm(c):
                return c
    return None

def page():
    st.header("🔁 VIS-003 — Controle de Status e Auditoria")

    # ---- base de solicitações (aba principal)
    tab_base = st.session_state["tabs_map"]["solicitacoes"]
    df = read_df(tab_base)
    if df.empty:
        st.info("Sem dados para auditoria.")
        return

    # identifica coluna do número (tolerante a variações)
    col_num = _pick_col(df, ["numero", "número", "num", "nº", "id", "protocolo"])
    if not col_num:
        st.error("Não encontrei a coluna de número da solicitação na aba base.")
        st.dataframe(df.head(10), use_container_width=True)
        return

    # Ordena por número (quando possível) para facilitar a escolha
    try:
        df_view = df.copy()
        df_view[col_num] = df_view[col_num].astype(str)
        numeros = df_view[col_num].dropna().unique().tolist()
        numeros = sorted(numeros, key=lambda x: (len(x), x))  # ordenação estável
    except Exception:
        numeros = df[col_num].dropna().astype(str).unique().tolist()

    colA, colB = st.columns([2, 1])
    with colA:
        numero = st.selectbox("Escolha a solicitação", numeros, index=0 if numeros else None)
    with colB:
        novo = st.selectbox("Novo status", STATUS, index=STATUS.index("AGENDADA"))

    justificativa = st.text_input("Justificativa da alteração", placeholder="Descreva o motivo da mudança")
    responsavel = st.text_input("Responsável (login/posto)", value="usuario")

    if st.button("Atualizar status", type="primary", disabled=not numero):
        try:
            atualizar_status(numero, novo, justificativa, responsavel)
            # limpa cache para refletir a mudança imediatamente
            try:
                read_df.clear()
            except Exception:
                pass
            st.success(f"Status da solicitação {numero} atualizado para {novo} e auditoria registrada.")
        except Exception as e:
            st.error(f"Falha ao atualizar: {e}")

    st.subheader("📜 Trilha de auditoria")
    if numero:
        try:
            hist = trilha(numero)
            if isinstance(hist, pd.DataFrame) and not hist.empty:
                st.dataframe(hist, use_container_width=True, height=360)
            else:
                st.info("Sem registros de auditoria para esta solicitação.")
        except Exception as e:
            st.warning(f"Não foi possível carregar a trilha de auditoria: {e}")
