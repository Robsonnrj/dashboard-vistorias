# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd

from core.config import TAB_SOLICITACOES
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

def page():
    st.header("🔁 VIS-003 — Controle de Status e Auditoria")
    df = read_df(TAB_SOLICITACOES)

    if df.empty:
        st.info(f"Sem dados na aba **{TAB_SOLICITACOES}**.")
        return

    if "numero" not in df.columns:
        st.error("A planilha base precisa ter a coluna 'numero'.")
        return

    colA, colB = st.columns([2,1])
    with colA:
        numeros = pd.to_numeric(df["numero"], errors="coerce").dropna().astype(int).astype(str).tolist()
        numero = st.selectbox("Escolha a solicitação", numeros)
    with colB:
        novo = st.selectbox("Novo status", STATUS, index=STATUS.index("AGENDADA"))

    justificativa = st.text_input("Justificativa da alteração")
    responsavel = st.text_input("Responsável (login/posto)", value="usuario")

    if st.button("Atualizar status"):
        try:
            atualizar_status(numero, novo, justificativa, responsavel)
            st.success("Status atualizado e auditoria registrada.")
            st.rerun()
        except Exception as e:
            st.error(str(e))

    st.subheader("📜 Trilha de auditoria")
    hist = trilha(numero)
    st.dataframe(hist, use_container_width=True)
