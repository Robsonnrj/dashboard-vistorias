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

def page():
    st.header("🔁 VIS-003 — Controle de Status e Auditoria")
    
    tab_val = st.session_state["tabs_map"]["validacao"]
    df_oms = read_df(tab_val)
    if df.empty:
        st.info("Ainda não existem solicitações.")
        return

    colA, colB = st.columns([2,1])
    with colA:
        numero = st.selectbox("Escolha a solicitação", df["numero"].tolist())
    with colB:
        novo = st.selectbox("Novo status", STATUS, index=STATUS.index("AGENDADA"))

    justificativa = st.text_input("Justificativa da alteração")
    responsavel = st.text_input("Responsável (login/posto)", value="usuario")

    if st.button("Atualizar status"):
        try:
            atualizar_status(numero, novo, justificativa, responsavel)
            st.success("Status atualizado e auditoria registrada.")
        except Exception as e:
            st.error(str(e))

    st.subheader("📜 Trilha de auditoria")
    hist = trilha(numero)
    st.dataframe(hist, use_container_width=True)
