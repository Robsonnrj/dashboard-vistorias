# -*- coding: utf-8 -*-

# Abas fixas usadas pelo sistema
TAB_SOLICITACOES = "ACOMPANHAMENTO VISTORIAS"   # base de registros (cadastro, dashboards, status)
TAB_VALIDACAO    = "Validacao_de_Dados"         # lista oficial de OMs

def has_gsheets() -> bool:
    import streamlit as st
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and st.secrets["gsheets"].get("spreadsheet_url")
    )
