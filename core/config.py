# -*- coding: utf-8 -*-
import streamlit as st

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# nomes de abas no Google Sheets
SHEET_TABS = {
    "solicitacoes": "Solicitacoes",  # VIS-001 (entrada)
    "historicos":   "Historicos",    # VIS-003 (auditoria de status)
    "relatorios":   "Relatorios",    # VIS-005 (metadados de PDFs gerados)
}

def has_gsheets() -> bool:
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and "spreadsheet_url" in st.secrets["gsheets"]
        and bool(st.secrets["gsheets"]["spreadsheet_url"])
    )
