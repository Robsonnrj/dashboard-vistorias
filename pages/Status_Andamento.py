from __future__ import annotations
import streamlit as st
try:
    from core.config import SHEET_TABS  # type: ignore
except Exception:
    SHEET_TABS = {"solicitacoes":"Solicitacoes","historicos":"Historicos","relatorios":"Relatorios"}
st.session_state.setdefault("tabs_map", dict(SHEET_TABS))

from features import status
from core.layout import hide_multipage_nav, top_nav
from features.status import page as status_page

def page():
    hide_multipage_nav()
    top_nav("Status")

    status_page()

status.page()
