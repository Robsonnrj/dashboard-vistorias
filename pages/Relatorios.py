from __future__ import annotations
import streamlit as st
try:
    from core.config import SHEET_TABS  # type: ignore
except Exception:
    SHEET_TABS = {"solicitacoes":"Solicitacoes","historicos":"Historicos","relatorios":"Relatorios"}
st.session_state.setdefault("tabs_map", dict(SHEET_TABS))
from core.layout import hide_multipage_nav, top_nav
from features.relatorios import page as relatorios_page

def page():
    hide_multipage_nav()
    top_nav("Relatorios")

    relatorios_page()

from features import relatorios
(getattr(relatorios, "page", getattr(relatorios, "ui_relatorios", None)))()
