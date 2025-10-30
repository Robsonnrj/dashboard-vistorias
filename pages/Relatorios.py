from __future__ import annotations
import streamlit as st
from core.config import SHEET_TABS
from features import relatorios

st.session_state.setdefault("tabs_map", SHEET_TABS.copy())

# algumas versões suas chamam ui_relatorios(); outras page()
fn = getattr(relatorios, "page", getattr(relatorios, "ui_relatorios", None))
if callable(fn):
    fn()
else:
    st.error("A página de Relatórios não expõe page() nem ui_relatorios().")
