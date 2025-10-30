from __future__ import annotations
import streamlit as st
from core.config import SHEET_TABS
from core import audit

st.session_state.setdefault("tabs_map", SHEET_TABS.copy())

fn = getattr(audit, "page", getattr(audit, "main", getattr(audit, "app", None)))
if callable(fn):
    fn()
else:
    st.error("A página de Auditoria não expõe page()/main()/app().")
