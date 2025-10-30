from __future__ import annotations
import streamlit as st
from core.config import SHEET_TABS
from features import cadastro

# garante o tabs_map quando a página for aberta direto por URL
st.session_state.setdefault("tabs_map", SHEET_TABS.copy())

cadastro.page()
