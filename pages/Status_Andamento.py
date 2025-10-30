from __future__ import annotations
import streamlit as st
from core.config import SHEET_TABS
from features import status

st.session_state.setdefault("tabs_map", SHEET_TABS.copy())

status.page()
