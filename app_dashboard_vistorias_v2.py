# -*- coding: utf-8 -*-
import streamlit as st
from streamlit_option_menu import option_menu

from core.config import has_gsheets
from features import cadastro, status, relatorios, dashboard

st.set_page_config(page_title="CRO1 — Seção de Vistorias", layout="wide")

LIGHT_CSS = """
<style>
.block-container { max-width: 1400px; padding-top: 1rem; }
[data-testid="stSidebar"] { background: #fff; border-right:1px solid #e5e7eb; }
section.main { background: #fff; }
.stButton>button { border-radius:10px; }
</style>
"""
st.markdown(LIGHT_CSS, unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Sistema CRO1 — Gestão de Vistorias")
    st.write("🔌 Google Sheets:", "ON ✅" if has_gsheets() else "OFF ❌")
    MENU = option_menu(
        "",
        ["📊 Dashboard", "📝 Cadastro", "🔁 Status/Auditoria", "📄 Relatórios (PDF)"],
        icons=["bar-chart","file-plus","arrow-repeat","file-earmark-pdf"],
        default_index=0,
        styles={
            "nav-link": {"font-size":"15px", "text-align":"left", "margin":"2px"},
            "nav-link-selected": {"background-color":"#f3f4f6"},
        }
    )

if MENU == "📊 Dashboard":
    dashboard.page()
elif MENU == "📝 Cadastro":
    cadastro.page()
elif MENU == "🔁 Status/Auditoria":
    status.page()
else:
    relatorios.page()
