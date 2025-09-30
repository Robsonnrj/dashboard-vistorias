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
# mapeamento de abas (fica salvo na sessão)
default_tabs = {
    "solicitacoes": "ACOMPANHAMENTO VISTORIAS",   # fonte dos dashboards/solicitações
    "validacao":    "Validacao_de_Dados",         # lista oficial de OMs
    "auditoria":    "Auditoria_Vistorias",        # se você tiver essa aba
}

if "tabs_map" not in st.session_state:
    st.session_state["tabs_map"] = default_tabs.copy()

with st.sidebar.expander("⚙️ Abas usadas pelo sistema", expanded=False):
    # lista todas as abas do arquivo para facilitar
    try:
        all_tabs = [ws.title for ws in core.data_loader._book().worksheets()]
    except Exception:
        all_tabs = []
    st.session_state["tabs_map"]["solicitacoes"] = st.selectbox(
        "Aba de Solicitações / Base dos Dashboards",
        options=all_tabs or [default_tabs["solicitacoes"]],
        index=(all_tabs.index(default_tabs["solicitacoes"]) if default_tabs["solicitacoes"] in all_tabs else 0),
        key="tab_solic"
    )
    st.session_state["tabs_map"]["validacao"] = st.selectbox(
        "Aba de Validação (OMs oficiais)",
        options=all_tabs or [default_tabs["validacao"]],
        index=(all_tabs.index(default_tabs["validacao"]) if default_tabs["validacao"] in all_tabs else 0),
        key="tab_valid"
    )
    st.session_state["tabs_map"]["auditoria"] = st.selectbox(
        "Aba de Auditoria (opcional)",
        options=["(não usar)"] + (all_tabs or []),
        index=0,
        key="tab_audit"
    )

    if st.button("🔁 Atualizar dados (limpar cache)"):
        from core.data_loader import read_df
        read_df.clear()
        st.toast("Cache limpo. Recarregando…")
        st.experimental_rerun()    

if MENU == "📊 Dashboard":
    dashboard.page()
elif MENU == "📝 Cadastro":
    cadastro.page()
elif MENU == "🔁 Status/Auditoria":
    status.page()
else:
    relatorios.page()
