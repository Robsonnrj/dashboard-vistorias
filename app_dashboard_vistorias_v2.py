# -*- coding: utf-8 -*-
"""
Aplicativo principal do Sistema de Vistorias CRO/1
Navegação: Cadastro | Dashboard Operacional | Relatórios | Status/Andamento | Auditoria
"""

from __future__ import annotations
import streamlit as st

# Importa módulos de features
from features import cadastro, status, relatorios, dashboard
from core import audit

# Configuração do layout e ícone
st.set_page_config(
    page_title="CRO/1 — Sistema de Vistorias",
    layout="wide",
    page_icon="🛠️"
)

# --------------------------------------------------------
# HOME (menu inicial com ícones)
# --------------------------------------------------------
if "secao" not in st.session_state:
    st.session_state["secao"] = "Home"

if st.session_state["secao"] == "Home":
    st.markdown("""
        <style>
        .block-container { padding-top: 2rem; padding-bottom: 3rem; text-align:center; }
        .icon-btn {
            display:inline-flex; flex-direction:column; align-items:center; 
            justify-content:center; margin:1rem;
        }
        .icon-btn button[kind="secondary"] {
            height:120px; width:120px; font-size:60px; border-radius:24px;
        }
        .icon-caption { margin-top:8px; font-weight:600; font-size:0.95rem; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1>Navegação</h1>", unsafe_allow_html=True)
    st.markdown("<p>Clique em um ícone para abrir a seção</p>", unsafe_allow_html=True)

    def icon_card(icon: str, label: str, target: str):
        col = st.container()
        with col:
            st.markdown("<div class='icon-btn'>", unsafe_allow_html=True)
            if st.button(icon, key=label, help=label, type="secondary"):
                st.session_state["secao"] = target
                st.experimental_rerun()
            st.markdown(f"<div class='icon-caption'>{label}</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: icon_card("🗂️", "Cadastro de Vistorias", "Cadastro")
    with c2: icon_card("📊", "Dashboard Operacional", "Dashboard")
    with c3: icon_card("📑", "Relatórios", "Relatorios")
    with c4: icon_card("🔄", "Status / Andamento", "Status")
    with c5: icon_card("🕵️", "Auditoria", "Auditoria")
    st.stop()  # impede que o restante do app carregue antes da seleção

# --------------------------------------------------------
# Dispatcher principal (após escolher no menu)
# --------------------------------------------------------
st.sidebar.title("📂 Navegação")
menu = st.sidebar.radio(
    "Selecione uma seção:",
    (
        "🏠 Início",
        "Cadastro de Vistorias",
        "Dashboard Operacional",
        "Relatórios",
        "Status / Andamento",
        "Auditoria"
    ),
)

if menu == "🏠 Início":
    st.session_state["secao"] = "Home"
    st.experimental_rerun()

def _run_page(module, label: str):
    for fn_name in ("page", "main", "app"):
        fn = getattr(module, fn_name, None)
        if callable(fn):
            try:
                return fn()
            except Exception as e:
                st.error(f"Erro ao executar `{fn_name}` no módulo **{label}**.\n\nDetalhes: `{e}`")
                return
    st.error(f"O módulo **{label}** não expõe uma função `page()`, `main()` ou `app()`.")

if st.session_state["secao"] == "Cadastro":
    st.markdown("### 📝 Cadastro de Vistorias")
    _run_page(cadastro, "features/cadastro.py")
elif st.session_state["secao"] == "Dashboard":
    st.markdown("### 📊 Dashboard Operacional — CRO/1 (Vistorias Técnicas)")
    _run_page(dashboard, "features/dashboard.py")
elif st.session_state["secao"] == "Relatorios":
    st.markdown("### 📑 Geração de Relatórios")
    _run_page(relatorios, "features/relatorios.py")
elif st.session_state["secao"] == "Status":
    st.markdown("### 📊 Status e Andamento das Solicitações")
    _run_page(status, "features/status.py")
elif st.session_state["secao"] == "Auditoria":
    st.markdown("### 🔎 Registro de Auditoria")
    _run_page(audit, "core/audit.py")
