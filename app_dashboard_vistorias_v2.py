# -*- coding: utf-8 -*-
"""
Aplicativo principal do Sistema de Vistorias CRO/1
Navegação: Cadastro | Dashboard Operacional | Relatórios | Status/Andamento | Auditoria
"""

from __future__ import annotations
import streamlit as st

# Importa as páginas (cada uma com função page())
from features import cadastro, status, relatorios, dashboard
from core import audit

# Config global da app (apenas aqui)
st.set_page_config(
    page_title="CRO/1 — Sistema de Vistorias",
    layout="wide",
    page_icon="🛠️",
)

st.sidebar.title("📂 Navegação")
menu = st.sidebar.radio(
    "Selecione uma seção:",
    (
        "Cadastro de Vistorias",
        "Dashboard Operacional",
        "Relatórios",
        "Status / Andamento",
        "Auditoria",
    ),
)

if menu == "Cadastro de Vistorias":
    st.markdown("### 📝 Cadastro de Vistorias")
    cadastro.page()

elif menu == "Dashboard Operacional":
    st.markdown("### 📊 Dashboard Operacional — CRO/1 (Vistorias Técnicas)")
    dashboard.page()

elif menu == "Relatórios":
    st.markdown("### 📑 Geração de Relatórios")
    relatorios.page()

elif menu == "Status / Andamento":
    st.markdown("### 📊 Status e Andamento das Solicitações")
    status.page()

elif menu == "Auditoria":
    st.markdown("### 🔎 Registro de Auditoria")
    audit.page()
