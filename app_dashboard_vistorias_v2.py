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

# Sidebar de navegação
st.sidebar.title("📂 Navegação")
menu = st.sidebar.radio(
    "Selecione uma seção:",
    (
        "Cadastro de Vistorias",
        "Dashboard Operacional",
        "Relatórios",
        "Status / Andamento",
        "Auditoria"
    ),
)

def _run_page(module, label: str):
    """
    Executa a página principal do módulo, garantindo fallback e feedback:
    - Tenta `page()`, depois `main()`, depois `app()`.
    - Caso não exista, exibe alerta para o desenvolvedor corrigir.
    """
    for fn_name in ("page", "main", "app"):
        fn = getattr(module, fn_name, None)
        if callable(fn):
            try:
                return fn()
            except Exception as e:
                # Feedback detalhado se alguma exception for disparada internamente
                st.error(
                    f"Erro ao executar `{fn_name}` no módulo **{label}**.\n"
                    f"\nDetalhes: `{e}`"
                )
                return
    st.error(
        f"O módulo **{label}** não expõe uma função `page()`, `main()` ou `app()`. "
        f"Abra `{getattr(module, '__file__', label)}` e defina uma dessas funções."
    )

# Dispatcher simples por menu
if menu == "Cadastro de Vistorias":
    st.markdown("### 📝 Cadastro de Vistorias")
    _run_page(cadastro, "features/cadastro.py")
elif menu == "Dashboard Operacional":
    st.markdown("### 📊 Dashboard Operacional — CRO/1 (Vistorias Técnicas)")
    _run_page(dashboard, "features/dashboard.py")
elif menu == "Relatórios":
    st.markdown("### 📑 Geração de Relatórios")
    _run_page(relatorios, "features/relatorios.py")
elif menu == "Status / Andamento":
    st.markdown("### 📊 Status e Andamento das Solicitações")
    _run_page(status, "features/status.py")
elif menu == "Auditoria":
    st.markdown("### 🔎 Registro de Auditoria")
    _run_page(audit, "core/audit.py")
