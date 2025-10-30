# -*- coding: utf-8 -*-
"""
Home — CRO/1 Sistema de Vistorias
Página inicial com menu de ícones
"""

from __future__ import annotations
import streamlit as st

st.set_page_config(
    page_title="CRO/1 — Sistema de Vistorias",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
.block-container { padding-top: 2rem; padding-bottom: 3rem; text-align:center; }
.icon-btn button[kind="secondary"]{height:120px;width:120px;font-size:60px;border-radius:24px;}
.icon-caption{margin-top:8px;font-weight:600;font-size:.95rem;text-align:center;}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1>Navegação</h1>", unsafe_allow_html=True)
st.write("Clique em um ícone para abrir a seção")
st.write("")

def icon(link_path: str, emoji: str, label: str):
    st.markdown("<div class='icon-btn'>", unsafe_allow_html=True)
    if st.button(emoji, help=label, type="secondary", key=label):
        st.switch_page(link_path)
    st.markdown(f"<div class='icon-caption'>{label}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

c1,c2,c3,c4,c5 = st.columns(5, gap="large")
with c1: icon("pages/Cadastro_de_vistorias.py", "🗂️", "Cadastro de Vistorias")
with c2: icon("pages/Dashboard_operacional.py", "📊", "Dashboard Operacional")
with c3: icon("pages/Relatorios.py", "📑", "Relatórios")
with c4: icon("pages/Status_Andamento.py", "🔄", "Status / Andamento")
with c5: icon("pages/Auditoria.py", "🕵️", "Auditoria")
