# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime
from core.models import SolicitacaoVistoria
from core.data_loader import append_row
from core.config import has_gsheets

def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    with st.form("form_solicitacao", clear_on_submit=True):
        col1, col2, col3 = st.columns([1,2,1.2])
        with col1:
            om_solicitante = st.text_input("OM (sigla)*", placeholder="Ex.: 1º BPE")
        with col2:
            om_nome = st.text_input("Organização Militar (nome completo)*")
        with col3:
            diretoria = st.text_input("Diretoria*")

        local = st.text_input("Local/Instalação*")
        coordenadas = st.text_input("Coordenadas (lat,lon)", placeholder="-22.90,-43.17")

        col4, col5, col6 = st.columns(3)
        with col4:
            tipo_vistoria = st.selectbox("Tipo de Vistoria*", ["Periódica","Emergencial","Preventiva","Extraordinária"])
        with col5:
            urgencia = st.selectbox("Urgência*", ["NÃO PRIORITÁRIO","PRIORIDADE","URGENTE"])
        with col6:
            data_limite = st.date_input("Data limite*", value=datetime.today()).isoformat()

        motivo = st.text_area("Motivo/Justificativa técnica*", height=120)
        anexos = st.text_input("Referência/Anexos (URL DIEx/Drive)")
        criado_por = st.text_input("Responsável pelo cadastro", value="usuario")

        numero = st.text_input("Número (NAOM)", placeholder="NAOM/2025-0001")
        if not numero:
            numero = f"NAOM/{datetime.now().year}-{datetime.now().strftime('%m%d%H%M%S')}"

        enviar = st.form_submit_button("💾 Registrar Solicitação")

    if enviar:
        obrigatorios = [om_solicitante, om_nome, diretoria, local, tipo_vistoria, urgencia, data_limite, motivo]
        if not all(obrigatorios):
            st.error("Preencha todos os campos obrigatórios (*).")
            return
        s = SolicitacaoVistoria(
            numero=numero,
            om_solicitante=om_solicitante.strip(),
            om_nome=om_nome.strip(),
            diretoria=diretoria.strip(),
            local=local.strip(),
            coordenadas=coordenadas.strip(),
            tipo_vistoria=tipo_vistoria,
            motivo=motivo.strip(),
            urgencia=urgencia,
            data_limite=data_limite,
            anexos=anexos.strip(),
            criado_por=criado_por.strip()
        )
        if not has_gsheets():
            st.error("Google Sheets OFF. Configure .streamlit/secrets.toml")
            return
        append_row("solicitacoes", s.to_row())
        st.success(f"Solicitação **{s.numero}** registrada.")
