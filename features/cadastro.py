import streamlit as st
from datetime import datetime
from core.models import SolicitacaoVistoria
from core.data_loader import append_row, read_df
from core.config import has_gsheets

def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    with st.form("cadastro_vistoria", clear_on_submit=True):
        om = st.text_input("OM solicitante *")
        local = st.text_input("Local / Instalação *")
        tipo = st.selectbox("Tipo de vistoria *", ["Periódica","Emergencial","Preventiva","Extraordinária"])
        motivo = st.text_area("Motivo / Justificativa *")
        urg = st.selectbox("Urgência", ["Não Prioritário","Prioritário","Urgente"])
        data_limite = st.date_input("Data limite")
        enviar = st.form_submit_button("Salvar")

    if not enviar:
        return

    # validações simples
    if not om or not local or not tipo or not motivo:
        st.error("Preencha todos os campos obrigatórios (*)")
        return

    # monta payload
    row = {
        "OM APOIADA": om,
        "LOCAL": local,
        "TIPO": tipo,
        "MOTIVO": motivo,
        "URGENCIA": urg,
        "DATA DA SOLICITACAO": datetime.now().strftime("%Y-%m-%d"),
        "DATA LIMITE": data_limite.strftime("%Y-%m-%d") if data_limite else "",
        "Situacao": "SOLICITADA",
    }

    # aba de destino
    tab_dest = st.session_state["tabs_map"]["solicitacoes"]
    headers = [
        "OBJETO DE VISTORIA","OM APOIADA","Diretoria Responsavel",
        "Classificacao da Urgencia","Situacao","DATA DA SOLICITACAO",
        "LOCAL","TIPO","MOTIVO","URGENCIA","DATA LIMITE"
    ]

    try:
        append_row(tab_dest, headers, row)      # grava no Sheets
        read_df.clear()                         # limpa cache de leitura
        st.success("Solicitação registrada no Google Sheets ✅")
    except Exception as e:
        st.error(f"Falha ao salvar no Sheets: {e}")
