# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime, date
import pandas as pd

from core.data_loader import read_df, append_row
from core.config import TAB_SOLICITACOES, TAB_VALIDACAO

SHEET_COLUMNS = [
    "OBJETO DE VISTORIA",
    "OM APOIADA",
    "Diretoria Responsável",
    "Classificação de Urgência",
    "Situação",
    "DATA DA SOLICITAÇÃO",
    "DATA DA SOLICITAÇÃO_2",
    "REFERÊNCIA OPUS",
    "OBJETIVO (ADICIONAR POSSÍVEL CONTATO)",
    "DATA DA VISTORIA",
    "VT EXECUTADA POR",
    "STATUS - ATUALIZAÇÃO SEMANAL",
    "DATA/PREVISÃO DE CONCLUSÃO",
    "MEIO DE RESPOSTA DA SOLICITAÇÃO",
    "DATA DA RESPOSTA A SOLICITAÇÃO",
    "Nº OPUS DA VISTORIA (SE FOR O CASO)",
    "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO",
    "QUANTIDADE DE DIAS PARA EXECUÇÃO",
    "OBSERVAÇÕES",
]

URGENCIAS = ["Não Prioridade", "Prioridade", "Urgente"]
SITUACOES = ["Não Atendida", "Em andamento", "Finalizada"]

def _clean_str(x) -> str:
    return (str(x).strip() if x is not None else "").strip()

def _date_to_str(d: date | None) -> str:
    if not d:
        return ""
    try:
        return pd.to_datetime(d).strftime("%Y-%m-%d")
    except Exception:
        return ""

def _load_oms_map() -> tuple[list[str], dict, dict]:
    """Carrega OMs e diretorias de Validacao_de_Dados ou Acompanhamento Vistorias."""
    sources = [TAB_VALIDACAO, TAB_SOLICITACOES]
    oms_df = pd.DataFrame()
    for tab in sources:
        try:
            df = read_df(tab)
        except Exception:
            continue
        if df.empty:
            continue
        # normaliza
        dfc = {c.lower(): c for c in df.columns}
        sigla = dfc.get("om") or dfc.get("om apoiada") or dfc.get("sigla")
        nome = dfc.get("organização militar") or dfc.get("organização") or dfc.get("om")
        diretoria = dfc.get("diretoria responsável") or dfc.get("diretoria")
        if sigla and diretoria:
            oms_df = df[[sigla]].rename(columns={sigla:"om_sigla"})
            oms_df["om_nome"] = df[nome] if nome else ""
            oms_df["diretoria"] = df[diretoria]
            break

    options, disp2sig, sig2dir = [], {}, {}
    if not oms_df.empty:
        for _, r in oms_df.iterrows():
            sig = _clean_str(r.get("om_sigla"))
            nom = _clean_str(r.get("om_nome"))
            dire = _clean_str(r.get("diretoria"))
            if not sig:
                continue
            label = f"{sig} — {nom}" if nom else sig
            options.append(label)
            disp2sig[label] = sig
            if sig not in sig2dir:
                sig2dir[sig] = dire
    options.append("Outra / não listada…")
    disp2sig["Outra / não listada…"] = ""
    return options, disp2sig, sig2dir

def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    options, disp2sig, sig2dir = _load_oms_map()

    with st.form("form_vistoria"):
        objeto = st.text_input("OBJETO DE VISTORIA *")
        om_display = st.selectbox("OM APOIADA *", options, index=None, placeholder="Selecione…")
        om_sigla = disp2sig.get(om_display or "", "")
        diretoria = st.text_input("Diretoria Responsável *", value=sig2dir.get(om_sigla, ""))
        urg = st.selectbox("Classificação de Urgência *", URGENCIAS, index=0)
        sit = st.selectbox("Situação *", SITUACOES, index=0)
        data_solic = st.date_input("DATA DA SOLICITAÇÃO", value=date.today())
        data_solic2 = st.date_input("DATA DA SOLICITAÇÃO_2", value=None)
        ref_opus = st.text_input("REFERÊNCIA OPUS")
        objetivo = st.text_input("OBJETIVO (ADICIONAR POSSÍVEL CONTATO)")
        data_vist = st.date_input("DATA DA VISTORIA", value=None)
        vt_exec = st.text_input("VT EXECUTADA POR")
        status_semana = st.text_input("STATUS - ATUALIZAÇÃO SEMANAL")
        previsao_conc = st.date_input("DATA/PREVISÃO DE CONCLUSÃO", value=None)
        meio_resp = st.text_input("MEIO DE RESPOSTA DA SOLICITAÇÃO")
        data_resp = st.date_input("DATA DA RESPOSTA A SOLICITAÇÃO", value=None)
        num_opus = st.text_input("Nº OPUS DA VISTORIA (SE FOR O CASO)")
        qt_total = st.number_input("QTD DIAS TOTAL ATENDIMENTO", min_value=0, step=1)
        qt_exec = st.number_input("QTD DIAS EXECUÇÃO", min_value=0, step=1)
        obs = st.text_area("OBSERVAÇÕES", height=100)

        salvar = st.form_submit_button("💾 Salvar", type="primary")

    if salvar:
        erros = []
        if not _clean_str(objeto): erros.append("OBJETO DE VISTORIA")
        if not _clean_str(om_sigla): erros.append("OM APOIADA")
        if not _clean_str(diretoria): erros.append("Diretoria Responsável")
        if erros:
            st.error("Preencha os campos obrigatórios: " + ", ".join(erros))
            return

        row = {
            "OBJETO DE VISTORIA": _clean_str(objeto),
            "OM APOIADA": _clean_str(om_sigla),
            "Diretoria Responsável": _clean_str(diretoria),
            "Classificação de Urgência": _clean_str(urg),
            "Situação": _clean_str(sit),
            "DATA DA SOLICITAÇÃO": _date_to_str(data_solic),
            "DATA DA SOLICITAÇÃO_2": _date_to_str(data_solic2),
            "REFERÊNCIA OPUS": _clean_str(ref_opus),
            "OBJETIVO (ADICIONAR POSSÍVEL CONTATO)": _clean_str(objetivo),
            "DATA DA VISTORIA": _date_to_str(data_vist),
            "VT EXECUTADA POR": _clean_str(vt_exec),
            "STATUS - ATUALIZAÇÃO SEMANAL": _clean_str(status_semana),
            "DATA/PREVISÃO DE CONCLUSÃO": _date_to_str(previsao_conc),
            "MEIO DE RESPOSTA DA SOLICITAÇÃO": _clean_str(meio_resp),
            "DATA DA RESPOSTA A SOLICITAÇÃO": _date_to_str(data_resp),
            "Nº OPUS DA VISTORIA (SE FOR O CASO)": _clean_str(num_opus),
            "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO": int(qt_total),
            "QUANTIDADE DE DIAS PARA EXECUÇÃO": int(qt_exec),
            "OBSERVAÇÕES": _clean_str(obs),
        }

        payload = {col: row.get(col, "") for col in SHEET_COLUMNS}
        try:
            append_row(TAB_SOLICITACOES, payload)
            st.success("✅ Registro salvo na aba ACOMPANHAMENTO VISTORIAS")
            st.rerun()
        except Exception as e:
            st.error(f"Falha ao salvar: {e}")
