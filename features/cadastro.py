# -*- coding: utf-8 -*-
import streamlit as st
from datetime import date
import pandas as pd

from core.data_loader import append_row, read_df
from core.config import TAB_SOLICITACOES, TAB_VALIDACAO


def _clean(x): return "" if pd.isna(x) else str(x).strip()


def _load_oms_map() -> tuple[list[str], dict[str, str], dict[str, str]]:
    """options, display->sigla, sigla->diretoria"""
    options, disp2sig, sig2dir = [], {}, {}

    for tab in (TAB_VALIDACAO, TAB_SOLICITACOES):
        try:
            df = read_df(tab)
        except Exception:
            continue
        if df.empty:
            continue

        cols = {c.lower(): c for c in df.columns}
        c_sig = cols.get("om") or cols.get("om apoiada") or cols.get("sigla")
        c_nom = cols.get("organização militar") or cols.get("organização") or cols.get("om")
        c_dir = cols.get("diretoria responsável") or cols.get("diretoria")
        if not c_sig or not c_dir:
            continue

        tmp = pd.DataFrame({"sig": df[c_sig].map(_clean)})
        tmp["nome"] = df[c_nom].map(_clean) if c_nom else ""
        tmp["dir"] = df[c_dir].map(_clean)
        tmp = tmp[tmp["sig"] != ""].drop_duplicates("sig")

        for _, r in tmp.iterrows():
            label = f"{r['sig']} — {r['nome']}" if r["nome"] else r["sig"]
            if label not in options:
                options.append(label)
                disp2sig[label] = r["sig"]
                sig2dir[r["sig"]] = r["dir"]
        break

    options.append("Outra / não listada…")
    disp2sig["Outra / não listada…"] = ""
    return options, disp2sig, sig2dir



def _input_row(oms_df: pd.DataFrame | None = None):
    """Coleta os campos do formulário e devolve (row, ok). 
       Se `oms_df` for passado, usa-o para autocomplete de OM e preenchimento de Diretoria."""
    st.subheader("📥 Nova solicitação de vistoria")

    # --- Autocomplete de OM + diretoria (se você já montou o dicionário em outro lugar, pode remover esta parte) ---
    display_to_sigla: dict[str, str] = {}
    sigla_to_dir: dict[str, str] = {}
    if isinstance(oms_df, pd.DataFrame) and not oms_df.empty:
        for _, r in oms_df.iterrows():
            sig = str(r.get("om_sigla", "") or "").strip()
            nom = str(r.get("om_nome", "") or "").strip()
            dir_ = str(r.get("diretoria", "") or "").strip()
            if not sig:
                continue
            display = f"{sig} — {nom}" if nom else sig
            display_to_sigla[display] = sig
            sigla_to_dir[sig] = dir_

    col1, col2 = st.columns(2)
    with col1:
        # OM: se temos base, usa autocomplete; senão, text_input
        if display_to_sigla:
            om_display = st.selectbox("OM solicitante (sigla)", [""] + sorted(display_to_sigla.keys()))
            om_solicitante = display_to_sigla.get(om_display, "")
        else:
            om_solicitante = st.text_input("OM solicitante (sigla)").strip()

        # Diretoria: se mapeada pela OM, preenche e desabilita
        diretoria_auto = sigla_to_dir.get(om_solicitante, "")
        if diretoria_auto:
            diretoria = st.text_input("Diretoria responsável", value=diretoria_auto, disabled=True)
        else:
            diretoria = st.text_input("Diretoria responsável").strip()

        tipo_vistoria = st.selectbox(
            "Tipo de vistoria",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
        )

    with col2:
        local = st.text_input("Local / instalação").strip()
        urgencia = st.selectbox("Urgência", ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"], index=0)
        data_limite = st.date_input("Data limite (se houver)", value=None)

    motivo = st.text_area("Motivo / justificativa (NAOM)", height=120)

    # --- Validações (sem unsafe_allow_html) ---
    erros = []
    if not om_solicitante:
        erros.append("Informe a **OM solicitante**.")
    if not local:
        erros.append("Informe o **local/instalação**.")
    if not (motivo or "").strip():
        erros.append("Descreva o **motivo/justificativa**.")

    if erros:
        # avisar sem HTML; se quiser bullets bonitas, use st.markdown com unsafe_allow_html=True
        st.warning("• " + "\n• ".join(erros))

    row = {
        "OBJETO DE VISTORIA": (motivo or "").strip(),  # ajuste se quiser outro campo como “objeto”
        "OM APOIADA": om_solicitante,
        "Diretoria Responsável": diretoria,
        "Classificação de Urgência": urgencia,
        "Situação": "SOLICITADA",  # estado inicial
        "DATA DA SOLICITAÇÃO": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "DATA DA SOLICITAÇÃO_2": data_limite.strftime("%Y-%m-%d") if data_limite else "",
        "REFERÊNCIA OPUS": "",
        "OBJETIVO (ADICIONAR POSSÍVEL CONTATO)": (motivo or "").strip(),
        "DATA DA VISTORIA": "",
        "VT EXECUTADA POR": "",
        "STATUS - ATUALIZAÇÃO SEMANAL": "",
        "DATA/PREVISÃO DE CONCLUSÃO": "",
        "MEIO DE RESPOSTA DA SOLICITAÇÃO": "",
        "DATA DA RESPOSTA A SOLICITAÇÃO": "",
        "Nº OPUS DA VISTORIA (SE FOR O CASO)": "",
        "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO": "",
        "QUANTIDADE DE DIAS PARA EXECUÇÃO": "",
        "OBSERVAÇÕES": "",
    }
    ok = (len(erros) == 0)
    return row, ok


def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    tabs_map = st.session_state.get("tabs_map", {})
    tab_base = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")   # fonte de dados/dash
    tab_valid = tabs_map.get("validacao", "Validacao_de_Dados")           # OMs oficiais
    tab_save  = "ACOMPANHAMENTO VISTORIAS"                                 # destino do salvamento

    # Carrega OMs para autocomplete (se existir)
    try:
        df_valid = read_df(tab_valid)
    except Exception:
        df_valid = pd.DataFrame()

    # Normaliza nomes de colunas esperados para o autocomplete (adeque à sua aba de validação)
    oms_df = pd.DataFrame()
    if not df_valid.empty:
        # tente mapear as colunas da validação (ajuste as opções conforme o seu sheet)
        def pick(df, *cands):
            for c in cands:
                if c in df.columns:
                    return c
            # tentativa por contains
            up = [c.upper() for c in df.columns]
            for c in cands:
                for i, u in enumerate(up):
                    if c.upper() in u:
                        return df.columns[i]
            return None

        c_sig = pick(df_valid, "OM", "Sigla", "OM Sigla")
        c_nom = pick(df_valid, "Organização Militar", "OM Nome", "Nome")
        c_dir = pick(df_valid, "Diretoria Responsável", "Diretoria", "DIR")

        if c_sig:
            oms_df = pd.DataFrame({
                "om_sigla": df_valid[c_sig].astype(str),
                "om_nome" : df_valid[c_nom].astype(str) if c_nom else "",
                "diretoria": df_valid[c_dir].astype(str) if c_dir else "",
            })

    # --- Formulário ---
    row, ok = _input_row(oms_df if not oms_df.empty else None)

    # --- Salvar ---
    if st.button("💾 Salvar na aba ACOMPANHAMENTO VISTORIAS", disabled=not ok):
        try:
            # limpa NaN -> "", e garante str
            clean_row = {k: ("" if pd.isna(v) else str(v)) for k, v in row.items()}
            append_row(tab_save, clean_row)
            st.success("Registro salvo com sucesso na aba **ACOMPANHAMENTO VISTORIAS**.")
            st.rerun()
        except Exception as e:
            st.error(f"Falha ao salvar: {e}")
