# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime
import pandas as pd

from core.data_loader import append_row, read_df
from core.config import TAB_SOLICITACOES, TAB_VALIDACAO

# -----------------------------------------------------------------------------
# Utilidades
# -----------------------------------------------------------------------------

def _clean(x) -> str:
    return "" if pd.isna(x) else str(x).strip()

def _pick(df: pd.DataFrame, *cands: str) -> str | None:
    """Encontra a coluna de df a partir de uma lista de candidatos (case/contains)."""
    if not isinstance(df, pd.DataFrame):
        return None
    cols = list(df.columns)
    up = [c.upper().strip() for c in cols]
    # match exato
    for c in cands:
        cu = c.upper().strip()
        if cu in up:
            return cols[up.index(cu)]
    # match por contains
    for c in cands:
        cu = c.upper().strip()
        for i, u in enumerate(up):
            if cu in u:
                return cols[i]
    return None

def _build_om_catalog() -> tuple[list[str], dict[str, str], dict[str, str]]:
    """
    Lê a aba de validação e (se estiver vazia) tenta a de solicitações,
    e monta:
      - options: lista para o select (ex.: "1º BPE — 1º Batalhão ...")
      - disp2sig: mapeia o rótulo escolhido -> sigla (OM)
      - sig2dir: mapeia sigla (OM) -> diretoria
    """
    options: list[str] = []
    disp2sig: dict[str, str] = {}
    sig2dir: dict[str, str] = {}

    for tab in (TAB_VALIDACAO, TAB_SOLICITACOES):
        try:
            df = read_df(tab)
        except Exception:
            continue
        if df is None or df.empty:
            continue

        c_sig = _pick(df, "OM", "OM APOIADA", "SIGLA", "OM SIGLA")
        c_nom = _pick(df, "Organização Militar", "OM NOME", "NOME")
        c_dir = _pick(df, "Diretoria Responsável", "Diretoria", "DIR")
        if not c_sig or not c_dir:
            continue

        tmp = pd.DataFrame({
            "sig": df[c_sig].map(_clean),
            "nome": df[c_nom].map(_clean) if c_nom else "",
            "dir":  df[c_dir].map(_clean),
        })
        tmp = tmp[tmp["sig"] != ""].drop_duplicates("sig")

        for _, r in tmp.iterrows():
            label = f"{r['sig']} — {r['nome']}" if r["nome"] else r["sig"]
            if label not in options:
                options.append(label)
                disp2sig[label] = r["sig"]
                sig2dir[r["sig"]] = r["dir"]
        break  # encontrou uma aba válida; para aqui

    # opção livre
    options = sorted(options)
    options.append("Outra / não listada…")
    disp2sig["Outra / não listada…"] = ""
    return options, disp2sig, sig2dir

# -----------------------------------------------------------------------------
# Formulário
# -----------------------------------------------------------------------------

def _input_row(om_options: list[str], disp2sig: dict[str, str], sig2dir: dict[str, str]):
    """Coleta os campos do formulário e devolve (row, ok)."""
    st.subheader("📥 Nova solicitação de vistoria")

    col1, col2 = st.columns(2)

    with col1:
        # Select de OMs (ou livre, se escolher "Outra…")
        choice = st.selectbox(
            "OM solicitante (sigla)",
            options=om_options,
            index=None,
            placeholder="Selecione a OM…",
        )

        om_solicitante = ""
        if choice:
            if choice == "Outra / não listada…":
                om_solicitante = st.text_input("Informe a OM (sigla)")
            else:
                om_solicitante = disp2sig.get(choice, "")

        # Diretoria auto (se existir no catálogo), mas editável
        auto_dir = sig2dir.get(om_solicitante, "") if om_solicitante else ""
        diretoria = st.text_input("Diretoria responsável", value=auto_dir)

        tipo_vistoria = st.selectbox(
            "Tipo de vistoria",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
        )

    with col2:
        local = st.text_input("Local / instalação")
        urgencia = st.selectbox("Urgência", ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"], index=0)
        data_limite = st.date_input("Data limite (se houver)", value=None)

    motivo = st.text_area("Motivo / justificativa (NAOM)", height=120)

    # Validações
    erros = []
    if not (om_solicitante or "").strip():
        erros.append("Informe a **OM solicitante**.")
    if not local.strip():
        erros.append("Informe o **local/instalação**.")
    if not motivo.strip():
        erros.append("Descreva o **motivo/justificativa**.")

    if erros:
        st.warning("• " + "<br>• ".join(erros), unsafe_allow_html=True)

    # Linha no formato da aba ACOMPANHAMENTO VISTORIAS
    row = {
        "OBJETO DE VISTORIA": (motivo or "").strip(),
        "OM APOIADA": (om_solicitante or "").strip(),
        "Diretoria Responsável": (diretoria or "").strip(),
        "Classificação de Urgência": urgencia,
        "Situação": "SOLICITADA",
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

    return row, (len(erros) == 0)

# -----------------------------------------------------------------------------
# Página
# -----------------------------------------------------------------------------

def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    # Mapeamento de abas definido na sidebar do app
    tabs_map = st.session_state.get("tabs_map", {})
    # Fonte principal do dashboard (continua igual)
    tab_base = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")
    # Aba de validação oficial (OMs)
    tab_valid = tabs_map.get("validacao", "Validacao_de_Dados")
    # Destino do salvamento (seu requisito)
    tab_save = "ACOMPANHAMENTO VISTORIAS"

    # Carrega catálogo de OMs (com fallback automático)
    om_options, disp2sig, sig2dir = _build_om_catalog()

    # Formulário
    row, ok = _input_row(om_options, disp2sig, sig2dir)

    # Salvar
    if st.button("💾 Salvar na aba ACOMPANHAMENTO VISTORIAS", disabled=not ok):
        try:
            clean_row = {k: ("" if pd.isna(v) else str(v)) for k, v in row.items()}
            append_row(tab_save, clean_row)
            st.success("Registro salvo com sucesso na aba **ACOMPANHAMENTO VISTORIAS**.")
            st.rerun()
        except Exception as e:
            st.error(f"Falha ao salvar: {e}")
