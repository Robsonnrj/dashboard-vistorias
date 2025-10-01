# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime, date
import pandas as pd

from core.data_loader import append_row, read_df
from core.config import TAB_SOLICITACOES, TAB_VALIDACAO


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def _clean(x) -> str:
    return "" if pd.isna(x) else str(x).strip()

def _pick(df: pd.DataFrame, *cands: str) -> str | None:
    if not isinstance(df, pd.DataFrame) or df.empty:
        return None
    cols = list(df.columns)
    up = [c.upper().strip() for c in cols]
    for c in cands:
        cu = c.upper().strip()
        if cu in up:
            return cols[up.index(cu)]
    for c in cands:
        cu = c.upper().strip()
        for i, u in enumerate(up):
            if cu in u:
                return cols[i]
    return None

def _build_om_catalog() -> tuple[list[str], dict[str, str], dict[str, str]]:
    """Cria catálogo de OMs com rótulo -> sigla e sigla -> diretoria."""
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
        break

    options = sorted(options)
    options.append("Outra / não listada…")
    disp2sig["Outra / não listada…"] = ""
    return options, disp2sig, sig2dir


# ------------------------------------------------------------------
# Formulário
# ------------------------------------------------------------------
def _input_row(om_options: list[str], disp2sig: dict[str, str], sig2dir: dict[str, str]):
    st.subheader("📥 Nova solicitação de vistoria")

    # Básico
    c1, c2 = st.columns(2)

    with c1:
        choice = st.selectbox(
            "OM solicitante (sigla)",
            options=om_options,
            index=None,
            placeholder="Selecione a OM…",
        )
        om_sigla = ""
        if choice:
            om_sigla = disp2sig.get(choice, "")
            if choice == "Outra / não listada…":
                om_sigla = st.text_input("Informe a OM (sigla)").strip()

        auto_dir = sig2dir.get(om_sigla, "") if om_sigla else ""
        diretoria = st.text_input("Diretoria responsável", value=auto_dir)

        tipo_vistoria = st.selectbox(
            "Tipo de vistoria",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
        )

    with c2:
        local = st.text_input("Local / instalação")
        urgencia = st.selectbox(
            "Urgência", ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"], index=0
        )
        data_limite: date | None = st.date_input(
            "Data limite (se houver)", value=None, format="YYYY/MM/DD"
        )

    motivo = st.text_area("Motivo / justificativa (NAOM)", height=120)

    # Complementares (todos os campos que você queria)
    st.markdown("### Complementares")
    cc1, cc2 = st.columns(2)
    with cc1:
        referencia_opus = st.text_input("REFERÊNCIA OPUS")
        objetivo_contato = st.text_input("OBJETIVO (ADICIONAR POSSÍVEL CONTATO)")
        vt_exec_por = st.text_input("VT EXECUTADA POR")
        status_atual = st.text_input("STATUS - ATUALIZAÇÃO SEMANAL")
        obs = st.text_area("OBSERVAÇÕES", height=90)

    with cc2:
        data_vistoria: date | None = st.date_input("DATA DA VISTORIA", value=None, format="YYYY/MM/DD")
        data_prev_conc: date | None = st.date_input("DATA/PREVISÃO DE CONCLUSÃO", value=None, format="YYYY/MM/DD")
        meio_resposta = st.text_input("MEIO DE RESPOSTA DA SOLICITAÇÃO")
        data_resposta: date | None = st.date_input("DATA DA RESPOSTA A SOLICITAÇÃO", value=None, format="YYYY/MM/DD")
        num_opus = st.text_input("Nº OPUS DA VISTORIA (SE FOR O CASO)")
        qd_total = st.number_input("QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO", min_value=0, step=1, value=0)
        qd_exec  = st.number_input("QUANTIDADE DE DIAS PARA EXECUÇÃO", min_value=0, step=1, value=0)

    # Validações
    erros = []
    if not om_sigla:
        erros.append("Informe a **OM solicitante**.")
    if not local.strip():
        erros.append("Informe o **local/instalação**.")
    if not motivo.strip():
        erros.append("Descreva o **motivo/justificativa**.")

    if erros:
        st.markdown(
            "⚠️ " + "<br>• ".join([""] + erros),
            unsafe_allow_html=True,
        )

    # Monta linha exatamente com os nomes da aba ACOMPANHAMENTO VISTORIAS
    row = {
        "OBJETO DE VISTORIA": motivo.strip(),
        "OM APOIADA": om_sigla,
        "Diretoria Responsável": diretoria.strip(),
        "Classificação de Urgência": urgencia,
        "Situação": "SOLICITADA",
        "DATA DA SOLICITAÇÃO": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "DATA DA SOLICITAÇÃO_2": data_limite.strftime("%Y-%m-%d") if data_limite else "",
        "REFERÊNCIA OPUS": referencia_opus.strip(),
        "OBJETIVO (ADICIONAR POSSÍVEL CONTATO)": objetivo_contato.strip(),
        "DATA DA VISTORIA": data_vistoria.strftime("%Y-%m-%d") if data_vistoria else "",
        "VT EXECUTADA POR": vt_exec_por.strip(),
        "STATUS - ATUALIZAÇÃO SEMANAL": status_atual.strip(),
        "DATA/PREVISÃO DE CONCLUSÃO": data_prev_conc.strftime("%Y-%m-%d") if data_prev_conc else "",
        "MEIO DE RESPOSTA DA SOLICITAÇÃO": meio_resposta.strip(),
        "DATA DA RESPOSTA A SOLICITAÇÃO": data_resposta.strftime("%Y-%m-%d") if data_resposta else "",
        "Nº OPUS DA VISTORIA (SE FOR O CASO)": num_opus.strip(),
        "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO": str(int(qd_total)) if qd_total else "",
        "QUANTIDADE DE DIAS PARA EXECUÇÃO": str(int(qd_exec)) if qd_exec else "",
        "OBSERVAÇÕES": obs.strip(),
    }

    ok = (len(erros) == 0)
    return row, ok


# ------------------------------------------------------------------
# Página
# ------------------------------------------------------------------
def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    # Abas do app (sidebar)
    tabs_map = st.session_state.get("tabs_map", {})
    # Mantém compatibilidade, mas o salvamento é na ACOMPANHAMENTO VISTORIAS
    _ = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")
    _ = tabs_map.get("validacao", "Validacao_de_Dados")
    tab_save = "ACOMPANHAMENTO VISTORIAS"

    # Catálogo de OMs (com diretoria automática)
    om_options, disp2sig, sig2dir = _build_om_catalog()

    # Form
    row, ok = _input_row(om_options, disp2sig, sig2dir)

    st.divider()
    # Botão de salvar SEMPRE aparece (só fica desabilitado se inválido)
    if st.button("💾 Salvar na aba ACOMPANHAMENTO VISTORIAS", type="primary", disabled=not ok):
        try:
            clean_row = {k: ("" if pd.isna(v) else str(v)) for k, v in row.items()}
            append_row(tab_save, clean_row)
            st.success("Registro salvo com sucesso na aba **ACOMPANHAMENTO VISTORIAS**.")
            st.rerun()
        except Exception as e:
            st.error(f"Falha ao salvar: {e}")
