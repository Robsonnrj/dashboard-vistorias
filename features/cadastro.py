# features/cadastro.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import streamlit as st
from datetime import datetime, date
import pandas as pd

from core.data_loader import append_row, read_df
from core.config import TAB_VALIDACAO  # usamos apenas a aba de validação


# =========================================================
# Helpers
# =========================================================
def _clean(x) -> str:
    """Normaliza valores vindos do Sheets."""
    s = "" if pd.isna(x) else str(x).strip()
    if s.upper() in {"#N/A", "N/A", "NA", "NAN", "NONE", "-"}:
        return ""
    return s


def _build_om_catalog() -> tuple[list[str], dict[str, str], dict[str, str]]:
    """
    Lê a aba de validação (TAB_VALIDACAO) e monta:
      - options: rótulos para o select (ex.: "1º BPE — 1º Batalhão ...")
      - disp2sig: rótulo -> sigla
      - sig2dir : sigla  -> diretoria
    """
    try:
        # Lê a aba de validação com header na linha 1 (índice 1)
        df = read_df(TAB_VALIDACAO, header=1)
        
        # Se read_df não suportar o parâmetro header, use esta alternativa:
        # df = pd.read_excel("caminho_do_arquivo.xlsx", sheet_name=TAB_VALIDACAO, header=1)
        
    except Exception as e:
        st.warning(f"Erro ao carregar dados de validação: {e}")
        df = pd.DataFrame()

    options: list[str] = []
    disp2sig: dict[str, str] = {}
    sig2dir: dict[str, str] = {}

    if not df.empty:
        # As colunas estão como "Unnamed: X", então precisamos usar os índices
        # Coluna 1: OM (sigla)
        # Coluna 2: Organização Militar (nome completo)
        # Coluna 3: Diretoria Responsável
        
        if len(df.columns) >= 4:
            # Renomeia as colunas para facilitar
            col_sig = df.columns[1]   # "Unnamed: 1" -> OM (sigla)
            col_nome = df.columns[2]  # "Unnamed: 2" -> Organização Militar
            col_dir = df.columns[3]   # "Unnamed: 3" -> Diretoria Responsável
            
            # Cria dataframe limpo
            tmp = pd.DataFrame({
                "sig": df[col_sig].map(_clean),
                "nome": df[col_nome].map(_clean),
                "dir": df[col_dir].map(_clean),
            })
            
            # Remove linha de cabeçalho duplicada (NR, OM, Organização Militar...)
            tmp = tmp[tmp["sig"] != "OM"]
            tmp = tmp[tmp["sig"] != "NR"]
            tmp = tmp[tmp["sig"] != ""]
            
            # Remove duplicatas
            tmp = tmp.drop_duplicates("sig")
            
            # Cria as opções para o selectbox
            for _, r in tmp.iterrows():
                if r['sig']:  # Só adiciona se tiver sigla
                    label = f"{r['sig']} — {r['nome']}" if r['nome'] else r['sig']
                    options.append(label)
                    disp2sig[label] = r["sig"]
                    sig2dir[r["sig"]] = r["dir"] if r["dir"] else ""

    # Ordena alfabeticamente
    options = sorted(options)
    
    # Adiciona opção para OM não listada
    options.append("Outra / não listada…")
    disp2sig["Outra / não listada…"] = ""
    
    return options, disp2sig, sig2dir


# ---------------------------------------------------------
# Callback: ao trocar a OM, atualiza Diretoria e limpa OM manual
# ---------------------------------------------------------
def _on_om_change(disp2sig: dict[str, str], sig2dir: dict[str, str]):
    choice = st.session_state.get("om_choice")
    sig = disp2sig.get(choice or "", "")
    st.session_state["dir_resp"] = sig2dir.get(sig, "") if sig else ""
    st.session_state["om_sigla_out"] = ""


# ---------------------------------------------------------
# Chaves do formulário para limpeza automática
# ---------------------------------------------------------
FORM_KEYS = [
    "om_choice", "om_sigla_out", "dir_resp",
    "tipo_vist", "local_inst", "urg", "data_limite", "motivo",
    "ref_opus", "objetivo", "vt_exec", "status_sem", "obs",
    "data_vist", "data_prev_conc", "meio_resp", "data_resp",
    "num_opus", "qd_total", "qd_exec",
]


def _reset_form():
    for k in FORM_KEYS:
        if k in st.session_state:
            del st.session_state[k]


# =========================================================
# Formulário (UI)
# =========================================================
def _input_row(om_options: list[str], disp2sig: dict[str, str], sig2dir: dict[str, str]):
    st.subheader("📥 Nova solicitação de vistoria")

    c1, c2 = st.columns(2)

    # -------- coluna 1 --------
    with c1:
        choice = st.selectbox(
            "OM solicitante (sigla)",
            options=om_options,
            index=None,
            placeholder="Selecione a OM…",
            key="om_choice",
            on_change=_on_om_change,
            kwargs={"disp2sig": disp2sig, "sig2dir": sig2dir},
        )

        # Sigla selecionada (ou digitada)
        om_sigla = ""
        if choice:
            om_sigla = disp2sig.get(choice, "")
            if choice == "Outra / não listada…":
                om_sigla = st.text_input("Informe a OM (sigla)", key="om_sigla_out").strip()

        # Diretoria sincronizada automaticamente pelo callback
        # Torna o campo editável caso o usuário precise ajustar
        diretoria = st.text_input(
            "Diretoria responsável", 
            key="dir_resp",
            help="Preenchimento automático. Você pode editar se necessário."
        )

        tipo_vistoria = st.selectbox(
            "Tipo de vistoria",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
            key="tipo_vist",
        )

    # -------- coluna 2 --------
    with c2:
        local = st.text_input("Local / instalação", key="local_inst")
        urgencia = st.selectbox(
            "Urgência", ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"], index=0, key="urg"
        )
        data_limite: date | None = st.date_input(
            "Data limite (se houver)", value=None, format="YYYY/MM/DD", key="data_limite"
        )

    motivo = st.text_area("Motivo / justificativa (NAOM)", height=120, key="motivo")

    # ---------- Complementares ----------
    st.markdown("### Complementares")
    cc1, cc2 = st.columns(2)
    with cc1:
        referencia_opus = st.text_input("REFERÊNCIA OPUS", key="ref_opus")
        objetivo_contato = st.text_input("OBJETIVO (ADICIONAR POSSÍVEL CONTATO)", key="objetivo")
        vt_exec_por = st.text_input("VT EXECUTADA POR", key="vt_exec")
        status_atual = st.text_input("STATUS - ATUALIZAÇÃO SEMANAL", key="status_sem")
        obs = st.text_area("OBSERVAÇÕES", height=90, key="obs")

    with cc2:
        data_vistoria: date | None = st.date_input(
            "DATA DA VISTORIA", value=None, format="YYYY/MM/DD", key="data_vist"
        )
        data_prev_conc: date | None = st.date_input(
            "DATA/PREVISÃO DE CONCLUSÃO", value=None, format="YYYY/MM/DD", key="data_prev_conc"
        )
        meio_resposta = st.text_input("MEIO DE RESPOSTA DA SOLICITAÇÃO", key="meio_resp")
        data_resposta: date | None = st.date_input(
            "DATA DA RESPOSTA A SOLICITAÇÃO", value=None, format="YYYY/MM/DD", key="data_resp"
        )
        num_opus = st.text_input("Nº OPUS DA VISTORIA (SE FOR O CASO)", key="num_opus")
        qd_total = st.number_input(
            "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO", min_value=0, step=1, value=0, key="qd_total"
        )
        qd_exec = st.number_input(
            "QUANTIDADE DE DIAS PARA EXECUÇÃO", min_value=0, step=1, value=0, key="qd_exec"
        )

    # ---------- Validações ----------
    erros = []
    if not om_sigla:
        erros.append("Informe a OM solicitante.")
    if not local.strip():
        erros.append("Informe o local/instalação.")
    if not motivo.strip():
        erros.append("Descreva o motivo/justificativa (NAOM).")

    if erros:
        st.warning("• " + "\n• ".join(erros))

    # ---------- Monta linha conforme a aba ACOMPANHAMENTO VISTORIAS ----------
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
        "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO": (str(int(qd_total)) if qd_total else ""),
        "QUANTIDADE DE DIAS PARA EXECUÇÃO": (str(int(qd_exec)) if qd_exec else ""),
        "OBSERVAÇÕES": obs.strip(),
    }

    ok = (len(erros) == 0)
    return row, ok


# =========================================================
# Página
# =========================================================
def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    # Catálogo de OMs (sigla + diretoria automática) da aba Validacao_de_Dados
    om_options, disp2sig, sig2dir = _build_om_catalog()

    # Debug (opcional - remova em produção)
    if st.checkbox("🔍 Mostrar dados carregados (debug)", value=False):
        st.info(f"Total de OMs carregadas: {len(om_options) - 1}")  # -1 por causa de "Outra/não listada"
        with st.expander("Ver detalhes"):
            st.write("**Opções disponíveis:**", om_options[:10])
            st.write("**Mapeamento sigla -> diretoria:**", dict(list(sig2dir.items())[:10]))

    # Formulário
    row, ok = _input_row(om_options, disp2sig, sig2dir)

    st.divider()
    if st.button("💾 Salvar na aba ACOMPANHAMENTO VISTORIAS", type="primary", disabled=not ok):
        try:
            # normaliza NaN -> "" e tudo como str
            clean_row = {k: ("" if pd.isna(v) else str(v)) for k, v in row.items()}
            append_row("ACOMPANHAMENTO VISTORIAS", clean_row)
            st.success("✅ Registro salvo com sucesso na aba **ACOMPANHAMENTO VISTORIAS**.")
            # limpa o formulário para um novo cadastro
            _reset_form()
            st.rerun()
        except Exception as e:
            st.error(f"❌ Falha ao salvar: {e}")
