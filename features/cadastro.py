# features/cadastro.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import streamlit as st
from datetime import datetime, date
import pandas as pd
from typing import Tuple, Dict, List

from core.data_loader import append_row, read_df
from core.config import TAB_VALIDACAO, TAB_SOLICITACOES


# =========================================================
# Funções Helper
# =========================================================
def _clean(x) -> str:
    """Normaliza valores vindos do Sheets."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.upper() in {"#N/A", "N/A", "NA", "NAN", "NONE", "-", ""}:
        return ""
    return s


@st.cache_data(ttl=600, show_spinner=False)  # Cache de 10 minutos
def _build_om_catalog() -> Tuple[List[str], Dict[str, str], Dict[str, str]]:
    """
    Lê a aba de validação e monta:
      - options: rótulos para o select (ex.: "1º BPE — 1º Batalhão ...")
      - disp2sig: rótulo -> sigla
      - sig2dir : sigla  -> diretoria
    """
    options: List[str] = []
    disp2sig: Dict[str, str] = {}
    sig2dir: Dict[str, str] = {}
    
    try:
        # Lê os dados sem processar header ainda
        df_raw = read_df(TAB_VALIDACAO, use_cache=True)
        
        if df_raw.empty:
            st.warning("⚠️ Aba de validação está vazia")
            return _add_fallback_option(options, disp2sig)
        
        # Procura pela linha que contém os cabeçalhos reais
        header_row = _find_header_row(df_raw)
        
        if header_row is None:
            st.warning("⚠️ Não foi possível localizar o cabeçalho na aba de validação")
            return _add_fallback_option(options, disp2sig)
        
        # Pega os dados a partir da linha seguinte ao header
        df = df_raw.iloc[header_row + 1:].reset_index(drop=True)
        
        # Define nomes únicos para as colunas
        header_values = df_raw.iloc[header_row].tolist()
        col_names = _create_unique_column_names(header_values)
        df.columns = col_names
        
        # Identifica as colunas importantes
        col_sig, col_nome, col_dir = _identify_columns(col_names)
        
        if not col_sig or len(df) == 0:
            st.warning("⚠️ Nenhuma coluna de OM identificada ou DataFrame vazio")
            return _add_fallback_option(options, disp2sig)
        
        # Cria dataframe limpo
        tmp = _create_clean_dataframe(df, col_sig, col_nome, col_dir)
        
        # Cria as opções para o selectbox
        for _, r in tmp.iterrows():
            if r['sig']:
                label = f"{r['sig']} — {r['nome']}" if r['nome'] else r['sig']
                options.append(label)
                disp2sig[label] = r["sig"]
                sig2dir[r["sig"]] = r["dir"] if r["dir"] else ""
        
        options = sorted(options)
        
    except Exception as e:
        st.error(f"❌ Erro ao carregar dados de validação: {e}")
        import traceback
        with st.expander("🔍 Detalhes do erro"):
            st.code(traceback.format_exc(), language="python")
    
    return _add_fallback_option(options, disp2sig)


def _find_header_row(df: pd.DataFrame) -> int | None:
    """Procura pela linha que contém os cabeçalhos."""
    for idx in range(min(5, len(df))):
        row_values = df.iloc[idx].astype(str).str.upper().tolist()
        if "OM" in row_values and any("ORGANIZA" in str(v) for v in row_values):
            return idx
    return None


def _create_unique_column_names(header_values: list) -> list:
    """Cria nomes únicos para colunas duplicadas."""
    col_names = []
    col_counts = {}
    
    for val in header_values:
        val_str = _clean(str(val))
        if not val_str or val_str == "<NA>":
            val_str = f"col_{len(col_names)}"
        
        if val_str in col_counts:
            col_counts[val_str] += 1
            val_str = f"{val_str}_{col_counts[val_str]}"
        else:
            col_counts[val_str] = 0
        
        col_names.append(val_str)
    
    return col_names


def _identify_columns(col_names: list) -> Tuple[str | None, str | None, str | None]:
    """Identifica as colunas importantes (sigla, nome, diretoria)."""
    col_sig = None
    col_nome = None
    col_dir = None
    
    for col in col_names:
        col_upper = col.upper()
        if col_sig is None and ("OM" == col_upper or "SIGLA" in col_upper):
            col_sig = col
        if col_nome is None and "ORGANIZA" in col_upper:
            col_nome = col
        if col_dir is None and "DIRETORIA" in col_upper:
            col_dir = col
    
    return col_sig, col_nome, col_dir


def _create_clean_dataframe(
    df: pd.DataFrame,
    col_sig: str,
    col_nome: str | None,
    col_dir: str | None
) -> pd.DataFrame:
    """Cria dataframe limpo removendo linhas inválidas."""
    tmp = pd.DataFrame({
        "sig": df[col_sig].apply(_clean),
        "nome": df[col_nome].apply(_clean) if col_nome else "",
        "dir": df[col_dir].apply(_clean) if col_dir else "",
    })
    
    # Remove linhas vazias e inválidas
    tmp = tmp[tmp["sig"] != ""]
    tmp = tmp[tmp["sig"].str.upper() != "OM"]
    tmp = tmp[tmp["sig"].str.upper() != "NR"]
    tmp = tmp[tmp["sig"].str.len() > 0]
    tmp = tmp[~tmp["sig"].str.contains("unnamed", case=False, na=False)]
    
    # Remove duplicatas (mantém a primeira ocorrência)
    tmp = tmp.drop_duplicates(subset=["sig"], keep="first")
    
    return tmp


def _add_fallback_option(
    options: list,
    disp2sig: dict
) -> Tuple[List[str], Dict[str, str], Dict[str, str]]:
    """Adiciona opção para OM não listada."""
    options.append("Outra / não listada…")
    disp2sig["Outra / não listada…"] = ""
    sig2dir: Dict[str, str] = {}
    return options, disp2sig, sig2dir


# =========================================================
# Callbacks
# =========================================================
def _on_om_change():
    """Callback quando a OM é alterada - atualiza diretoria automaticamente."""
    choice = st.session_state.get("om_choice")
    disp2sig = st.session_state.get("_disp2sig", {})
    sig2dir = st.session_state.get("_sig2dir", {})
    
    sig = disp2sig.get(choice or "", "")
    st.session_state["dir_resp"] = sig2dir.get(sig, "") if sig else ""
    st.session_state["om_sigla_out"] = ""


# =========================================================
# Chaves do formulário
# =========================================================
FORM_KEYS = [
    "om_choice", "om_sigla_out", "dir_resp",
    "tipo_vist", "local_inst", "urg", "data_limite", "motivo",
    "ref_opus", "objetivo", "vt_exec", "status_sem", "obs",
    "data_vist", "data_prev_conc", "meio_resp", "data_resp",
    "num_opus", "qd_total", "qd_exec",
]


def _reset_form():
    """Limpa todos os campos do formulário."""
    for k in FORM_KEYS:
        if k in st.session_state:
            del st.session_state[k]


# =========================================================
# Formulário Principal
# =========================================================
def _render_form(
    om_options: List[str],
    disp2sig: Dict[str, str],
    sig2dir: Dict[str, str]
) -> Tuple[dict, bool]:
    """Renderiza o formulário e retorna os dados e validação."""
    
    st.subheader("📥 Nova solicitação de vistoria")

    c1, c2 = st.columns(2)

    # -------- Coluna 1 --------
    with c1:
        choice = st.selectbox(
            "OM solicitante (sigla) *",
            options=om_options,
            index=None,
            placeholder="Selecione a OM…",
            key="om_choice",
            on_change=_on_om_change,
            help="Selecione a Organização Militar solicitante da vistoria"
        )

        om_sigla = ""
        if choice:
            om_sigla = disp2sig.get(choice, "")
            if choice == "Outra / não listada…":
                om_sigla = st.text_input(
                    "Informe a OM (sigla) *",
                    key="om_sigla_out",
                    help="Digite a sigla da OM não listada"
                ).strip()

        diretoria = st.text_input(
            "Diretoria responsável *",
            key="dir_resp",
            help="Preenchimento automático baseado na OM selecionada"
        )

        tipo_vistoria = st.selectbox(
            "Tipo de vistoria *",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
            key="tipo_vist",
            help="Selecione o tipo de vistoria a ser realizada"
        )

    # -------- Coluna 2 --------
    with c2:
        local = st.text_input(
            "Local / instalação *",
            key="local_inst",
            help="Endereço ou identificação do local da vistoria"
        )
        
        urgencia = st.selectbox(
            "Urgência *",
            ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"],
            index=0,
            key="urg",
            help="Classificação de urgência da solicitação"
        )
        
        data_limite: date | None = st.date_input(
            "Data limite (opcional)",
            value=None,
            format="YYYY/MM/DD",
            key="data_limite",
            help="Prazo para realização da vistoria, se houver"
        )

    motivo = st.text_area(
        "Motivo / justificativa (NAOM) *",
        height=120,
        key="motivo",
        help="Descreva o motivo e justificativa para a solicitação de vistoria"
    )

    # -------- Seção Complementares --------
    with st.expander("📋 Campos Complementares (opcional)", expanded=False):
        cc1, cc2 = st.columns(2)
        
        with cc1:
            referencia_opus = st.text_input("REFERÊNCIA OPUS", key="ref_opus")
            objetivo_contato = st.text_input("OBJETIVO (ADICIONAR POSSÍVEL CONTATO)", key="objetivo")
            vt_exec_por = st.text_input("VT EXECUTADA POR", key="vt_exec")
            status_atual = st.text_input("STATUS - ATUALIZAÇÃO SEMANAL", key="status_sem")
            obs = st.text_area("OBSERVAÇÕES", height=90, key="obs")

        with cc2:
            data_vistoria: date | None = st.date_input(
                "DATA DA VISTORIA",
                value=None,
                format="YYYY/MM/DD",
                key="data_vist"
            )
            data_prev_conc: date | None = st.date_input(
                "DATA/PREVISÃO DE CONCLUSÃO",
                value=None,
                format="YYYY/MM/DD",
                key="data_prev_conc"
            )
            meio_resposta = st.text_input("MEIO DE RESPOSTA DA SOLICITAÇÃO", key="meio_resp")
            data_resposta: date | None = st.date_input(
                "DATA DA RESPOSTA A SOLICITAÇÃO",
                value=None,
                format="YYYY/MM/DD",
                key="data_resp"
            )
            num_opus = st.text_input("Nº OPUS DA VISTORIA (SE FOR O CASO)", key="num_opus")
            qd_total = st.number_input(
                "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO",
                min_value=0,
                step=1,
                value=0,
                key="qd_total"
            )
            qd_exec = st.number_input(
                "QUANTIDADE DE DIAS PARA EXECUÇÃO",
                min_value=0,
                step=1,
                value=0,
                key="qd_exec"
            )

    # -------- Validações --------
    erros = []
    if not om_sigla:
        erros.append("Informe a OM solicitante")
    if not local.strip():
        erros.append("Informe o local/instalação")
    if not motivo.strip():
        erros.append("Descreva o motivo/justificativa (NAOM)")
    if not diretoria.strip():
        erros.append("Informe a diretoria responsável")

    if erros:
        st.warning("⚠️ **Campos obrigatórios não preenchidos:**\n\n• " + "\n• ".join(erros))

    # -------- Monta a linha de dados --------
    row = {
        "OBJETO DE VISTORIA": motivo.strip(),
        "OM APOIADA": om_sigla,
        "Diretoria Responsável": diretoria.strip(),
        "Classificação de Urgência": urgencia,
        "Situação": "SOLICITADA",
        "DATA DA SOLICITAÇÃO": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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

    ok = len(erros) == 0
    return row, ok


# =========================================================
# Página Principal
# =========================================================
def page():
    """Página de cadastro de solicitações de vistoria."""
    
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")
    
    st.markdown("""
    **Instruções:** Preencha todos os campos obrigatórios (*) para registrar uma nova 
    solicitação de vistoria. Os campos complementares são opcionais e podem ser preenchidos 
    posteriormente na tela de Status/Auditoria.
    """)
    
    st.divider()

    # Catálogo de OMs (com cache)
    om_options, disp2sig, sig2dir = _build_om_catalog()
    
    # Armazena no session_state para uso no callback
    st.session_state["_disp2sig"] = disp2sig
    st.session_state["_sig2dir"] = sig2dir
    
    # Mostra informação sobre quantas OMs foram carregadas
    total_oms = len(om_options) - 1  # -1 por causa de "Outra / não listada..."
    if total_oms > 0:
        st.info(f"ℹ️ {total_oms} Organizações Militares carregadas da base de validação")

    # Renderiza o formulário
    row, ok = _render_form(om_options, disp2sig, sig2dir)

    # -------- Botões de ação --------
    st.divider()
    
    col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 1])
    
    with col_btn1:
        if st.button(
            "💾 Salvar Solicitação",
            type="primary",
            disabled=not ok,
            use_container_width=True
        ):
            try:
                # Normaliza valores NaN
                clean_row = {k: ("" if pd.isna(v) else str(v)) for k, v in row.items()}
                
                # Salva no Google Sheets
                with st.spinner("Salvando no Google Sheets..."):
                    append_row(TAB_SOLICITACOES, clean_row)
                
                st.success("✅ Solicitação registrada com sucesso!")
                st.balloons()
                
                # Mostra informações do registro
                with st.expander("📄 Dados salvos"):
                    st.json(clean_row)
                
                # Limpa o formulário
                _reset_form()
                
                # Aguarda 2 segundos antes de recarregar
                import time
                time.sleep(2)
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ Falha ao salvar solicitação: {e}")
                import traceback
                with st.expander("🔍 Detalhes do erro"):
                    st.code(traceback.format_exc(), language="python")
    
    with col_btn2:
        if st.button("🔄 Limpar Formulário", use_container_width=True):
            _reset_form()
            st.rerun()
    
    with col_btn3:
        if st.button("📊 Ver Dashboard", use_container_width=True):
            st.switch_page("pages/dashboard.py")
