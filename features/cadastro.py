# features/cadastro.py
# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime
import pandas as pd

from core.data_loader import append_row, read_df

# ===============================
# Helpers para normalização e busca de colunas
# ===============================
def _normalize(text: str) -> str:
    return str(text).strip().lower()

def _find_column(df: pd.DataFrame, *column_names: str) -> str:
    columns_map = {_normalize(c): c for c in df.columns}
    for name in column_names:
        norm_name = _normalize(name)
        if norm_name in columns_map:
            return columns_map[norm_name]
    for name in column_names:
        norm_name = _normalize(name)
        for norm_col, orig_col in columns_map.items():
            if norm_name in norm_col:
                return orig_col
    raise KeyError(f"Coluna(s) '{column_names}' não encontradas em {df.columns.tolist()}")

# ===============================
# Carrega OMs e Diretorias da aba de validação
# ===============================
@st.cache_data(ttl=600)
def _load_oms_validadas() -> pd.DataFrame:
    df_raw = read_df("Validacao_de_Dados")

    def _try_promote_header(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        first_row = df.iloc[0].astype(str).str.lower()
        cols_lower = [str(c).lower() for c in df.columns]
        has_expected = lambda x: any(k in x for k in ("om", "sigla", "diretoria", "organizacao"))
        if (not any(has_expected(col) for col in cols_lower)) and any(has_expected(x) for x in first_row):
            df_new = df.copy()
            df_new.columns = df_new.iloc[0]
            df_new = df_new.drop(df_new.index[0]).reset_index(drop=True)
            return df_new
        return df

    df = _try_promote_header(df_raw)

    try:
        col_sigla = _find_column(df, "om", "sigla")
        col_nome = _find_column(df, "organização militar", "organizacao militar", "om nome", "nome")
        col_diretoria = _find_column(df, "diretoria responsável", "diretoria")
    except KeyError:
        st.error("Erro ao localizar colunas OM, Nome ou Diretoria na aba de validação.")
        return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])

    out = df[[col_sigla, col_nome, col_diretoria]].copy()
    out.columns = ["om_sigla", "om_nome", "diretoria"]
    for c in out.columns:
        out[c] = out[c].astype(str).str.strip()
    out = out[(out["om_sigla"] != "") & (out["diretoria"] != "")]
    out = out.drop_duplicates()

    return out

def _build_om_options(oms_df: pd.DataFrame):
    options_display, disp_to_sigla, sigla_to_diretoria, disp_to_diretoria = [], {}, {}, {}

    for _, r in oms_df.iterrows():
        sigla = r["om_sigla"]
        nome = r["om_nome"]
        diretoria = r["diretoria"]
        display = f"{sigla} — {nome}" if nome else sigla
        options_display.append(display)
        disp_to_sigla[display] = sigla
        sigla_to_diretoria[sigla] = diretoria
        disp_to_diretoria[display] = diretoria

    options_display.append("Outra / não listada…")
    disp_to_sigla["Outra / não listada…"] = ""
    disp_to_diretoria["Outra / não listada…"] = ""

    return options_display, disp_to_sigla, sigla_to_diretoria, disp_to_diretoria

# ===============================
# Interface principal do formulário
# ===============================
def _input_row():
    st.subheader("📥 Nova solicitação de vistoria")
    oms_df = _load_oms_validadas()
    options, disp2sig, sig2dir, disp2dir = _build_om_options(oms_df)

    st.caption(f"{len(oms_df)} Organizações Militares carregadas da base de validação")

    col1, col2 = st.columns(2)

    with col1:
        om_display = st.selectbox(
            "OM solicitante (sigla) *",
            options=options,
            key="om_choice",
            placeholder="Selecione ou digite…"
        )

        om_sigla = disp2sig.get(om_display, "")

        if om_display == "Outra / não listada…":
            om_sigla = st.text_input("Sigla da OM (manual)", key="om_sigla_out")
            diretoria = st.text_input("Diretoria responsável (manual)", key="diretoria_manual")
            st.session_state["diretoria_auto"] = ""
        else:
            default_dir = disp2dir.get(om_display, "") or sig2dir.get(om_sigla, "")
            st.session_state["diretoria_auto"] = default_dir
            diretoria = st.text_input(
                "Diretoria responsável *",
                value=st.session_state.get("diretoria_auto", ""),
                disabled=True,
                key="diretoria_auto"
            )

        tipo_vistoria = st.selectbox(
            "Tipo de vistoria *",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
        )

    with col2:
        local = st.text_input("Local / instalação *")
        urgencia = st.selectbox(
            "Urgência *",
            ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"],
            index=0
        )
        data_limite = st.date_input("Data limite (opcional)", value=None)

    motivo = st.text_area("Motivo / justificativa (NAOM) *", height=120)

    # Campos complementares opcionais
    with st.expander("📋 Campos Complementares (opcional)", expanded=False):
        referencia_opus = st.text_input("REFERÊNCIA OPUS")
        objetivo_contato = st.text_input("OBJETIVO (ADICIONAR POSSÍVEL CONTATO)")
        vt_exec_por = st.text_input("VT EXECUTADA POR")
        status_atual = st.text_input("STATUS - ATUALIZAÇÃO SEMANAL")
        obs = st.text_area("OBSERVAÇÕES", height=90)
        data_vistoria = st.date_input("DATA DA VISTORIA", value=None)
        data_prev_conc = st.date_input("DATA/PREVISÃO DE CONCLUSÃO", value=None)
        meio_resposta = st.text_input("MEIO DE RESPOSTA DA SOLICITAÇÃO")
        data_resposta = st.date_input("DATA DA RESPOSTA A SOLICITAÇÃO", value=None)
        num_opus = st.text_input("Nº OPUS DA VISTORIA (SE FOR O CASO)")
        qd_total = st.number_input("QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO", min_value=0, step=1, value=0)
        qd_exec = st.number_input("QUANTIDADE DE DIAS PARA EXECUÇÃO", min_value=0, step=1, value=0)

    erros = []
    if not (om_sigla or "").strip():
        erros.append("Informe a **OM**.")
    if om_display == "Outra / não listada…":
        if not (diretoria or "").strip():
            erros.append("Informe a **diretoria** (manual).")
    else:
        if not (diretoria or "").strip():
            erros.append("Diretoria não encontrada para a OM selecionada.")
    if not (local or "").strip():
        erros.append("Informe o **local/instalação**.")
    if not (motivo or "").strip():
        erros.append("Descreva o **motivo/justificativa**.")

    if erros:
        st.warning("Campos obrigatórios não preenchidos:\n\n• " + "\n• ".join(erros))

    row = {
        "numero": "",
        "data_solicitacao": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "om_solicitante": (om_sigla or "").strip(),
        "diretoria": (diretoria or "").strip(),
        "tipo_vistoria": tipo_vistoria,
        "local": (local or "").strip(),
        "urgencia": urgencia,
        "data_limite": data_limite.strftime("%Y-%m-%d") if data_limite else "",
        "motivo": (motivo or "").strip(),
        "status_atual": "SOLICITADA",
        "REFERÊNCIA OPUS": referencia_opus.strip(),
        "OBJETIVO (ADICIONAR POSSÍVEL CONTATO)": objetivo_contato.strip(),
        "VT EXECUTADA POR": vt_exec_por.strip(),
        "STATUS - ATUALIZAÇÃO SEMANAL": status_atual.strip(),
        "DATA DA VISTORIA": data_vistoria.strftime("%Y-%m-%d") if data_vistoria else "",
        "DATA/PREVISÃO DE CONCLUSÃO": data_prev_conc.strftime("%Y-%m-%d") if data_prev_conc else "",
        "MEIO DE RESPOSTA DA SOLICITAÇÃO": meio_resposta.strip(),
        "DATA DA RESPOSTA A SOLICITAÇÃO": data_resposta.strftime("%Y-%m-%d") if data_resposta else "",
        "Nº OPUS DA VISTORIA (SE FOR O CASO)": num_opus.strip(),
        "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO": str(int(qd_total)) if qd_total else "",
        "QUANTIDADE DE DIAS PARA EXECUÇÃO": str(int(qd_exec)) if qd_exec else "",
        "OBSERVAÇÕES": obs.strip(),
    }

    return row, (len(erros) == 0)

# ===============================
# Página principal
# ===============================
def page():

    if "main_menu" not in st.session_state:
        st.session_state["main_menu"] = "📊 Dashboard"
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")
    tabs_map = st.session_state.get("tabs_map", {})
    tab_solic = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")

    try:
        df_existente = read_df(tab_solic)
    except Exception:
        df_existente = pd.DataFrame()

    row, ok = _input_row()

    # Botões de ação
    col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 1])
    with col_btn1:
        if st.button("💾 Salvar Solicitação", type="primary", disabled=not ok, use_container_width=True):
            try:
                proximo = 1
                if not df_existente.empty and "numero" in df_existente.columns:
                    nums = pd.to_numeric(df_existente["numero"], errors="coerce").dropna()
                    if not nums.empty:
                        proximo = int(nums.max()) + 1
                row["numero"] = str(proximo)
                append_row(tab_solic, row)
                st.success(f"Solicitação **#{row['numero']}** cadastrada com sucesso!")
                st.rerun()
            except Exception as e:
                st.error(f"Falha ao salvar: {e}")
    with col_btn2:
        if st.button("🧹 Limpar Formulário", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    with col_btn3:
        if st.button("📊 Ver Dashboard", use_container_width=True):
            st.session_state["main_menu"] = "📊 Dashboard"  # Usa a string completa do menu
            st.rerun()


