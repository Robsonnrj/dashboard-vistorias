# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime
import pandas as pd
from pathlib import Path
import base64

from core.data_loader import append_row, read_df
from core.utils import norm, pick_col

# ============================
# Helper para converter imagens locais em Base64 (robusto)
# ============================
def _img_to_base64(path: str) -> str:
    """Lê uma imagem local e converte para base64 (melhor método para Streamlit)."""
    file = Path(path)
    if not file.exists():
        st.error(f"Imagem não encontrada: {path}")
        return ""
    return base64.b64encode(file.read_bytes()).decode()


# ===================================
# VIS-001 — Ícone local convertido para base64
# ===================================
ICON_VIS_001_B64 = _img_to_base64("folder_15779310.png")


def _N(x: str) -> str:
    """Normaliza string: strip + upper (para siglas/chaves)."""
    return (str(x or "")).strip().upper()


@st.cache_data(ttl=600)
def _load_oms_validadas() -> pd.DataFrame:
    df = read_df("Validacao_de_Dados", use_cache=False)
    if df is None or df.empty:
        st.error("Base de validação vazia ou não encontrada.")
        return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])

    # >>>>> AQUI ESTÁ O AJUSTE IMPORTANTE <<<<<
    # Preferimos "OM" antes de "Sigla" para não confundir com a coluna de órgão setorial
    col_sigla = pick_col(df, [
        "om", "OM", "om (sigla)", "om sigla", "om solicitante (sigla)",
        "sigla om", "sigla da om", "sigla/om", "sigla"
    ])
    col_nome = pick_col(df, [
        "organização militar", "organizacao militar", "om nome", "nome da om",
        "nome", "unidade/om", "unidade", "om solicitante (nome)"
    ])
    col_dir = pick_col(df, [
        "diretoria responsável", "diretoria responsavel", "diretoria",
        "dir responsável", "dir responsavel", "direção", "diretoria/om"
    ])

    missings = []
    if not col_sigla: missings.append("OM/Sigla")
    if not col_nome:  missings.append("Nome da OM")
    if not col_dir:   missings.append("Diretoria")

    if missings:
        st.error("Colunas não encontradas: " + ", ".join(missings))
        return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])

    df_out = (
        df.rename(columns={col_sigla: "om_sigla", col_nome: "om_nome", col_dir: "diretoria"})
          [["om_sigla", "om_nome", "diretoria"]].copy()
    )

    for c in ["om_sigla", "om_nome", "diretoria"]:
        df_out[c] = df_out[c].astype(str).str.strip()

    # mantém apenas linhas com OM e diretoria preenchidas
    df_out = df_out[(df_out["om_sigla"] != "") & (df_out["diretoria"] != "")]
    # normaliza sigla para evitar duplicados com diferença de caixa/espaço
    df_out["om_sigla_norm"] = df_out["om_sigla"].map(_N)
    df_out = df_out.drop_duplicates(subset=["om_sigla_norm"], keep="first").reset_index(drop=True)

    # (Opcional) diagnóstico rápido
    with st.expander("🧭 Mapeamento detectado na validação"):
        st.write(f"OMs válidas encontradas: **{len(df_out)}**")
        st.dataframe(df_out[["om_sigla", "om_nome", "diretoria"]].head(20), use_container_width=True)

    return df_out[["om_sigla", "om_nome", "diretoria", "om_sigla_norm"]]


def _build_om_options(oms_df: pd.DataFrame):
    if oms_df.empty:
        return ["Outra / não listada…"], {"Outra / não listada…": ""}, {}, {"Outra / não listada…": ""}

    options_display, disp_to_sigla, sig_to_dir, disp_to_dir = [], {}, {}, {}

    for _, r in oms_df.iterrows():
        sigla = r["om_sigla"].strip()
        sig_norm = _N(sigla)
        nome = r["om_nome"].strip()
        diretoria = r["diretoria"].strip()

        display = f"{sigla} — {nome}" if nome else sigla
        options_display.append(display)

        disp_to_sigla[display] = sigla
        sig_to_dir[sig_norm] = diretoria
        disp_to_dir[display] = diretoria

    options_display.append("Outra / não listada…")
    disp_to_sigla["Outra / não listada…"] = ""
    disp_to_dir["Outra / não listada…"] = ""

    return options_display, disp_to_sigla, sig_to_dir, disp_to_dir


def _input_row():
    st.subheader("📥 Nova solicitação de vistoria")

    oms_df = _load_oms_validadas()
    options, disp2sig, sig2dir, disp2dir = _build_om_options(oms_df)
    st.caption(f"{len(oms_df)} Organizações Militares carregadas da base de validação")

    col1, col2 = st.columns(2)

    with col1:
        om_display = st.selectbox(
            "OM solicitante (sigla) *",
            options=options, key="om_choice",
            placeholder="Selecione ou digite…"
        )

        if om_display == "Outra / não listada…":
            om_sigla = st.text_input("Sigla da OM (manual)", key="om_sigla_out")
            diretoria = st.text_input("Diretoria responsável (manual)", key="diretoria_manual")
        else:
            om_sigla = disp2sig.get(om_display, "")
            diretoria = disp2dir.get(om_display) or sig2dir.get(_N(om_sigla), "")

            st.session_state["diretoria_view"] = diretoria
            st.text_input("Diretoria responsável *", key="diretoria_view", disabled=True)

        tipo_vistoria = st.selectbox(
            "Tipo de vistoria *",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"]
        )

    with col2:
        local = st.text_input("Local / instalação *")
        urgencia = st.selectbox("Urgência *", ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"])
        data_limite = st.date_input("Data limite (opcional)", value=None)

    motivo = st.text_area("Motivo / justificativa (NAOM) *", height=120)

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
        qd_total = st.number_input("QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO", min_value=0)
        qd_exec = st.number_input("QUANTIDADE DE DIAS PARA EXECUÇÃO", min_value=0)

    erros = []
    if not (om_sigla or "").strip(): erros.append("Informe a **OM**.")
    if not (diretoria or "").strip(): erros.append("Informe a **diretoria**.")
    if not (local or "").strip(): erros.append("Informe o **local/instalação**.")
    if not (motivo or "").strip(): erros.append("Descreva o **motivo/justificativa**.")

    if erros:
        st.warning("Campos obrigatórios não preenchidos:\n\n• " + "\n• ".join(erros))

    row = {
        "numero": "",
        "data_solicitacao": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "om_solicitante": om_sigla.strip(),
        "diretoria": diretoria.strip(),
        "tipo_vistoria": tipo_vistoria,
        "local": local.strip(),
        "urgencia": urgencia,
        "data_limite": data_limite.strftime("%Y-%m-%d") if data_limite else "",
        "motivo": motivo.strip(),
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
        "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO": str(int(qd_total)),
        "QUANTIDADE DE DIAS PARA EXECUÇÃO": str(int(qd_exec)),
        "OBSERVAÇÕES": obs.strip(),
    }

    return row, (len(erros) == 0)


def page():
    # Header com ícone Base64 (robusto)
    st.markdown(
        f"""
        <div style="display:flex; align-items:center; gap:10px; margin-bottom:0.8rem;">
            <img src="data:image/png;base64,{ICON_VIS_001_B64}" style="width:32px; height:32px;" />
            <h1 style="margin:0; font-size:1.9rem;"> Cadastro de Solicitação de Vistoria</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if "tabs_map" not in st.session_state:
        st.session_state["tabs_map"] = {
            "solicitacoes": "ACOMPANHAMENTO VISTORIAS",
            "validacao": "Validacao_de_Dados",
            "auditoria": "Auditoria_Vistorias",
        }

    tabs_map = st.session_state.get("tabs_map", {})
    tab_solic = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")

    try:
        df_existente = read_df(tab_solic, use_cache=False)
    except Exception:
        df_existente = pd.DataFrame()

    row, ok = _input_row()

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
