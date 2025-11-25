# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime
import pandas as pd

from core.data_loader import append_row, read_df
from core.utils import norm, pick_col

# Ícone de pasta (VIS-001) hospedado no GitHub
ICON_VIS_001 = "https://raw.githubusercontent.com/Robsonnrj/dashboard-vistorias/main/folder_15779310.png"


def _N(x: str) -> str:
    """Normaliza string: strip + upper (para siglas/chaves)."""
    return (str(x or "")).strip().upper()


@st.cache_data(ttl=600)
def _load_oms_validadas() -> pd.DataFrame:
    """
    Lê a aba 'Validacao_de_Dados' já tratada pelo read_df e mapeia as colunas por sinônimos.
    Retorna um DF com colunas: om_sigla, om_nome, diretoria (todas sem espaços extras).
    """
    df = read_df("Validacao_de_Dados", use_cache=False)
    if df is None or df.empty:
        st.error("Base de validação vazia ou não encontrada.")
        return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])

    col_sigla = pick_col(df, [
        "sigla", "sigla om", "om (sigla)", "sigla da om", "om sigla", "sigla/om",
        "om", "om solicitante (sigla)"
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
    if not col_sigla:
        missings.append("OM/Sigla")
    if not col_nome:
        missings.append("Nome da OM")
    if not col_dir:
        missings.append("Diretoria")
    if missings:
        with st.expander("🔎 Diagnóstico — colunas disponíveis na aba de validação"):
            st.write(list(df.columns))
        st.error(
            "Colunas não encontradas na aba de validação: "
            + ", ".join(missings)
            + ". Corrija a planilha ou verifique os nomes das colunas."
        )
        return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])

    df_out = (
        df.rename(columns={col_sigla: "om_sigla", col_nome: "om_nome", col_dir: "diretoria"})
        [["om_sigla", "om_nome", "diretoria"]]
        .copy()
    )
    # limpeza e normalização
    for c in ["om_sigla", "om_nome", "diretoria"]:
        df_out[c] = df_out[c].astype(str).str.strip()
    # remove linhas vazias
    df_out = df_out[(df_out["om_sigla"] != "") & (df_out["diretoria"] != "")]
    # normaliza sigla para UPPER para chaves consistentes
    df_out["om_sigla_norm"] = df_out["om_sigla"].map(_N)
    # remove duplicados por sigla normalizada
    df_out = df_out.drop_duplicates(subset=["om_sigla_norm"], keep="first").reset_index(drop=True)

    # diagnóstico opcional
    with st.expander("🧭 Mapeamento detectado na validação"):
        st.write(f"OMs válidas: {len(df_out)}")
        st.write(df_out.head(10))

    return df_out[["om_sigla", "om_nome", "diretoria", "om_sigla_norm"]]


def _build_om_options(oms_df: pd.DataFrame):
    """
    Monta:
      - options_display: lista para o select (ex.: 'IME — Instituto Militar de Engenharia')
      - disp_to_sigla:   display -> SIGLA (original)
      - sigla_to_dir:    SIGLA_NORMALIZADA -> Diretoria
      - disp_to_dir:     display -> Diretoria
    """
    if oms_df.empty:
        return ["Outra / não listada…"], {"Outra / não listada…": ""}, {}, {"Outra / não listada…": ""}

    options_display, disp_to_sigla, sigla_to_dir, disp_to_dir = [], {}, {}, {}
    for _, r in oms_df.iterrows():
        sigla = (r["om_sigla"] or "").strip()
        sigla_norm = _N(sigla)
        nome = (r["om_nome"] or "").strip()
        diretoria = (r["diretoria"] or "").strip()

        display = f"{sigla} — {nome}" if nome else sigla
        options_display.append(display)
        disp_to_sigla[display] = sigla
        sigla_to_dir[sigla_norm] = diretoria
        disp_to_dir[display] = diretoria

    # opção manual
    options_display.append("Outra / não listada…")
    disp_to_sigla["Outra / não listada…"] = ""
    disp_to_dir["Outra / não listada…"] = ""

    return options_display, disp_to_sigla, sigla_to_dir, disp_to_dir


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

        # OM e diretoria derivadas da seleção (determinístico)
        if om_display == "Outra / não listada…":
            om_sigla = st.text_input("Sigla da OM (manual)", key="om_sigla_out")
            diretoria = st.text_input("Diretoria responsável (manual)", key="diretoria_manual")
        else:
            om_sigla = disp2sig.get(om_display, "")
            diretoria = disp2dir.get(om_display) or sig2dir.get(_N(om_sigla), "")
            # atualiza sessão e mostra campo somente-leitura que sempre reflete a seleção
            st.session_state["diretoria_view"] = diretoria
            st.text_input("Diretoria responsável *", key="diretoria_view", disabled=True)

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

    # Validações
    erros = []
    if om_display == "Outra / não listada…":
        if not (om_sigla or "").strip():
            erros.append("Informe a **OM**.")
        if not (diretoria or "").strip():
            erros.append("Informe a **diretoria** (manual).")
    else:
        if not (om_sigla or "").strip():
            erros.append("Informe a **OM**.")
        # diretoria vem do dicionário; garante
        if not (diretoria or "").strip():
            erros.append("Diretoria não encontrada para a OM selecionada.")

    if not (local or "").strip():
        erros.append("Informe o **local/instalação**.")
    if not (motivo or "").strip():
        erros.append("Descreva o **motivo/justificativa**.")

    if erros:
        st.warning("Campos obrigatórios não preenchidos:\n\n• " + "\n• ".join(erros))

    # linhas finais (usa diretoria mapeada automaticamente quando não for manual)
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
        "REFERÊNCIA OPUS": (referencia_opus or "").strip(),
        "OBJETIVO (ADICIONAR POSSÍVEL CONTATO)": (objetivo_contato or "").strip(),
        "VT EXECUTADA POR": (vt_exec_por or "").strip(),
        "STATUS - ATUALIZAÇÃO SEMANAL": (status_atual or "").strip(),
        "DATA DA VISTORIA": data_vistoria.strftime("%Y-%m-%d") if data_vistoria else "",
        "DATA/PREVISÃO DE CONCLUSÃO": data_prev_conc.strftime("%Y-%m-%d") if data_prev_conc else "",
        "MEIO DE RESPOSTA DA SOLICITAÇÃO": (meio_resposta or "").strip(),
        "DATA DA RESPOSTA A SOLICITAÇÃO": data_resposta.strftime("%Y-%m-%d") if data_resposta else "",
        "Nº OPUS DA VISTORIA (SE FOR O CASO)": (num_opus or "").strip(),
        "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO": str(int(qd_total)) if qd_total else "",
        "QUANTIDADE DE DIAS PARA EXECUÇÃO": str(int(qd_exec)) if qd_exec else "",
        "OBSERVAÇÕES": (obs or "").strip(),
    }
    return row, (len(erros) == 0)


def page():
    if "tabs_map" not in st.session_state:
        st.session_state["tabs_map"] = {
            "solicitacoes": "ACOMPANHAMENTO VISTORIAS",
            "validacao": "Validacao_de_Dados",
            "auditoria": "Auditoria_Vistorias",
        }

    # Cabeçalho com ícone de pasta + título VIS-001
    st.markdown(
        f"""
        <div style="display:flex; align-items:center; gap:10px; margin-bottom:0.8rem;">
          <img src="{ICON_VIS_001}" style="width:32px; height:32px;" />
          <h1 style="margin:0; font-size:1.9rem;">VIS-001 — Cadastro de Solicitação de Vistoria</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )

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
