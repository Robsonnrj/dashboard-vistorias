# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime, date
import pandas as pd

from core.data_loader import read_df, append_row
from core.config import TAB_SOLICITACOES, TAB_VALIDACAO

# ----------------------------
# Colunas-alvo na aba ACOMPANHAMENTO VISTORIAS
# ----------------------------
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

# Opções padrão
URGENCIAS = ["Não Prioridade", "Prioridade", "Urgente"]
SITUACOES = ["Não Atendida", "Em andamento", "Finalizada"]

# ----------------------------
# Helpers
# ----------------------------
def _clean_str(x) -> str:
    return (str(x).strip() if x is not None else "").strip()

def _date_to_str(d: date | None) -> str:
    if not d:
        return ""
    try:
        return pd.to_datetime(d).strftime("%Y-%m-%d")
    except Exception:
        return ""

def _load_oms_df() -> pd.DataFrame:
    """
    Tenta carregar a lista oficial de OMs:
      1) Validacao_de_Dados (preferência)
      2) ACOMPANHAMENTO VISTORIAS (fallback)
    Retorna DataFrame com colunas: om_sigla, om_nome, diretoria
    """
    # 1) Validacao_de_Dados
    for tab in (TAB_VALIDACAO, TAB_SOLICITACOES):
        try:
            df = read_df(tab)
        except Exception:
            df = pd.DataFrame()

        if df.empty:
            continue

        cols = {c.lower().strip(): c for c in df.columns}

        # Em Validacao_de_Dados os nomes costumam ser: "OM" (sigla), "Organização Militar" (nome), "Diretoria Responsável"
        cand_sigla = None
        for key in cols:
            if key in ("om", "sigla") or "om" in key and "apoia" not in key:
                cand_sigla = cols[key]; break

        cand_nome = None
        for key in cols:
            if "organiza" in key or key == "om":
                cand_nome = cols[key]; break

        cand_dir = None
        for key in cols:
            if "diretoria" in key:
                cand_dir = cols[key]; break

        # No fallback (aba base), podemos ter "OM APOIADA" e "Diretoria Responsável"
        if tab == TAB_SOLICITACOES and cand_sigla is None:
            cand_sigla = cols.get("om apoiada") or cols.get("om")

        if all(c is None for c in (cand_sigla, cand_nome, cand_dir)):
            continue

        out = pd.DataFrame({
            "om_sigla": df[cand_sigla] if cand_sigla in df.columns else "",
            "om_nome":  df[cand_nome]  if cand_nome  in df.columns else "",
            "diretoria": df[cand_dir]  if cand_dir  in df.columns else "",
        })
        for c in out.columns:
            out[c] = out[c].fillna("").astype(str).str.strip()
        out = out[out["om_sigla"] != ""]
        out = out.drop_duplicates(subset=["om_sigla"], keep="first")
        if not out.empty:
            return out

    return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])

def _build_om_options(oms_df: pd.DataFrame):
    """
    Retorna:
      - options_display: lista de rótulos "SIGLA — Nome"
      - disp_to_sigla:   mapeia rótulo -> sigla
      - sigla_to_dir:    mapeia sigla  -> diretoria
    """
    options_display, disp_to_sigla, sigla_to_dir = [], {}, {}
    if not oms_df.empty:
        for _, r in oms_df.iterrows():
            sig = _clean_str(r.get("om_sigla"))
            nom = _clean_str(r.get("om_nome"))
            dire = _clean_str(r.get("diretoria"))
            if not sig:
                continue
            display = f"{sig} — {nom}" if nom else sig
            options_display.append(display)
            disp_to_sigla[display] = sig
            # primeira diretoria encontrada vence
            if sig not in sigla_to_dir:
                sigla_to_dir[sig] = dire

    options_display.append("Outra / não listada…")
    disp_to_sigla["Outra / não listada…"] = ""
    return options_display, disp_to_sigla, sigla_to_dir

# ----------------------------
# Página
# ----------------------------
def page():
    st.header("📝 Cadastro — Acompanhamento de Vistorias (Grava direto na aba)")

    # Carrega OMs
    oms_df = _load_oms_df()
    options, disp2sig, sig2dir = _build_om_options(oms_df)

    # Formulário
    with st.form("form_vistoria", clear_on_submit=False):
        st.subheader("Dados principais")
        c1, c2 = st.columns([1.3, 1])
        with c1:
            objeto = st.text_input("OBJETO DE VISTORIA *", placeholder="Ex.: Inspeção elétrica na sala de máquinas")
            om_display = st.selectbox("OM APOIADA *", options=options, index=None, placeholder="Selecione…")
            om_sigla = disp2sig.get(om_display or "", "")
            dir_auto = sig2dir.get(om_sigla, "")
            diretoria = st.text_input("Diretoria Responsável *", value=dir_auto if om_sigla else "")
            urg = st.selectbox("Classificação de Urgência *", URGENCIAS, index=0)
            sit = st.selectbox("Situação *", SITUACOES, index=0)
        with c2:
            data_solic_1 = st.date_input("DATA DA SOLICITAÇÃO *", value=date.today())
            data_solic_2 = st.date_input("DATA DA SOLICITAÇÃO_2", value=None)
            data_vist = st.date_input("DATA DA VISTORIA", value=None)
            previsao_conc = st.date_input("DATA/PREVISÃO DE CONCLUSÃO", value=None)
            data_resp = st.date_input("DATA DA RESPOSTA A SOLICITAÇÃO", value=None)

        st.subheader("Complementares")
        c3, c4 = st.columns(2)
        with c3:
            ref_opus = st.text_input("REFERÊNCIA OPUS")
            objetivo_contato = st.text_input("OBJETIVO (ADICIONAR POSSÍVEL CONTATO)")
            vt_exec_por = st.text_input("VT EXECUTADA POR")
            status_semana = st.text_input("STATUS - ATUALIZAÇÃO SEMANAL")
        with c4:
            meio_resposta = st.text_input("MEIO DE RESPOSTA DA SOLICITAÇÃO")
            num_opus = st.text_input("Nº OPUS DA VISTORIA (SE FOR O CASO)")
            qt_dias_total = st.number_input("QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO", min_value=0, step=1)
            qt_dias_exec = st.number_input("QUANTIDADE DE DIAS PARA EXECUÇÃO", min_value=0, step=1)
        obs = st.text_area("OBSERVAÇÕES", height=120)

        # Validação mínima
        erros = []
        if not _clean_str(objeto):
            erros.append("Informe o **OBJETO DE VISTORIA**.")
        if not _clean_str(om_sigla):
            erros.append("Selecione/informe a **OM APOIADA**.")
        if not _clean_str(diretoria):
            erros.append("Informe a **Diretoria Responsável**.")
        if erros:
            st.warning("• " + "\n• ".join(erros))

        salvar = st.form_submit_button("💾 Salvar na aba ACOMPANHAMENTO VISTORIAS", type="primary", disabled=bool(erros))

    if salvar:
        try:
            # Monta payload exatamente com as colunas do sheet
            row = {
                "OBJETO DE VISTORIA": _clean_str(objeto),
                "OM APOIADA": _clean_str(om_sigla),
                "Diretoria Responsável": _clean_str(diretoria),
                "Classificação de Urgência": _clean_str(urg),
                "Situação": _clean_str(sit),
                "DATA DA SOLICITAÇÃO": _date_to_str(data_solic_1),
                "DATA DA SOLICITAÇÃO_2": _date_to_str(data_solic_2),
                "REFERÊNCIA OPUS": _clean_str(ref_opus),
                "OBJETIVO (ADICIONAR POSSÍVEL CONTATO)": _clean_str(objetivo_contato),
                "DATA DA VISTORIA": _date_to_str(data_vist),
                "VT EXECUTADA POR": _clean_str(vt_exec_por),
                "STATUS - ATUALIZAÇÃO SEMANAL": _clean_str(status_semana),
                "DATA/PREVISÃO DE CONCLUSÃO": _date_to_str(previsao_conc),
                "MEIO DE RESPOSTA DA SOLICITAÇÃO": _clean_str(meio_resposta),
                "DATA DA RESPOSTA A SOLICITAÇÃO": _date_to_str(data_resp),
                "Nº OPUS DA VISTORIA (SE FOR O CASO)": _clean_str(num_opus),
                "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO": int(qt_dias_total) if pd.notna(qt_dias_total) else "",
                "QUANTIDADE DE DIAS PARA EXECUÇÃO": int(qt_dias_exec) if pd.notna(qt_dias_exec) else "",
                "OBSERVAÇÕES": _clean_str(obs),
            }

            # Garante que só existam as chaves esperadas e que todas existam
            payload = {col: row.get(col, "") for col in SHEET_COLUMNS}

            # Grava: usa a ordem do cabeçalho da aba (append_row já faz o mapeamento pelo header da linha 1)
            append_row(TAB_SOLICITACOES, payload)

            st.success("✅ Registro incluído em **ACOMPANHAMENTO VISTORIAS**.")
            st.rerun()

        except Exception as e:
            st.error(f"Falha ao salvar: {e}")

    # Visual: últimas 30 inserções para conferência (se existir a aba)
    try:
        df_base = read_df(TAB_SOLICITACOES)
        if not df_base.empty:
            # tenta ordenar por DATA DA SOLICITAÇÃO se existir
            order_col = None
            for cand in ["DATA DA SOLICITAÇÃO", "DATA DA VISTORIA"]:
                if cand in df_base.columns:
                    order_col = cand; break
            if order_col:
                df_base[order_col] = pd.to_datetime(df_base[order_col], errors="coerce")
                df_base = df_base.sort_values(order_col, ascending=False)
            st.divider()
            st.caption("Visualização rápida da base (últimos 30):")
            st.dataframe(df_base.tail(30), use_container_width=True, height=360)
    except Exception:
        pass
