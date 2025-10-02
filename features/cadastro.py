# features/cadastro.py
# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime
import pandas as pd

from core.data_loader import append_row, read_df

# -------------------------------
# Helpers p/ detectar colunas pelo nome
# -------------------------------
def _norm(s: str) -> str:
    return str(s).strip().lower()

def _find_col(df: pd.DataFrame, *cands: str) -> str:
    cols = { _norm(c): c for c in df.columns }
    for c in cands:
        k = _norm(c)
        if k in cols: 
            return cols[k]
    # fallback: tenta por "contém"
    for want in cands:
        for k, orig in cols.items():
            if _norm(want) in k:
                return orig
    raise KeyError(f"Não achei colunas {cands} em {list(df.columns)}")

# -------------------------------
# Carrega OMs e diretorias diretamente da aba de validação
# -------------------------------
def _load_oms_validadas() -> pd.DataFrame:
    """
    Lê a aba 'Validacao_de_Dados' e retorna colunas: om_sigla, om_nome, diretoria
    """
    # se o header da planilha começa na 2ª linha, mantenha header=1; se for 1ª, use header=0.
    df = read_df("Validacao_de_Dados", header=0)

    col_sigla = _find_col(df, "om", "sigla")
    col_nome  = _find_col(df, "organização militar", "organizacao militar", "om nome", "nome")
    col_dir   = _find_col(df, "diretoria responsável", "diretoria responsavel", "diretoria")

    out = df.rename(columns={
        col_sigla: "om_sigla",
        col_nome : "om_nome",
        col_dir  : "diretoria",
    })[["om_sigla", "om_nome", "diretoria"]].copy()

    # limpa
    for c in ["om_sigla", "om_nome", "diretoria"]:
        out[c] = out[c].astype(str).str.strip()

    out = out[(out["om_sigla"] != "") & (out["diretoria"] != "")]
    out = out.drop_duplicates(subset=["om_sigla", "diretoria"])
    return out

def _build_om_options(oms_df: pd.DataFrame):
    options_display, disp_to_sigla, sigla_to_dir = [], {}, {}
    for _, r in oms_df.iterrows():
        sig, nom, dire = r["om_sigla"], r["om_nome"], r["diretoria"]
        display = f"{sig} — {nom}" if nom else sig
        options_display.append(display)
        disp_to_sigla[display] = sig
        # se houver mais de uma diretoria para a mesma sigla, a última prevalece
        sigla_to_dir[sig] = dire
    options_display.append("Outra / não listada…")
    disp_to_sigla["Outra / não listada…"] = ""
    return options_display, disp_to_sigla, sigla_to_dir

def _on_om_change(disp2sig: dict, sig2dir: dict):
    """Quando a OM muda, atualiza a diretoria automática e limpa campos manuais."""
    choice = st.session_state.get("om_choice")
    sig = disp2sig.get(choice or "", "")
    st.session_state["om_sigla_out"] = ""  # limpa o manual
    st.session_state["diretoria_manual"] = ""
    st.session_state["diretoria_auto"] = sig2dir.get(sig, "")

def _input_row():
    st.subheader("📥 Nova solicitação de vistoria")
    oms_df = _load_oms_validadas()
    options, disp2sig, sig2dir = _build_om_options(oms_df)
    st.caption(f"{len(oms_df)} Organizações Militares carregadas da base de validação")

    col1, col2 = st.columns(2)

    with col1:
        om_display = st.selectbox(
            "OM solicitante (sigla) *",
            options=options,
            index=None,
            placeholder="Selecione ou digite…",
            key="om_choice",
            on_change=_on_om_change,
            args=(disp2sig, sig2dir),
        )
        om_sigla = disp2sig.get(om_display or "", "")

        # diretoria: automática para OMs conhecidas; manual se "Outra…"
        if om_display == "Outra / não listada…":
            om_sigla = st.text_input("Sigla da OM (manual)", key="om_sigla_out")
            diretoria = st.text_input("Diretoria responsável (manual)", key="diretoria_manual")
        else:
            diretoria_auto = st.session_state.get("diretoria_auto", sig2dir.get(om_sigla, ""))
            diretoria = diretoria_auto
            st.text_input("Diretoria responsável *", value=diretoria_auto, disabled=True, key="diretoria_auto_show")

        tipo_vistoria = st.selectbox(
            "Tipo de vistoria *",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
        )

    with col2:
        local = st.text_input("Local / instalação *")
        urgencia = st.selectbox("Urgência *", ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"], index=0)
        data_limite = st.date_input("Data limite (opcional)", value=None)

    motivo = st.text_area("Motivo / justificativa (NAOM) *", height=120)

    # validação
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
    }
    return row, (len(erros) == 0)
    
def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")
    tabs_map = st.session_state.get("tabs_map", {})
    tab_solic = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")
    try:
        df_existente = read_df(tab_solic)
    except Exception:
        df_existente = pd.DataFrame()
    row, ok = _input_row()
    if st.button("💾 Salvar solicitação", type="primary", disabled=not ok):
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
    st.divider()
    st.subheader("📄 Últimas solicitações")
    if not df_existente.empty:
        c_data = next((c for c in df_existente.columns if "data" in c.lower()), None)
        if c_data:
            df_existente[c_data] = pd.to_datetime(df_existente[c_data], errors="coerce")
            df_existente = df_existente.sort_values(c_data, ascending=False)
        st.dataframe(df_existente.head(50), use_container_width=True, height=360)
    else:
        st.caption("Ainda não há registros.")
