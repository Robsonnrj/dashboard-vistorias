# features/cadastro.py
# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime
import pandas as pd

from core.data_loader import append_row, read_df

# -------------------------------
# Carrega OMs e diretorias da planilha
# -------------------------------
def _load_oms_df() -> pd.DataFrame:
    for tab in ("Validacao_de_Dados", "ACOMPANHAMENTO VISTORIAS"):
        try:
            df = read_df(tab)
        except Exception:
            df = pd.DataFrame()
        if df.empty:
            continue
        cols = {c.lower().strip(): c for c in df.columns}
        sigla = next((cols[k] for k in cols if "sigla" in k or k in ("om", "om apoiada")), None)
        nome  = next((cols[k] for k in cols if "organiza" in k or "om" == k), None)
        diret = next((cols[k] for k in cols if "diretoria" in k), None)
        if not sigla and "OM" in df.columns:
            sigla = "OM"
        if not diret: continue
        out = pd.DataFrame({
            "om_sigla": df[sigla] if sigla in df.columns else pd.Series(dtype=str),
            "om_nome":  df[nome]  if nome  in df.columns else pd.Series(dtype=str),
            "diretoria": df[diret],
        }).copy()
        for c in ("om_sigla", "om_nome", "diretoria"):
            if c in out.columns:
                out[c] = out[c].fillna("").astype(str).str.strip()
        out = out[out["diretoria"] != ""]
        out = out.drop_duplicates(subset=["om_sigla", "om_nome", "diretoria"])
        if not out.empty:
            return out
    return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])

def _build_om_options(oms_df: pd.DataFrame):
    options_display, disp_to_sigla = [], {}
    sigla_to_dir = {}
    if not oms_df.empty:
        for _, r in oms_df.iterrows():
            sig = str(r.get("om_sigla", "") or "").strip()
            nom = str(r.get("om_nome", "") or "").strip()
            dire = str(r.get("diretoria", "") or "").strip()
            if not sig:
                continue
            display = f"{sig} — {nom}" if nom else sig
            options_display.append(display)
            disp_to_sigla[display] = sig
            # sempre pega a última diretoria válida para cada sigla, isso cobre reprocessamentos.
            sigla_to_dir[sig] = dire
    options_display.append("Outra / não listada…")
    disp_to_sigla["Outra / não listada…"] = ""
    return options_display, disp_to_sigla, sigla_to_dir

def _input_row(oms_df: pd.DataFrame):
    st.subheader("📥 Nova solicitação de vistoria")
    options, disp2sig, sig2dir = _build_om_options(oms_df)
    col1, col2 = st.columns(2)
    with col1:
        om_display = st.selectbox(
            "OM solicitante",
            options=options,
            index=None,
            placeholder="Selecione ou digite…",
            help="Escolha a OM ou selecione 'Outra / não listada…' para inserir manualmente."
        )
        om_sigla = disp2sig.get(om_display, "")
        diretoria_auto = sig2dir.get(om_sigla, "")
        if om_display == "Outra / não listada…":
            om_sigla = st.text_input("Sigla da OM (manual)", "")
            diretoria = st.text_input("Diretoria responsável (manual)", "")
            diretoria_field_disabled = False
        else:
            diretoria = st.text_input(
                "Diretoria responsável",
                value=diretoria_auto,
                disabled=True,
                help="Preenchido automaticamente conforme OM"
            )
            diretoria_field_disabled = True
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
    erros = []
    if not om_sigla.strip():
        erros.append("Informe a **OM**.")
    if not diretoria.strip():
        if not diretoria_field_disabled:  # Só pede se o campo for manual
            erros.append("Informe a **diretoria** (selecione uma OM conhecida ou preencha manualmente).")
    if not local.strip():
        erros.append("Informe o **local/instalação**.")
    if not motivo.strip():
        erros.append("Descreva o **motivo/justificativa**.")
    if erros:
        st.warning("• " + "\n• ".join(erros))
    row = {
        "numero": "",  # será preenchido ao salvar
        "data_solicitacao": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "om_solicitante": om_sigla.strip(),
        "diretoria": diretoria.strip(),
        "tipo_vistoria": tipo_vistoria,
        "local": local.strip(),
        "urgencia": urgencia,
        "data_limite": data_limite.strftime("%Y-%m-%d") if data_limite else "",
        "motivo": motivo.strip(),
        "status_atual": "SOLICITADA",
    }
    return row, (len(erros) == 0)

def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")
    # Abas definidas na sidebar (mantém compatibilidade)
    tabs_map = st.session_state.get("tabs_map", {})
    tab_solic = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")
    try:
        df_existente = read_df(tab_solic)
    except Exception:
        df_existente = pd.DataFrame()
    # 🔹 Carrega OMs e diretorias para autocomplete
    oms_df = _load_oms_df()
    # Formulário
    row, ok = _input_row(oms_df)
    # Salvar
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
    # Últimos registros
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
