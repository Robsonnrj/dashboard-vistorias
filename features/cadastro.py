# features/cadastro.py
# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime
import pandas as pd

from core.data_loader import append_row, read_df

# -------------------------------
# Carrega OMs e diretorias diretamente da aba de validação
# -------------------------------
def _load_oms_validadas() -> pd.DataFrame:
    """
    Carrega OM, nome e diretoria da aba 'Validacao_de_Dados'.
    """
    df = read_df("Validacao_de_Dados", header=1)
    # Ajusta os nomes das colunas com base em sua planilha:
    df = df.rename(columns={
        df.columns[1]: "om_sigla",          # 'OM'
        df.columns[2]: "om_nome",           # 'Organização Militar'
        df.columns[3]: "diretoria"          # 'Diretoria Responsável'
    })
    out = df[["om_sigla", "om_nome", "diretoria"]].copy()
    out = out[out["om_sigla"].notna() & out["diretoria"].notna()]
    out = out[(out["om_sigla"] != "") & (out["diretoria"] != "")]
    out = out.drop_duplicates(subset=["om_sigla", "diretoria"])
    return out

def _build_om_options(oms_df: pd.DataFrame):
    options_display, disp_to_sigla, sigla_to_dir = [], {}, {}
    for _, r in oms_df.iterrows():
        sig, nom, dire = str(r["om_sigla"]), str(r["om_nome"]), str(r["diretoria"])
        display = f"{sig} — {nom}" if nom else sig
        options_display.append(display)
        disp_to_sigla[display] = sig
        sigla_to_dir[sig] = dire
    options_display.append("Outra / não listada…")
    disp_to_sigla["Outra / não listada…"] = ""
    return options_display, disp_to_sigla, sigla_to_dir

def _input_row():
    st.subheader("📥 Nova solicitação de vistoria")
    oms_df = _load_oms_validadas()
    options, disp2sig, sig2dir = _build_om_options(oms_df)
    col1, col2 = st.columns(2)

    with col1:
        om_display = st.selectbox(
            "OM solicitante",
            options=options,
            index=None,
            placeholder="Selecione ou digite…",
            key="om_choice"
        )
        om_sigla = disp2sig.get(om_display, "")
        if om_display == "Outra / não listada…":
            om_sigla = st.text_input("Sigla da OM (manual)", "")
            diretoria = st.text_input("Diretoria responsável (manual)", key="diretoria_manual")
        else:
            diretoria = sig2dir.get(om_sigla, "")
            st.text_input("Diretoria responsável", value=diretoria, disabled=True, key="diretoria_auto")
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
    if om_display == "Outra / não listada…":
        if not diretoria.strip():
            erros.append("Informe a **diretoria** (selecione uma OM conhecida ou preencha manualmente).")
    else:
        if not diretoria.strip():
            erros.append("Diretoria não encontrada para a OM selecionada.")
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
