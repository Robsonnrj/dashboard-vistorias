# features/cadastro.py
# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime
import pandas as pd
from core.data_loader import append_row, read_df


def _input_row():
    """Coleta os campos do formulário e devolve um dicionário pronto para gravar."""
    st.subheader("📥 Nova solicitação de vistoria")
    col1, col2 = st.columns(2)

    with col1:
        om_solicitante = st.text_input("OM solicitante (sigla)")
        diretoria = st.text_input("Diretoria responsável")
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

    # Validações simples
    erros = []
    if not om_solicitante.strip():
        erros.append("Informe a **OM solicitante**.")
    if not local.strip():
        erros.append("Informe o **local/instalação**.")
    if not motivo.strip():
        erros.append("Descreva o **motivo/justificativa**.")

    if erros:
        st.warning("• " + "\n• ".join(erros))

    row = {
        "numero": "",  # será atribuído ao salvar (sequencial)
        "data_solicitacao": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "om_solicitante": om_solicitante.strip(),
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

    # Abas configuradas na sidebar (já existentes no seu app)
    tabs_map = st.session_state.get("tabs_map", {})
    tab_solic = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")

    # Carrega para mostrar as últimas solicitações (não é obrigatório para salvar)
    try:
        df_existente = read_df(tab_solic)
    except Exception:
        df_existente = pd.DataFrame()

    # Formulário
    row, ok = _input_row()

    # Salvar
    if st.button("💾 Salvar solicitação", disabled=not ok):
        try:
            # Gera número sequencial simples baseado na quantidade atual
            proximo = 1
            if not df_existente.empty and "numero" in df_existente.columns:
                try:
                    # tenta converter para inteiro ignorando vazios
                    nums = pd.to_numeric(df_existente["numero"], errors="coerce").dropna()
                    if not nums.empty:
                        proximo = int(nums.max()) + 1
                except Exception:
                    pass
            row["numero"] = str(proximo)

            append_row(tab_solic, row)
            st.success(f"Solicitação **#{row['numero']}** cadastrada com sucesso!")
            st.rerun()
        except Exception as e:
            st.error(f"Falha ao salvar: {e}")

    st.divider()
    st.subheader("📄 Últimas solicitações")
    if not df_existente.empty:
        # Ordena por data se existir
        c_data = None
        for c in df_existente.columns:
            if "data" in c.lower():
                c_data = c
                break
        if c_data:
            df_existente[c_data] = pd.to_datetime(df_existente[c_data], errors="coerce")
            df_existente = df_existente.sort_values(c_data, ascending=False)

        st.dataframe(df_existente.head(50), use_container_width=True, height=360)
    else:
        st.caption("Ainda não há registros.")
