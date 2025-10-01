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
    """
    Tenta carregar a lista oficial de OMs e diretorias.
    1) Prioriza a aba 'Validacao_de_Dados' (se existir).
    2) Se não tiver, cai para 'ACOMPANHAMENTO VISTORIAS'.
    Normaliza para as colunas: ['om_sigla','om_nome','diretoria'].
    """
    # 1) origem preferencial
    for tab in ("Validacao_de_Dados", "ACOMPANHAMENTO VISTORIAS"):
        try:
            df = read_df(tab)
        except Exception:
            df = pd.DataFrame()

        if df.empty:
            continue

        # normalização de nomes possíveis
        cols = {c.lower().strip(): c for c in df.columns}

        # tenta mapear as colunas para sigla/nome/diretoria
        sigla = next((cols[k] for k in cols if "sigla" in k or k in ("om", "om apoiada")), None)
        nome  = next((cols[k] for k in cols if "organiza" in k or "om" == k), None)
        diret = next((cols[k] for k in cols if "diretoria" in k), None)

        # caso na aba de acompanhamento: muitas vezes a sigla já está em "OM"
        if not sigla and "OM" in df.columns:
            sigla = "OM"

        if not diret:
            # sem diretoria não serve para o autocomplete
            continue

        out = pd.DataFrame({
            "om_sigla": df[sigla] if sigla in df.columns else pd.Series(dtype=str),
            "om_nome":  df[nome]  if nome  in df.columns else pd.Series(dtype=str),
            "diretoria": df[diret],
        }).copy()

        # higieniza
        for c in ("om_sigla", "om_nome", "diretoria"):
            if c in out.columns:
                out[c] = out[c].fillna("").astype(str).str.strip()

        out = out[out["diretoria"] != ""].drop_duplicates()
        if not out.empty:
            return out

    # fallback vazio
    return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])


def _build_om_options(oms_df: pd.DataFrame):
    """
    Monta opções de exibição e dicionários de lookup.
    Retorna:
      options_display: lista de strings mostradas no selectbox
      disp_to_sigla:  display -> sigla
      sigla_to_dir:   sigla  -> diretoria
    """
    options_display, disp_to_sigla, sigla_to_dir = [], {}, {}

    if not oms_df.empty:
        for _, r in oms_df.iterrows():
            sig = str(r.get("om_sigla", "") or "").strip()
            nom = str(r.get("om_nome", "") or "").strip()
            dire = str(r.get("diretoria", "") or "").strip()
            if not sig:  # evita NA/None
                continue
            display = f"{sig} — {nom}" if nom else sig
            options_display.append(display)
            disp_to_sigla[display] = sig
            # guarda a 1ª diretoria encontrada para a sigla
            sigla_to_dir.setdefault(sig, dire)

    # opção manual (fora da lista)
    options_display.append("Outra / não listada…")
    disp_to_sigla["Outra / não listada…"] = ""

    return options_display, disp_to_sigla, sigla_to_dir


# -------------------------------
# Formulário
# -------------------------------
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
        )
        om_sigla = disp2sig.get(om_display, "")

        # diretoria automática se OM veio da lista
        diretoria_auto = sig2dir.get(om_sigla, "")

        # quando "Outra / não listada…", pede manualmente
        if om_sigla == "":
            om_sigla = st.text_input("Sigla da OM (manual)", "")
            diretoria = st.text_input("Diretoria responsável (manual)", "")
        else:
            diretoria = st.text_input(
                "Diretoria responsável (auto)",
                value=diretoria_auto,
                disabled=True
            )

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

    # Validação simples
    erros = []
    if not om_sigla.strip():
        erros.append("Informe a **OM**.")
    if not diretoria.strip():
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


# -------------------------------
# Página
# -------------------------------
def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    # 🔹 Agora fixo na aba ACOMPANHAMENTO VISTORIAS
    
    tab_solic = "ACOMPANHAMENTO VISTORIAS"

    # DataFrame existente (só p/ mostrar últimos)
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
            # numeração sequencial simples
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
