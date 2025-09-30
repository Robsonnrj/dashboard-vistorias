# features/cadastro.py
# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import datetime
import pandas as pd
import streamlit as st

from core.data_loader import read_df, append_row


# --------------------- helpers ---------------------
def _norm(s: str) -> str:
    return str(s or "").strip().lower()


def _pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Tenta achar uma coluna pelo nome ou por 'parecido'."""
    cols = list(df.columns)
    # exata
    for want in candidates:
        for c in cols:
            if _norm(c) == _norm(want):
                return c
    # contém
    for want in candidates:
        w = _norm(want)
        for c in cols:
            if w in _norm(c):
                return c
    return None


def _load_oms_df() -> pd.DataFrame:
    """
    Carrega lista de OMs e suas diretorias de:
      1) aba de validação (st.session_state['tabs_map']['validacao']), se existir
      2) senão, da aba base de solicitações (ACOMPANHAMENTO VISTORIAS).
    Retorna DF com colunas padronizadas: ['om_sigla','om_nome','diretoria'] (faltantes viram "")
    """
    tabs_map = st.session_state.get("tabs_map", {})
    tab_valid = tabs_map.get("validacao")
    tab_base = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")

    df = pd.DataFrame()

    # 1) tentar validação
    if tab_valid:
        try:
            dv = read_df(tab_valid)
            if not dv.empty:
                c_sig = _pick_col(dv, ["sigla", "om", "om_sigla"])
                c_nom = _pick_col(dv, ["organização militar", "organização", "om_nome", "nome"])
                c_dir = _pick_col(dv, ["diretoria responsável", "diretoria", "dr"])
                if c_sig:
                    df = dv[[c for c in [c_sig, c_nom, c_dir] if c in dv.columns]].copy()
                    df.columns = [("om_sigla" if i == 0 else ("om_nome" if (i == 1 and c_nom) else "diretoria"))
                                  for i, _ in enumerate(df.columns)]
        except Exception:
            pass

    # 2) fallback na aba base
    if df.empty:
        try:
            db = read_df(tab_base)
            if not db.empty:
                c_sig = _pick_col(db, ["om", "om apoiada", "om_solicitante"])
                c_dir = _pick_col(db, ["diretoria responsável", "diretoria"])
                df = pd.DataFrame({
                    "om_sigla": db[c_sig] if c_sig in db.columns else "",
                    "om_nome":  "",
                    "diretoria": db[c_dir] if c_dir in db.columns else "",
                })
        except Exception:
            pass

    if df.empty:
        return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])

    # limpar NA e duplicados
    for c in ["om_sigla", "om_nome", "diretoria"]:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].map(lambda x: "" if pd.isna(x) else str(x).strip())

    df = df.drop_duplicates(subset=["om_sigla"]).sort_values("om_sigla").reset_index(drop=True)
    return df[["om_sigla", "om_nome", "diretoria"]]


def _input_row(oms_df: pd.DataFrame) -> tuple[dict, bool]:
    """Coleta os campos do formulário e devolve dicionário pronto para gravar + flag de validação."""
    st.subheader("📥 Nova solicitação de vistoria")

    # --- opções de OM com autocomplete + mapeamentos
    options: list[str] = []
    label_to_sigla: dict[str, str] = {}
    sigla_to_dir: dict[str, str] = {}

    for _, r in oms_df.iterrows():
        sig = "" if pd.isna(r.get("om_sigla")) else str(r.get("om_sigla")).strip()
        nom = "" if pd.isna(r.get("om_nome")) else str(r.get("om_nome")).strip()
        dire = "" if pd.isna(r.get("diretoria")) else str(r.get("diretoria")).strip()
        if not sig:
            continue
        label = sig
        if nom:
            label += f" — {nom}"
        if dire:
            label += f"  ({dire})"
        options.append(label)
        label_to_sigla[label] = sig
        sigla_to_dir[sig] = dire

    col1, col2 = st.columns(2)

    with col1:
        om_label = st.selectbox(
            "OM solicitante (sigla)",
            options,
            index=None,
            placeholder="Digite/seleciona a OM…",
        )
        om_solicitante = label_to_sigla.get(om_label, "")  # sigla limpa
        tipo_vistoria = st.selectbox(
            "Tipo de vistoria",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
        )

    with col2:
        local = st.text_input("Local / instalação")
        urgencia = st.selectbox("Urgência", ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"], index=0)

    # Diretoria preenchida automaticamente pela OM escolhida
    auto_dir = sigla_to_dir.get(om_solicitante, "")
    diretoria = st.text_input("Diretoria responsável", value=auto_dir, disabled=True)

    data_limite = st.date_input("Data limite (se houver)", value=None)
    motivo = st.text_area("Motivo / justificativa (NAOM)", height=120)

    # validações
    erros = []
    if not om_solicitante:
        erros.append("Selecione a **OM solicitante**.")
    if not local.strip():
        erros.append("Informe o **local/instalação**.")
    if not motivo.strip():
        erros.append("Descreva o **motivo/justificativa**.")

    if erros:
         st.warning("⚠️ Corrija os campos:\n• " + "\n• ".join(erros))
    

    row = {
        # campos operacionais próprios do sistema
        "numero": "",  # preenchido no salvar
        "data_solicitacao": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "om_solicitante": om_solicitante,
        "diretoria": diretoria,
        "tipo_vistoria": tipo_vistoria,
        "local": local.strip(),
        "urgencia": urgencia,
        "data_limite": data_limite.strftime("%Y-%m-%d") if data_limite else "",
        "motivo": motivo.strip(),
        "status_atual": "SOLICITADA",

        # mapeamento para a aba ACOMPANHAMENTO VISTORIAS
        "OM APOIADA": om_solicitante,
        "Diretoria Responsável": diretoria,
    }
    return row, (len(erros) == 0)


# --------------------- página ---------------------
def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    tabs_map = st.session_state.get("tabs_map", {})
    tab_solic = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")

    # Carrega para gerar próximo número
    try:
        df_exist = read_df(tab_solic)
    except Exception:
        df_exist = pd.DataFrame()

    # Carregar lista de OMs
    oms_df = _load_oms_df()

    # Formulário
    row, ok = _input_row(oms_df)

    # Salvar
row, ok = _input_row(oms_df)

# Salvar (sem desabilitar o botão — validamos por dentro)
if st.button("💾 Salvar solicitação", type="primary"):
    if not ok:
        st.warning("⚠️ Preencha os campos obrigatórios antes de salvar.")
    else:
        try:
            # número sequencial simples
            proximo = 1
            if not df_exist.empty and "numero" in df_exist.columns:
                try:
                    nums = pd.to_numeric(df_exist["numero"], errors="coerce").dropna()
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
    if not df_exist.empty:
        # Ordena por primeira coluna de data encontrada
        c_data = None
        for c in df_exist.columns:
            if "data" in c.lower():
                c_data = c
                break
        if c_data:
            df_exist[c_data] = pd.to_datetime(df_exist[c_data], errors="coerce")
            df_exist = df_exist.sort_values(c_data, ascending=False)
        st.dataframe(df_exist.head(50), use_container_width=True, height=360)
    else:
        st.caption("Ainda não há registros.")
