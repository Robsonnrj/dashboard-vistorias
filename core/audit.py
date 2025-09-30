# -*- coding: utf-8 -*-
import pandas as pd
from .data_loader import read_df, append_row, overwrite_tab_from_df  # <- nome correto

# --- helpers simples para achar colunas, tolerante a variações ---
def _norm(s: str) -> str:
    import unicodedata
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    return s.strip().lower()

def _pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    if df is None or df.empty:
        return None
    # match exato (normalizado)
    for cand in candidates:
        for c in df.columns:
            if _norm(c) == _norm(cand):
                return c
    # contém
    for cand in candidates:
        tgt = _norm(cand)
        for c in df.columns:
            if tgt in _norm(c):
                return c
    return None

# --- trilha de auditoria -----------------------------------------------------
def registrar_historico(row: dict):
    """
    Adiciona um evento na aba 'historicos' (criada automaticamente pelo data_loader).
    """
    append_row("historicos", row)

def trilha(numero: str) -> pd.DataFrame:
    """
    Retorna o histórico de auditoria para um número de solicitação.
    """
    try:
        hist = read_df("historicos")
        if hist.empty:
            return pd.DataFrame()
        col_num = _pick_col(hist, ["numero", "número", "num", "nº", "id", "protocolo"])
        if not col_num:
            return pd.DataFrame()
        return hist[hist[col_num].astype(str) == str(numero)].sort_index(ascending=False)
    except Exception:
        return pd.DataFrame()

# --- atualização de status ----------------------------------------------------
def atualizar_status(numero: str, novo_status: str, justificativa: str, responsavel: str):
    """
    Atualiza o status da solicitação na aba base e registra a auditoria.
    - Busca a aba base em st.session_state['tabs_map']['solicitacoes'] (configurada no app).
    - Escreve de volta a aba inteira com overwrite_tab_from_df (nome correto no data_loader).
    """
    import streamlit as st
    from datetime import datetime

    # 1) Carrega a aba base (solicitações)
    tab_base = st.session_state["tabs_map"]["solicitacoes"]
    df = read_df(tab_base)
    if df.empty:
        raise RuntimeError("A aba base de solicitações está vazia.")

    # 2) Descobre colunas
    col_num = _pick_col(df, ["numero", "número", "num", "nº", "id", "protocolo"])
    col_stt = _pick_col(df, ["status_atual", "status", "situação", "situacao"])

    if not col_num:
        raise RuntimeError("Não foi possível localizar a coluna 'número' da solicitação.")
    if not col_stt:
        raise RuntimeError("Não foi possível localizar a coluna de 'status' na planilha.")

    # 3) Localiza a linha e atualiza
    mask = df[col_num].astype(str) == str(numero)
    if not mask.any():
        raise RuntimeError(f"Solicitação '{numero}' não encontrada.")
    df.loc[mask, col_stt] = str(novo_status)

    # 4) Escreve a aba inteira de volta (usa o nome CORRETO da função)
    overwrite_tab_from_df(tab_base, df, keep_header=True)

    # 5) Registra auditoria
    registrar_historico({
        "numero": str(numero),
        "novo_status": str(novo_status),
        "justificativa": justificativa or "",
        "responsavel": responsavel or "",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
