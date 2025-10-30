# core/data_loader.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import time
from typing import Optional

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _make_unique_headers(headers: list[str]) -> list[str]:
    """Cria headers únicos para evitar colunas duplicadas."""
    seen, out = {}, []
    for h in headers:
        h = (h or "").strip() or "col"
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        out.append(h)
    return out


def has_gsheets() -> bool:
    """Verifica se as credenciais do Google Sheets estão disponíveis."""
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and st.secrets["gsheets"].get("spreadsheet_url")
    )


@st.cache_resource(show_spinner=False, ttl=3600)  # Cache de 1 hora
def _client():
    """Cria cliente autenticado do Google Sheets (cached)."""
    if not has_gsheets():
        raise ValueError("Credenciais do Google Sheets não configuradas em st.secrets")

    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False, ttl=3600)
def _book():
    """Retorna a planilha (workbook) do Google Sheets (cached)."""
    url = st.secrets["gsheets"]["spreadsheet_url"]
    return _client().open_by_url(url)


def _ensure_ws(title: str, header: list[str]):
    """Garante que a worksheet existe e tem o header correto."""
    sh = _book()
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=2000, cols=max(10, len(header)))
        ws.update("1:1", [header], value_input_option="USER_ENTERED")
        return ws

    # Verifica se o header está correto
    existing_header = ws.row_values(1)
    if existing_header != header:
        ws.update("1:1", [header], value_input_option="USER_ENTERED")

    return ws


@st.cache_data(ttl=300, show_spinner=False)  # Cache de 5 minutos
def read_df(tab_name: str, use_cache: bool = True) -> pd.DataFrame:
    """
    Lê uma aba do Sheets como DataFrame, de forma robusta.
    - Detecta automaticamente a linha de header (1ª linha não totalmente vazia).
    - Mantém todas as colunas (faz padding nas linhas mais curtas).
    - Converte "", "—", "–" e "-" isolado para NA.
    - Faz strip de espaços em headers e valores.
    - Remove linhas totalmente vazias após a limpeza.

    Args:
        tab_name: Nome da aba
        use_cache: Mantido só para compor a chave do cache. Para forçar recarga,
                   chame read_df(tab_name, use_cache=False).
    """
    try:
        ws = _book().worksheet(tab_name)
        values = ws.get_all_values()  # lista de listas

        if not values:
            return pd.DataFrame()

        # -------- localizar a linha de cabeçalho (primeira linha não totalmente vazia) --------
        def _is_all_empty(row: list[str]) -> bool:
            return all((c is None) or (str(c).strip() == "") for c in row)

        header_row_idx = None
        for i, row in enumerate(values):
            if not _is_all_empty(row):
                header_row_idx = i
                break
        if header_row_idx is None:
            return pd.DataFrame()  # planilha sem conteúdo útil

        raw_header = values[header_row_idx]

        # Largura máxima (para padronizar linhas curtas)
        width = max(len(r) for r in values[header_row_idx:]) if values[header_row_idx:] else len(raw_header)

        # Normalizar header: strip e headers únicos
        norm_header = [(h or "").strip() for h in raw_header]
        if len(norm_header) < width:
            norm_header += [""] * (width - len(norm_header))
        norm_header = _make_unique_headers(norm_header)

        # -------- montar corpo (linhas após o header) com padding --------
        body = []
        for row in values[header_row_idx + 1:]:
            r = row[:width] + [""] * max(0, width - len(row))
            body.append(r)
        if not body:
            return pd.DataFrame(columns=norm_header)

        df = pd.DataFrame(body, columns=norm_header)

        # -------- limpeza de strings --------
        for c in df.columns:
            df[c] = df[c].map(lambda x: x.strip() if isinstance(x, str) else x)

        # mapear vazios e traços para NA
        df = df.replace(to_replace={"": pd.NA, "—": pd.NA, "–": pd.NA, "-": pd.NA})

        # remover linhas totalmente vazias
        df = df.dropna(how="all")

        return df

    except gspread.WorksheetNotFound:
        st.warning(f"⚠️ Aba '{tab_name}' não encontrada no Google Sheets")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erro ao ler aba '{tab_name}': {e}")
        return pd.DataFrame()


def overwrite_tab_from_df(
    tab_name: str,
    df: pd.DataFrame,
    keep_header: bool = True,
    batch_size: int = 1000,
):
    """
    Sobrescreve a aba inteira com o DataFrame.
    """
    sh = _book()

    try:
        ws = sh.worksheet(tab_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(
            title=tab_name,
            rows=max(2000, len(df) + 100),
            cols=max(10, len(df.columns)),
        )

    # Prepara os dados
    if keep_header:
        values = [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist()
    else:
        values = df.fillna("").astype(str).values.tolist()

    # Upload em lotes para evitar timeout
    try:
        if len(values) <= batch_size:
            ws.update("A1", values, value_input_option="USER_ENTERED")
        else:
            for i in range(0, len(values), batch_size):
                batch = values[i : i + batch_size]
                start_row = i + 1
                ws.update(f"A{start_row}", batch, value_input_option="USER_ENTERED")
                time.sleep(0.5)  # Evita rate limiting

        # Limpa o cache
        clear_caches()

    except Exception as e:
        st.error(f"❌ Erro ao escrever na aba '{tab_name}': {e}")
        raise


def append_row(tab_name: str, row: dict):
    """
    Adiciona uma linha respeitando o cabeçalho existente.
    """
    sh = _book()

    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        # Cria a aba com o header do dict
        headers = list(row.keys())
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=max(10, len(headers)))
        ws.update("1:1", [headers], value_input_option="USER_ENTERED")

    # Pega o header atual
    headers = ws.row_values(1)

    # Monta a linha respeitando a ordem do header
    payload = [str(row.get(h, "")) for h in headers]

    # Adiciona a linha
    try:
        ws.append_row(payload, value_input_option="USER_ENTERED")
        clear_caches()
    except Exception as e:
        st.error(f"❌ Erro ao adicionar linha na aba '{tab_name}': {e}")
        raise


def clear_caches():
    """Limpa todos os caches de leitura."""
    try:
        read_df.clear()
    except Exception:
        pass


# Retrocompatibilidade
overwrite_tab = overwrite_tab_from_df

__all__ = [
    "read_df",
    "append_row",
    "overwrite_tab_from_df",
    "overwrite_tab",
    "clear_caches",
    "has_gsheets",
]
