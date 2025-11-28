# core/data_loader.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import time
from typing import Optional

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# =========================================================
# Configurações e escopos de acesso
# =========================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# =========================================================
# Utilidades internas
# =========================================================
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
    """Verifica se as credenciais do Google Sheets estão disponíveis..."""
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and st.secrets["gsheets"].get("spreadsheet_url")
    )


# =========================================================
# Conexão e autenticação
# =========================================================
@st.cache_resource(show_spinner=False, ttl=3600)
def _client() -> gspread.Client:
    """Cria cliente autenticado do Google Sheets (cached)."""
    if not has_gsheets():
        raise ValueError("⚠️ Credenciais do Google Sheets não configuradas em st.secrets")

    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False, ttl=3600)
def _book() -> gspread.Spreadsheet:
    """Retorna a planilha principal (cached)."""
    url = st.secrets["gsheets"]["spreadsheet_url"]
    return _client().open_by_url(url)


# =========================================================
# Leitura de planilhas
# =========================================================
@st.cache_data(ttl=300, show_spinner=False)
def read_df(tab_name: str, use_cache: bool = True) -> pd.DataFrame:
    """
    Lê uma aba do Sheets como DataFrame de forma robusta.
    - Detecta linha de cabeçalho automaticamente.
    - Normaliza colunas e remove vazios.
    """
    try:
        ws = _book().worksheet(tab_name)
        values = ws.get_all_values()

        if not values:
            return pd.DataFrame()

        # Detecta linha de cabeçalho (primeira linha não vazia)
        header_row_idx = next(
            (i for i, row in enumerate(values) if any(str(c).strip() for c in row)), None
        )
        if header_row_idx is None:
            return pd.DataFrame()

        raw_header = values[header_row_idx]
        width = max(len(r) for r in values[header_row_idx:]) or len(raw_header)

        # Header normalizado e único
        norm_header = _make_unique_headers([(h or "").strip() for h in raw_header])
        if len(norm_header) < width:
            norm_header += [f"col_{i}" for i in range(len(norm_header), width)]

        # Corpo da planilha
        body = [
            (row[:width] + [""] * max(0, width - len(row)))
            for row in values[header_row_idx + 1 :]
        ]
        if not body:
            return pd.DataFrame(columns=norm_header)

        df = pd.DataFrame(body, columns=norm_header)

        # Limpeza de strings e vazios
        for c in df.columns:
            df[c] = df[c].map(lambda x: x.strip() if isinstance(x, str) else x)
        df = df.replace({"": pd.NA, "—": pd.NA, "–": pd.NA, "-": pd.NA}).dropna(how="all")

        return df

    except gspread.WorksheetNotFound:
        st.warning(f"⚠️ Aba '{tab_name}' não encontrada no Google Sheets.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erro ao ler aba '{tab_name}': {e}")
        return pd.DataFrame()


# =========================================================
# Escrita / atualização de planilhas
# =========================================================
def overwrite_tab_from_df(
    tab_name: str,
    df: pd.DataFrame,
    keep_header: bool = True,
    batch_size: int = 1000,
):
    """Sobrescreve a aba inteira com o DataFrame informado."""
    sh = _book()
    try:
        ws = sh.worksheet(tab_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(
            title=tab_name,
            rows=max(2000, len(df) + 50),
            cols=max(10, len(df.columns)),
        )

    # Monta os valores a serem gravados
    values = (
        [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist()
        if keep_header
        else df.fillna("").astype(str).values.tolist()
    )

    try:
        for i in range(0, len(values), batch_size):
            chunk = values[i : i + batch_size]
            ws.update(f"A{i+1}", chunk, value_input_option="USER_ENTERED")
            time.sleep(0.5)
        clear_caches()
    except Exception as e:
        st.error(f"❌ Erro ao gravar na aba '{tab_name}': {e}")
        raise


def append_row(tab_name: str, row: dict):
    """Adiciona uma linha respeitando o cabeçalho existente."""
    sh = _book()
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        headers = list(row.keys())
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=len(headers))
        ws.update("1:1", [headers], value_input_option="USER_ENTERED")

    headers = ws.row_values(1)
    payload = [str(row.get(h, "")) for h in headers]

    try:
        ws.append_row(payload, value_input_option="USER_ENTERED")
        clear_caches()
    except Exception as e:
        st.error(f"❌ Erro ao adicionar linha na aba '{tab_name}': {e}")
        raise


def clear_caches():
    """Limpa os caches de leitura."""
    try:
        read_df.clear()
    except Exception:
        pass
    try:
        st.cache_resource.clear()
    except Exception:
        pass


# Alias para compatibilidade
overwrite_tab = overwrite_tab_from_df

__all__ = [
    "read_df",
    "append_row",
    "overwrite_tab_from_df",
    "overwrite_tab",
    "clear_caches",
    "has_gsheets",
]
