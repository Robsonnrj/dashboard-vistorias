# -*- coding: utf-8 -*-
from __future__ import annotations

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import streamlit as st

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _make_unique_headers(headers: list[str]) -> list[str]:
    seen, out = {}, []
    for h in headers:
        h = (h or "").strip() or "col"
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 1
        out.append(h)
    return out


def has_gsheets() -> bool:
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and st.secrets["gsheets"].get("spreadsheet_url")
    )


@st.cache_resource(show_spinner=False)
def _client():
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def _book():
    return _client().open_by_url(st.secrets["gsheets"]["spreadsheet_url"])


def _ensure_ws(title: str, header: list[str]):
    sh = _book()
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=2000, cols=max(10, len(header)))
        ws.update("1:1", [header])
        return ws
    if ws.row_values(1) != header:
        ws.update("1:1", [header])
    return ws


@st.cache_data(ttl=300, show_spinner=False)
def read_df(tab_name: str) -> pd.DataFrame:
    """Lê uma aba do Sheets como DataFrame (tolerante)."""
    ws = _book().worksheet(tab_name)
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    headers = _make_unique_headers(values[0])
    df = pd.DataFrame(values[1:], columns=headers).replace("", pd.NA)
    return df


def overwrite_tab_from_df(tab_name: str, df: pd.DataFrame, keep_header: bool = True):
    """Sobrescreve a aba inteira com o DataFrame."""
    sh = _book()
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(
            title=tab_name,
            rows=max(2000, len(df) + 10),
            cols=max(10, len(df.columns)),
        )
    else:
        ws.clear()

    if keep_header:
        values = [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist()
    else:
        values = df.fillna("").astype(str).values.tolist()

    ws.update("A1", values, value_input_option="USER_ENTERED")

    # invalida cache de leitura
    try:
        read_df.clear()
    except Exception:
        pass


def append_row(tab_name: str, row: dict):
    """Acrescenta uma linha (dict) respeitando o cabeçalho existente."""
    sh = _book()
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        headers = list(row.keys())
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=max(10, len(headers)))
        ws.update("1:1", [headers])

    headers = ws.row_values(1)
    payload = [row.get(h, "") for h in headers]
    ws.append_row(payload, value_input_option="USER_ENTERED")

    try:
        read_df.clear()
    except Exception:
        pass


def clear_caches():
    """Limpa somente os caches deste módulo."""
    try:
        read_df.clear()
    except Exception:
        pass


# retrocompatibilidade
overwrite_tab = overwrite_tab_from_df
__all__ = ["read_df", "append_row", "overwrite_tab_from_df", "overwrite_tab", "clear_caches", "has_gsheets"]
