from __future__ import annotations
import unicodedata
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import streamlit as st

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
def _make_unique_headers(headers: list[str]) -> list[str]:
    seen = {}
    result = []
    for h in headers:
        h = (h or "").strip()
        if not h:
            h = "col"
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 1
        result.append(h)
    return result
def has_gsheets() -> bool:
    return ("gcp_service_account" in st.secrets
            and "gsheets" in st.secrets
            and st.secrets["gsheets"].get("spreadsheet_url"))

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
def read_df(tab_name: str) -> pd.DataFrame:
    """Lê uma aba do Sheets como DataFrame."""
    ws = _book().worksheet(tab_name)
    # use sua função de leitura tolerante, por ex. read_ws_loose(ws)
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    df = pd.DataFrame(values[1:], columns=values[0]).replace("", pd.NA)
    # normalização de datas opcional…
    return df

def overwrite_tab_from_df(tab_name: str, df: pd.DataFrame, keep_header: bool = True):
    """Sobrescreve a aba inteira com o DataFrame."""
    sh = _book()
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=max(2000, len(df)+10), cols=max(10, len(df.columns)))
    else:
        ws.clear()
    values = [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist() if keep_header \
             else df.fillna("").astype(str).values.tolist()
    ws.update("A1", values, value_input_option="USER_ENTERED")
    # invalida cache, se você usa @st.cache_data em read_df:
    try:
        read_df.clear()
    except Exception:
        pass

def append_row(tab_name: str, row: dict):
    """Acrescenta uma linha (dict) mantendo a ordem do cabeçalho."""
    sh = _book()
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        # cria com o cabeçalho vindo da chave do dict
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
     """ Limpa somente os caches deste modulo."""
     try:
         read_DF.clear()
     except Exception:
         pass

# Alias para compatibilidade com nomes antigos:
overwrite_tab = overwrite_tab_from_df

# Opcional, ajuda a evitar import errors “fantasmas”
__all__ = ["read_df", "append_row", "overwrite_tab_from_df", "overwrite_tab"]


