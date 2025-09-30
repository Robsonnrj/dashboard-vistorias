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

@st.cache_data(ttl=300, show_spinner=False)
def read_df(tab_name: str) -> pd.DataFrame:
    """Lê a aba inteira a partir da linha 1 como cabeçalho (tolerante a vazios)."""
    ws = _book().worksheet(tab_name)
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    header = values[0] if values else []
    data = values[1:] if len(values) > 1 else []
    df = pd.DataFrame(data, columns=header).replace("", pd.NA)
    # normaliza datas
    for c in df.columns:
        if "DATA" in c.upper():
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

def append_row(tab_name: str, header: list[str], row_dict: dict):
    """Inclui uma linha na aba; cria aba/cabeçalho se necessário."""
    ws = _ensure_ws(tab_name, header)
    row = [row_dict.get(h, "") for h in header]
    ws.append_row(row, value_input_option="USER_ENTERED")
    read_df.clear()   # invalida o cache dessa leitura
