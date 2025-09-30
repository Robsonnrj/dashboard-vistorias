# -*- coding: utf-8 -*-
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
import streamlit as st

from .config import SCOPES, SHEET_TABS, has_gsheets

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
    # garante cabeçalho
    if ws.row_values(1) != header:
        ws.update("1:1", [header])
    return ws

# ---------- CRUD ----------
def read_df(tab_key: str) -> pd.DataFrame:
    title = SHEET_TABS[tab_key]
    ws = _ensure_ws(title, _headers_for(tab_key))
    rows = ws.get_all_records()
    return pd.DataFrame(rows)

def append_row(tab_key: str, row_dict: dict):
    title = SHEET_TABS[tab_key]
    ws = _ensure_ws(title, _headers_for(tab_key))
    row = [row_dict.get(h, "") for h in _headers_for(tab_key)]
    ws.append_row(row, value_input_option="USER_ENTERED")

def overwrite_tab(tab_key: str, df: pd.DataFrame):
    title = SHEET_TABS[tab_key]
    ws = _ensure_ws(title, list(df.columns))
    ws.clear()
    ws.update("A1", [list(df.columns)] + df.fillna("").astype(str).values.tolist(),
              value_input_option="USER_ENTERED")

def _headers_for(tab_key: str) -> list[str]:
    if tab_key == "solicitacoes":
        return [
            "numero","om_solicitante","om_nome","diretoria","local","coordenadas",
            "tipo_vistoria","motivo","urgencia","data_limite","anexos",
            "status_atual","criado_por","criado_em"
        ]
    if tab_key == "historicos":
        return ["numero","status_de","status_para","justificativa","responsavel","timestamp"]
    if tab_key == "relatorios":
        return ["numero","titulo","arquivo_pdf","gerado_por","gerado_em"]
    return []
