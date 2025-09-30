# core/data_loader.py
from __future__ import annotations

import unicodedata
from typing import Iterable, List, Dict

import pandas as pd
import gspread
import streamlit as st
from google.oauth2.service_account import Credentials

# ===== Config =====
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Mapeie aqui as abas “lógicas” do sistema para os títulos reais no Sheets
SHEET_TABS = {
    "solicitacoes": "ACOMPANHAMENTO VISTORIAS",
    "validacao": "Validacao_de_Dados",
    "auditoria": "AUDITORIA_VISTORIAS",
}

REQUIRED_HEADERS = {
    # cabeçalhos mínimos (ajuste conforme sua planilha)
    "solicitacoes": [
        "OBJETO DE VISTORIA",
        "OM APOIADA",
        "Diretoria Responsável",
        "Classificação de Urgência",
        "Situação",
        "DATA DA SOLICITAÇÃO",
    ],
}

# ===== Helpers =====
def _norm(s: str) -> str:
    s = str(s or "")
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.strip().lower()

def has_gsheets() -> bool:
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and st.secrets["gsheets"].get("spreadsheet_url")
    )

@st.cache_resource(show_spinner=False)
def _client() -> gspread.Client:
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def _book():
    return _client().open_by_url(st.secrets["gsheets"]["spreadsheet_url"])

def _ensure_ws(title: str, headers: list[str] | None = None):
    """Garante que a aba exista; se headers for dado, garante a linha 1."""
    sh = _book()
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        # cria com um mínimo de linhas/colunas
        cols = max(10, (len(headers) if headers else 0))
        ws = sh.add_worksheet(title=title, rows=2000, cols=cols or 10)
        if headers:
            ws.update("1:1", [headers])
        return ws

    if headers:
        row1 = ws.row_values(1)
        if row1 != headers:
            # só atualiza se estiver diferente (evita quota desnecessária)
            ws.update("1:1", [headers])
    return ws

def _make_unique_headers(raw_headers: Iterable[str]) -> list[str]:
    out, seen = [], {}
    for j, h in enumerate(raw_headers, start=1):
        h = (h or "").strip()
        if not h:
            h = f"col_{j}"
        base = h
        if base in seen:
            seen[base] += 1
            h = f"{base}_{seen[base]}"
        else:
            seen[base] = 1
        out.append(h)
    return out

def _read_ws_loose(ws) -> pd.DataFrame:
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    # acha primeira linha não totalmente vazia como cabeçalho
    hdr_idx = next((i for i, row in enumerate(values) if any(c.strip() for c in row)), 0)
    headers = _make_unique_headers(values[hdr_idx])
    body = values[hdr_idx + 1 :]
    # remove rodapé vazio
    while body and not any((c or "").strip() for c in body[-1]):
        body.pop()
    df = pd.DataFrame(body, columns=headers).replace("", pd.NA)

    # normaliza datas por nome
    for c in df.columns:
        if "data" in _norm(c):
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

# ===== API pública usada pelo app =====
@st.cache_data(ttl=60, show_spinner=False)
def read_df(tab_key: str) -> pd.DataFrame:
    """Lê uma aba do Sheets, tolerante a cabeçalhos imperfeitos."""
    assert has_gsheets(), "Google Sheets OFF (secrets ausente)"
    title = SHEET_TABS.get(tab_key, tab_key)
    need_headers = REQUIRED_HEADERS.get(tab_key)
    ws = _ensure_ws(title, headers=need_headers)
    # se temos REQUIRED_HEADERS, podemos ler por get_all_records(); senão, leitura 'loose'
    if need_headers:
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        # datas
        for c in df.columns:
            if "data" in _norm(c):
                df[c] = pd.to_datetime(df[c], errors="coerce")
        return df
    else:
        return _read_ws_loose(ws)

def overwrite_tab(tab_key: str, df: pd.DataFrame, keep_header: bool = True) -> None:
    """Sobrescreve completamente a aba com o DataFrame informado."""
    assert has_gsheets(), "Google Sheets OFF (secrets ausente)"
    title = SHEET_TABS.get(tab_key, tab_key)
    ws = _ensure_ws(title, headers=(list(df.columns) if keep_header else None))
    ws.clear()
    if keep_header:
        values = [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist()
    else:
        values = df.fillna("").astype(str).values.tolist()
    ws.update("A1", values, value_input_option="USER_ENTERED")
    read_df.clear()

def append_row(tab_key: str, row_dict: Dict[str, str | int | float]) -> None:
    """Acrescenta uma linha respeitando os headers definidos para a aba."""
    assert has_gsheets(), "Google Sheets OFF (secrets ausente)"
    title = SHEET_TABS.get(tab_key, tab_key)
    headers = REQUIRED_HEADERS.get(tab_key)
    if not headers:
        # sem cabeçalho definido, não sabemos a ordem
        raise ValueError(f"Não há REQUIRED_HEADERS para '{tab_key}'.")
    ws = _ensure_ws(title, headers=headers)
    row = [row_dict.get(h, "") for h in headers]
    ws.append_row(row, value_input_option="USER_ENTERED")
    read_df.clear()
