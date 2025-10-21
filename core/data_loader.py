# core/data_loader.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import pandas as pd
import gspread
import streamlit as st
from google.oauth2.service_account import Credentials
from core.utils import pick_col

# ===== Google APIs =====
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ---------- Utilidades ----------
def _make_unique_headers(headers: list[str]) -> list[str]:
    """Gera cabeçalhos únicos (vazios -> col_X; duplicados -> nome_2, nome_3, …)."""
    seen: dict[str, int] = {}
    out: list[str] = []
    for j, h in enumerate(headers, start=1):
        h = (h or "").strip()
        if not h:
            h = f"col_{j}"
        base = h
        seen[base] = seen.get(base, 0) + 1
        if seen[base] > 1:
            h = f"{base}_{seen[base]}"
        out.append(h)
    return out


def has_gsheets() -> bool:
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and bool(st.secrets["gsheets"].get("spreadsheet_url"))
    )


# ---------- Conexão (cacheada) ----------
@st.cache_resource(show_spinner=False)
def _client() -> gspread.Client:
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def _book() -> gspread.Spreadsheet:
    return _client().open_by_url(st.secrets["gsheets"]["spreadsheet_url"])


def _ensure_ws(title: str, header: list[str]) -> gspread.Worksheet:
    """Garante worksheet existente com cabeçalho informado."""
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


# ---------- Leitura tolerante ----------
def _read_ws_loose(ws: gspread.Worksheet, header_row: int | None = None) -> pd.DataFrame:
    """
    Lê a worksheet tolerando cabeçalho vazio/duplicado/mesclado.
    - Se header_row não for dado, usa a primeira linha que tiver algo.
    """
    values = ws.get_all_values()  # lista de listas
    if not values:
        return pd.DataFrame()

    if header_row is None:
        # primeira linha com algum conteúdo
        hdr_idx = next((i for i, row in enumerate(values) if any(str(c).strip() for c in row)), 0)
    else:
        hdr_idx = max(0, int(header_row) - 1)

    headers = _make_unique_headers(values[hdr_idx])
    body = values[hdr_idx + 1 :]

    # remove linhas finais 100% vazias (opcional)
    while body and not any(str(c).strip() for c in body[-1]):
        body.pop()

    df = pd.DataFrame(body, columns=headers).replace("", pd.NA)
    return df


# ---------- API pública deste módulo ----------
@st.cache_data(ttl=60, show_spinner=False)
def read_df(tab_name: str) -> pd.DataFrame:
    """Lê uma aba do Sheets como DataFrame (com leitura tolerante + cache)."""
    ws = _book().worksheet(tab_name)
    df = _read_ws_loose(ws)

    # normalização de datas: qualquer coluna que contenha "DATA"
    for c in list(df.columns):
        if "DATA" in c.upper():
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def overwrite_tab_from_df(tab_name: str, df: pd.DataFrame, keep_header: bool = True) -> None:
    """Sobrescreve a aba inteira com o DataFrame fornecido."""
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

    values = (
        [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist()
        if keep_header
        else df.fillna("").astype(str).values.tolist()
    )
    ws.update("A1", values, value_input_option="USER_ENTERED")

    # invalida cache de leitura
    read_df.clear()


def append_row(tab_name: str, row: dict):
    """Acrescenta uma linha garantindo que todas as chaves do dict existam no cabeçalho."""
    sh = _book()
    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        # cria com cabeçalho vindo do dict
        headers = list(row.keys())
        ws = sh.add_worksheet(title=tab_name, rows=2000, cols=max(10, len(headers)))
        ws.update("1:1", [headers])
    else:
        # garante que o cabeçalho tenha TODAS as chaves do dict
        headers = ws.row_values(1)
        missing = [k for k in row.keys() if k not in headers]
        if missing:
            new_headers = headers + missing
            ws.resize(rows=ws.row_count, cols=max(ws.col_count, len(new_headers)))
            ws.update("1:1", [new_headers])
            headers = new_headers

    # monta a linha respeitando a ordem do cabeçalho
    payload = [row.get(h, "") for h in headers]
    ws.append_row(payload, value_input_option="USER_ENTERED")

    # invalida cache de leitura, se existir
    try:
        read_df.clear()   # st.cache_data clear
    except Exception:
        pass



def clear_caches() -> None:
    """Limpa os caches usados por este módulo."""
    try:
        read_df.clear()
    except Exception:
        pass
    try:
        st.cache_resource.clear()
    except Exception:
        pass


# Alias para compatibilidade com nomes antigos:
overwrite_tab = overwrite_tab_from_df

__all__ = ["has_gsheets", "read_df", "append_row", "overwrite_tab_from_df", "overwrite_tab", "clear_caches"]
