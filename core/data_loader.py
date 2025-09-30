# core/data_loader.py
from _future_ import annotations
import unicodedata
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import streamlit as st

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def _norm(s: str) -> str:
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower().strip()

def _make_unique_headers(raw_headers):
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

@st.cache_resource(show_spinner=False)
def _client():
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def _book():
    return _client().open_by_url(st.secrets["gsheets"]["spreadsheet_url"])

def _read_ws_loose(ws, header_row=None) -> pd.DataFrame:
    """Lê tolerando cabeçalho repetido/mesclado/vazio e gera nomes únicos."""
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    # acha a linha do cabeçalho
    if header_row is None:
        hdr_idx = next((i for i, row in enumerate(values) if any(str(c).strip() for c in row)), 0)
    else:
        hdr_idx = max(0, int(header_row) - 1)

    headers = _make_unique_headers(values[hdr_idx])
    body = values[hdr_idx + 1:]

    # remove linhas finais totalmente vazias
    while body and not any(str(c).strip() for c in body[-1]):
        body.pop()

    df = pd.DataFrame(body, columns=headers).replace("", pd.NA)
    # tenta converter datas
    for c in df.columns:
        if "DATA" in c.upper():
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

def _ensure_ws(title: str, _headers_for: list[str] | None = None):
    """
    Garante que a worksheet exista.
    Se não houver, cria (com headers se fornecidos). Não exige cabeçalho igual para ler.
    """
    sh = _book()
    try:
        ws = sh.worksheet(title)
        return ws
    except gspread.WorksheetNotFound:
        rows = 2000
        cols = max(10, len(_headers_for or []))
        ws = sh.add_worksheet(title=title, rows=rows, cols=cols)
        if _headers_for:
            ws.update("1:1", [_headers_for])
        return ws

@st.cache_data(ttl=120, show_spinner=False)
def read_df(title: str, _headers_for: list[str] | None = None) -> pd.DataFrame:
    """
    Lê uma aba por título, criando caso não exista. Tolerante a cabeçalho fora do padrão.
    """
    try:
        ws = _ensure_ws(title, _headers_for=_headers_for)
        return _read_ws_loose(ws)
    except gspread.exceptions.APIError as e:
        # Mensagem amigável
        raise RuntimeError(
            "Falha ao acessar a planilha no Google Sheets.\n"
            "Verifique:\n"
            "• A URL em [gsheets.spreadsheet_url] no secrets.toml está correta;\n"
            "• A planilha foi compartilhada com o e-mail da service account (st.secrets['gcp_service_account']['client_email']) com permissão de EDITOR;\n"
            f"• A aba '{title}' existe (ou será criada) e não está protegida.\n\n"
            f"Detalhe técnico: {e}"
        )

def write_df(title: str, df: pd.DataFrame, keep_header=True):
    """Sobrescreve a aba com o DataFrame."""
    ws = _ensure_ws(title, _headers_for=list(df.columns) if keep_header else None)
    ws.clear()
    values = [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist() if keep_header \
             else df.fillna("").astype(str).values.tolist()
    ws.update("A1", values, value_input_option="USER_ENTERED")
    read_df.clear()
