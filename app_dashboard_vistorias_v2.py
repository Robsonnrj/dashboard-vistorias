# sheets_editor.py
# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime

st.set_page_config(page_title="Editor de Abas • Google Sheets", layout="wide")

# ========= Config / Constantes =========
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Cabeçalhos (caso queira usar o modo "acrescentar novas linhas" nas abas de vistorias/OMs)
OM_HEADERS = ["Nome", "Sigla", "Diretoria", "Criado em"]
VT_HEADERS = [
    "OBJETO DE VISTORIA",
    "OM APOIADA",
    "Diretoria Responsavel",
    "Classificacao da Urgencia",
    "Situacao",
    "DATA DA SOLICITACAO",
]

# ========= Helpers Google Sheets =========
def has_gsheets() -> bool:
    return (
        "gcp_service_account" in st.secrets and
        "gsheets" in st.secrets and
        "spreadsheet_url" in st.secrets["gsheets"]
    )

@st.cache_resource(show_spinner=False)
def _gs_client():
    import gspread
    from google.oauth2.service_account import Credentials
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def _book():
    return _gs_client().open_by_url(st.secrets["gsheets"]["spreadsheet_url"])

def _norm(s: str) -> str:
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower().strip()

def list_tabs():
    sh = _book()
    return [ws.title for ws in sh.worksheets()]

def read_tab_as_df(tab_name: str) -> pd.DataFrame:
    ws = _book().worksheet(tab_name)
    rows = ws.get_all_records()
    df = pd.DataFrame(rows)
    # tenta converter colunas com "DATA" no nome
    for c in df.columns:
        if "DATA" in c.upper():
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

def overwrite_tab_from_df(tab_name: str, df: pd.DataFrame, keep_header=True):
    """
    Sobrescreve toda a aba com o conteúdo do DF.
    - keep_header=True: usa as COLUNAS ATUAIS do DF como cabeçalho (linha 1).
    Observação: preserva o nome da aba; não duplica.
    """
    ws = _book().worksheet(tab_name)
    # garante que tudo seja string-friendly
    df_to_write = df.copy()
    # datas -> string ISO (Google Sheets entende bem)
    for c in df_to_write.columns:
        if pd.api.types.is_datetime64_any_dtype(df_to_write[c]):
            df_to_write[c] = df_to_write[c].dt.strftime("%Y-%m-%d")
    values = [df_to_write.columns.tolist()] + df_to_write.fillna("").astype(str).values.tolist()

    # Limpa e atualiza
    ws.clear()
    if keep_header:
        ws.update("1:1", [values[0]])
        if len(values) > 1:
            ws.update(f"A2", values[1:], value_input_option="USER_ENTERED")
    else:
        ws.update("A1", values, value_input_option="USER_ENTERED")

def append_new_rows_dedup(tab_name: str, df_new: pd.DataFrame, key_cols: list[str]):
    """
    Acrescenta somente as linhas NOVAS (com base nas colunas de chave 'key_cols').
    - Não apaga nada, não altera linhas existentes.
    - key_cols devem EXISTIR na aba de destino e no df_new.
    - Converte datas para 'YYYY-MM-DD'.
    """
    ws = _book().worksheet(tab_name)
    exist = ws.get_all_records()
    df_exist = pd.DataFrame(exist)

    # se a aba estiver vazia, cria com o cabeçalho do df_new
    if df_exist.empty:
        overwrite_tab_from_df(tab_name, df_new, keep_header=True)
        return len(df_new)

    # normaliza datas para comparação
    for c in df_exist.columns:
        if "DATA" in c.upper():
            df_exist[c] = pd.to_datetime(df_exist[c], errors="coerce")
    df_tmp = df_new.copy()
    for c in df_tmp.columns:
        if "DATA" in c.upper():
            df_tmp[c] = pd.to_datetime(df_tmp[c], errors="coerce")

    # valida chaves
    for k in key_cols:
        if k not in df_exist.columns or k not in df_tmp.columns:
            raise ValueError(f"Coluna de chave ausente: {k}")

    # cria set de chaves existentes
    def key_tuple(row, cols):
        out = []
        for k in cols:
            v = row.get(k)
            if isinstance(v, pd.Timestamp):
                out.append(str(v.date()) if pd.notna(v) else "")
            else:
                out.append(str(v).strip() if v is not None else "")
        return tuple(out)

    exist_keys = { key_tuple(df_exist.loc[i], key_cols) for i in df_exist.index }

    # prepara novas linhas
    df_out = []
    for _, r in df_tmp.iterrows():
        k = key_tuple(r, key_cols)
        if k in exist_keys:
            continue
        df_out.append(r)

    if not df_out:
        return 0

    df_out = pd.DataFrame(df_out)
    # respeita a ordem de colunas da aba (se possível)
    cols_order = [c for c in df_exist.columns if c in df_out.columns] + [c for c in df_out.columns if c not in df_exist.columns]
    df_out = df_out[cols_order]

    # conversão final para string
    for c in df_out.columns:
        if pd.api.types.is_datetime64_any_dtype(df_out[c]):
            df_out[c] = df_out[c].dt.strftime("%Y-%m-%d")
    values = df_out.fillna("").astype(str).values.tolist()

    # append em lote
    ws.append_rows(values, value_input_option="USER_ENTERED")
    return len(values)

# ========= UI =========
st.title("📄 Editor de Abas — Google Sheets")

if not has_gsheets():
    st.error("Google Sheets OFF. Configure `.streamlit/secrets.toml` e compartilhe a planilha com o Service Account.")
    st.stop()

st.success("Google Sheets conectado ✅")
st.caption(f"Planilha: {st.secrets['gsheets']['spreadsheet_url']}")

abas = list_tabs()
if not abas:
    st.warning("Não encontrei abas na planilha.")
    st.stop()

col_a, col_b = st.columns([2,1])
with col_a:
    tab = st.selectbox("Escolha a aba para visualizar/editar:", abas)
with col_b:
    st.write("")  # espaçamento
    if st.button("↻ Recarregar aba selecionada"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.rerun()

st.divider()

# Carrega a aba escolhida
df_tab = read_tab_as_df(tab)
st.caption(f"Linhas: {len(df_tab)} • Colunas: {list(df_tab.columns)}")

# Editor interativo
edited = st.data_editor(
    df_tab,
    num_rows="dynamic",
    use_container_width=True,
    key=f"editor_{tab}"
)

st.info("💡 Você pode alterar valores e adicionar/remover linhas no editor acima. Em seguida, escolha uma das opções de salvamento abaixo.")

# Ações de salvamento
st.subheader("💾 Salvar alterações na aba")
col1, col2 = st.columns(2)

with col1:
    if st.button("🧹 Sobrescrever a aba inteira (mantendo o cabeçalho do editor)"):
        try:
            overwrite_tab_from_df(tab, edited, keep_header=True)
            st.success("Aba sobrescrita com sucesso!")
        except Exception as e:
            st.error(f"Falha ao sobrescrever: {e}")

with col2:
    modo = st.selectbox(
        "Acrescentar apenas novas linhas (deduplicação por chave)",
        [
            "(Escolha a chave)",
            "Vistorias: OBJETO+OM+DATA",
            "OMs: Nome",
            "Personalizada…"
        ],
        index=0
    )

    if modo == "Vistorias: OBJETO+OM+DATA":
        key_cols = ["OBJETO DE VISTORIA", "OM APOIADA", "DATA DA SOLICITACAO"]
    elif modo == "OMs: Nome":
        key_cols = ["Nome"]
    elif modo == "Personalizada…":
        # permite escolher as colunas existentes da aba
        key_cols = st.multiselect("Selecione as colunas que formam a chave única:", list(edited.columns))
    else:
        key_cols = None

    if key_cols:
        if st.button("➕ Acrescentar somente novas linhas (com base na chave)"):
            try:
                added = append_new_rows_dedup(tab, edited, key_cols)
                if added == 0:
                    st.warning("Nenhuma linha nova para acrescentar (todas já existiam pela chave escolhida).")
                else:
                    st.success(f"{added} novas linha(s) acrescentadas!")
            except Exception as e:
                st.error(f"Falha ao acrescentar: {e}")
