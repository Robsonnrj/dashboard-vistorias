# app_dashboard_vistorias_v2.py
# -*- coding: utf-8 -*-

# ╭──────────────────────────────────────────────────────────────────────────╮
# │  CRO1 — Sistema de Vistorias (Streamlit + Google Sheets)                 │
# │  - GSheets via st.secrets (sem experimental_rerun)                       │
# │  - Importar Excel -> Sheets (sem duplicar)                               │
# │  - Fallback local (CSV)                                                  │
# │  - Dashboards e filtros                                                  │
# ╰──────────────────────────────────────────────────────────────────────────╯

# ── silenciar avisos chatos do openpyxl ──────────────────────────────────────
import warnings
warnings.filterwarnings(
    "ignore",
    message=".*outside the limits for dates.*",
    category=UserWarning,
    module="openpyxl",
)
warnings.filterwarnings(
    "ignore",
    message=".*Data Validation extension is not supported and will be removed.*",
    category=UserWarning,
    module="openpyxl",
)

# ── imports principais ──────────────────────────────────────────────────────
import streamlit as st
from streamlit_option_menu import option_menu
import pandas as pd
import plotly.express as px
import unicodedata
from datetime import datetime
from pathlib import Path

# ── streamlit config ────────────────────────────────────────────────────────
st.set_page_config(page_title="CRO1 - Sistema de Vistorias (GSheets)", layout="wide")

# ── constantes ──────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

OM_HEADERS = ["Nome", "Sigla", "Diretoria", "Criado em"]
VT_HEADERS = [
    "OBJETO DE VISTORIA",
    "OM APOIADA",
    "Diretoria Responsavel",
    "Classificacao da Urgencia",
    "Situacao",
    "DATA DA SOLICITACAO",
]

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║               BLOCO ÚNICO: GOOGLE SHEETS HELPERS                        ║
# ╚══════════════════════════════════════════════════════════════════════════╝
def has_gsheets() -> bool:
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and "spreadsheet_url" in st.secrets["gsheets"]
    )

@st.cache_resource(show_spinner=False)
def _gs_client():
    # importar aqui para permitir rodar sem libs quando secrets não existem
    import gspread
    from google.oauth2.service_account import Credentials

    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def _gs_book():
    return _gs_client().open_by_url(st.secrets["gsheets"]["spreadsheet_url"])

def _get_ws(title: str, header: list[str]):
    """Garante worksheet com cabeçalho correto."""
    import gspread

    sh = _gs_book()
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=2000, cols=max(10, len(header)))
        ws.append_row(header)
        return ws
    # corrige cabeçalho se necessário
    if ws.row_values(1) != header:
        ws.update("1:1", [header])
    return ws

@st.cache_data(ttl=60, show_spinner=False)
def gs_read_df(title: str, header: list[str]) -> pd.DataFrame:
    """Leitura em cache (60s) de uma aba do Sheets."""
    ws = _get_ws(title, header)
    data = ws.get_all_records()
    df = pd.DataFrame(data)
    # normaliza datas
    for c in df.columns:
        if "DATA" in c.upper():
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df

def gs_append_row(title: str, header: list[str], row_dict: dict):
    """Acrescenta uma linha e limpa o cache de leitura daquela aba."""
    ws = _get_ws(title, header)
    row = [row_dict.get(h, "") for h in header]
    ws.append_row(row, value_input_option="USER_ENTERED")
    gs_read_df.clear()

def gs_upsert_om(nome: str, sigla: str, diretoria: str):
    """Evita duplicar OM pelo nome (case-insensitive)."""
    df = gs_read_df("OMs", OM_HEADERS)
    if not df.empty and "Nome" in df.columns:
        existe = df["Nome"].astype(str).str.strip().str.lower().eq(nome.strip().lower()).any()
        if existe:
            return
    gs_append_row("OMs", OM_HEADERS, {
        "Nome": nome.strip(),
        "Sigla": sigla.strip(),
        "Diretoria": diretoria.strip(),
        "Criado em": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                     HELPERS DE TEXTO / COLUNAS                           ║
# ╚══════════════════════════════════════════════════════════════════════════╝
def _norm(s: str) -> str:
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower().strip()

def achar(col_alvo, candidatos):
    alvo_n = _norm(col_alvo)
    for c in candidatos:
        if _norm(c) == alvo_n:
            return c
    for c in candidatos:
        if alvo_n in _norm(c):
            return c
    return None

def achar_multi(alvos, candidatos):
    for a in alvos:
        x = achar(a, candidatos)
        if x:
            return x
    return None

def col_or_none(df: pd.DataFrame, opts: list[str]) -> str | None:
    cols = list(df.columns)
    # exata
    for o in opts:
        for c in cols:
            if _norm(c) == _norm(o):
                return c
    # contém
    for o in opts:
        alvo = _norm(o)
        for c in cols:
            if alvo in _norm(c):
                return c
    return None

def card_title(txt: str):
    st.markdown(
        "<div style='padding:8px 12px;border-radius:10px;background:#f6f6f9;"
        "border:1px solid #e5e7eb;font-weight:700;font-size:20px;'>"
        f"{txt}</div>",
        unsafe_allow_html=True
    )

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                  PERSISTÊNCIA LOCAL (FALLBACK CSV)                       ║
# ╚══════════════════════════════════════════════════════════════════════════╝
def init_om_store():
    if "om_cad" not in st.session_state:
        if has_gsheets():
            try:
                st.session_state.om_cad = gs_read_df("OMs", OM_HEADERS)
                return
            except Exception:
                pass
        # fallback
        try:
            st.session_state.om_cad = pd.read_csv("om_cadastro.csv")
        except Exception:
            st.session_state.om_cad = pd.DataFrame(columns=OM_HEADERS)

def init_vist_store():
    if "vist_cad" not in st.session_state:
        if has_gsheets():
            try:
                st.session_state.vist_cad = gs_read_df("Vistorias", VT_HEADERS)
                return
            except Exception:
                pass
        # fallback
        try:
            st.session_state.vist_cad = pd.read_csv(
                "vistorias_local.csv", parse_dates=["DATA DA SOLICITACAO"]
            )
        except Exception:
            st.session_state.vist_cad = pd.DataFrame(columns=VT_HEADERS)

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                      LEITURA DO EXCEL (UPLOAD)                           ║
# ╚══════════════════════════════════════════════════════════════════════════╝
arquivo = st.sidebar.file_uploader(
    "Envie o Excel (.xlsx) ou deixe em branco para usar o arquivo da pasta",
    type=["xlsx"],
    help="Se vazio, o app tenta abrir 'Acomp. de Vistorias CRO1 - 2025.xlsx' na raiz do projeto.",
)

@st.cache_data
def carregar_excel(file_like):
    if file_like is None:
        try:
            xls = pd.ExcelFile("Acomp. de Vistorias CRO1 - 2025.xlsx", engine="openpyxl")
        except Exception:
            return None, None, []
    else:
        xls = pd.ExcelFile(file_like, engine="openpyxl")

    preferidas = ["ACOMPANHAMENTO VISTORIAS", "Acompanhamento Vistorias"]
    alvo = next((n for n in xls.sheet_names if n in preferidas), xls.sheet_names[0])
    df = pd.read_excel(xls, sheet_name=alvo)
    return df, alvo, xls.sheet_names

df_raw, aba, abas = carregar_excel(arquivo)

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                 IMPORTAR EXCEL -> GOOGLE SHEETS                          ║
# ╚══════════════════════════════════════════════════════════════════════════╝
def _ensure_ws_with_header(sheet, title: str, header: list[str]):
    import gspread
    try:
        ws = sheet.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sheet.add_worksheet(title=title, rows=2000, cols=max(10, len(header)))
        ws.update("1:1", [header])
        return ws
    if ws.row_values(1) != header:
        ws.update("1:1", [header])
    return ws

def _append_batch(ws, rows: list[list], chunk=200):
    if not rows:
        return
    for i in range(0, len(rows), chunk):
        ws.append_rows(rows[i:i+chunk], value_input_option="USER_ENTERED")

def importar_excel_para_sheets(df: pd.DataFrame) -> dict:
    """Envia vistorias e OMs do DataFrame para o Sheets sem duplicar."""
    import gspread

    resumo = {"vistorias_importadas": 0, "vistorias_puladas_dup": 0,
              "oms_importadas": 0, "oms_puladas_existentes": 0}

    sh = _gs_book()

    # mapear colunas
    c_obj = col_or_none(df, ["OBJETO DE VISTORIA", "OBJETO"])
    c_om  = col_or_none(df, ["OM APOIADA", "OM APOIADORA", "OM"])
    c_dir = col_or_none(df, ["Diretoria Responsavel", "Diretoria Responsável", "Diretoria"])
    c_urg = col_or_none(df, ["Classificacao da Urgencia","Classificação da Urgência","Urgencia"])
    c_sit = col_or_none(df, ["Situacao", "Situação"])
    c_data_solic = col_or_none(df, ["DATA DA SOLICITACAO", "DATA DA SOLICITAÇÃO"])

    # ---------- Vistorias ----------
    if all([c_obj, c_om, c_data_solic]):
        df[c_data_solic] = pd.to_datetime(df[c_data_solic], errors="coerce")
        header_v = VT_HEADERS
        ws_v = _ensure_ws_with_header(sh, "Vistorias", header_v)

        exist = ws_v.get_all_records()
        if exist:
            de = pd.DataFrame(exist)
            if "DATA DA SOLICITACAO" in de.columns:
                de["DATA DA SOLICITACAO"] = pd.to_datetime(de["DATA DA SOLICITACAO"], errors="coerce")
            keys = set(
                (str(de.at[i, "OBJETO DE VISTORIA"]).strip(),
                 str(de.at[i, "OM APOIADA"]).strip(),
                 str(de.at[i, "DATA DA SOLICITACAO"].date()) if pd.notna(de.at[i, "DATA DA SOLICITACAO"]) else "")
                for i in de.index
            )
        else:
            keys = set()

        rows = []
        for _, r in df.iterrows():
            obj = str(r.get(c_obj, "")).strip()
            om  = str(r.get(c_om, "")).strip()
            d   = r.get(c_data_solic, pd.NaT)
            d_s = str(pd.to_datetime(d).date()) if pd.notna(d) else ""
            if not obj or not om:
                continue
            key = (obj, om, d_s)
            if key in keys:
                resumo["vistorias_puladas_dup"] += 1
                continue
            diretoria = str(r.get(c_dir, "")).strip() if c_dir else ""
            urg      = str(r.get(c_urg, "")).strip() if c_urg else ""
            sit      = str(r.get(c_sit, "")).strip() if c_sit else ""
            rows.append([obj, om, diretoria, urg, sit, d_s])

        if rows:
            _append_batch(ws_v, rows, chunk=200)
            resumo["vistorias_importadas"] = len(rows)

    # ---------- OMs ----------
    if c_om:
        pares_dir = {}
        if c_dir:
            tmp = df[[c_om, c_dir]].dropna().drop_duplicates()
            for _, r2 in tmp.iterrows():
                pares_dir[str(r2[c_om]).strip()] = str(r2[c_dir]).strip()
        om_unicas = sorted([str(x).strip() for x in df[c_om].dropna().unique()])

        ws_o = _ensure_ws_with_header(sh, "OMs", OM_HEADERS)
        exist_o = ws_o.get_all_records()
        de_o = pd.DataFrame(exist_o) if exist_o else pd.DataFrame(columns=OM_HEADERS)
        nomes_exist = set(de_o["Nome"].astype(str).str.strip()) if "Nome" in de_o.columns else set()

        rows_om = []
        for nome in om_unicas:
            if not nome or nome in nomes_exist:
                resumo["oms_puladas_existentes"] += 1
                continue
            rows_om.append([
                nome, "", pares_dir.get(nome, ""), datetime.now().strftime("%Y-%m-%d %H:%M")
            ])

        if rows_om:
            _append_batch(ws_o, rows_om, chunk=200)
            resumo["oms_importadas"] = len(rows_om)

    # limpa cache
    gs_read_df.clear()
    return resumo

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                           NAVEGAÇÃO / MENU                               ║
# ╚══════════════════════════════════════════════════════════════════════════╝
PAGES = ["🏠 Início", "📝 Nova Vistoria", "🏢 Nova OM", "📑 Novo Relatório", "🔍 Consulta", "📊 Resumos (Dashboards)"]
if "page" not in st.session_state:
    st.session_state.page = PAGES[0]

def goto(page):
    st.session_state.page = page
    st.session_state.menu_index = PAGES.index(page) if page in PAGES else 0
    st.session_state["_nav_ts"] = datetime.now().isoformat()  # evita cache agressivo

with st.sidebar:
    st.write("🔌 Google Sheets:", "ON ✅" if has_gsheets() else "OFF ❌")
    if not has_gsheets():
        st.error("Secrets não detectado. Verifique .streamlit/secrets.toml e a seção [gsheets].")

    # ações utilitárias
    colA, colB = st.columns(2)
    with colA:
        if st.button("🔄 Recarregar do Sheets", disabled=not has_gsheets()):
            gs_read_df.clear()
            for k in ("om_cad","vist_cad"):
                if k in st.session_state: del st.session_state[k]
            init_om_store(); init_vist_store()
            st.success("Recarregado do Google Sheets.")

    with colB:
        if st.button("🧹 Limpar caches"):
            gs_read_df.clear()
            st.cache_data.clear()
            st.cache_resource.clear()
            st.success("Caches limpos. Atualize a página (Ctrl+R).")

    # Importar Excel -> Sheets
    if df_raw is not None:
        if has_gsheets():
            st.markdown("### ⚙️ Importar Excel para o Sheets")
            st.caption("Vistorias → **Vistorias** • OMs → **OMs** (sem duplicar).")
            if st.button("📤 Enviar Excel ao Sheets"):
                try:
                    resumo = importar_excel_para_sheets(df_raw)
                    st.success(
                        f"✅ Concluído!\n\n"
                        f"- Vistorias importadas: **{resumo['vistorias_importadas']}** "
                        f"(puladas: {resumo['vistorias_puladas_dup']})\n"
                        f"- OMs importadas: **{resumo['oms_importadas']}** "
                        f"(puladas: {resumo['oms_puladas_existentes']})"
                    )
                except Exception as e:
                    st.error(f"Falha ao importar: {e}")
        else:
            st.info("Para enviar o Excel ao Google Sheets, ative os *secrets* primeiro.")

    # menu
    idx = st.session_state.get("menu_index", 0)
    escolha = option_menu(
        "Seção de Vistorias — CRO1",
        PAGES,
        icons=["house", "clipboard-check", "building", "file-earmark-text", "search", "bar-chart"],
        menu_icon="list",
        default_index=idx
    )
    if escolha != st.session_state.page:
        goto(escolha)

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                                 PÁGINAS                                   ║
# ╚══════════════════════════════════════════════════════════════════════════╝

# Início
if st.session_state.page == "🏠 Início":
    st.title("Seção de Vistorias — CRO1")
    st.write("Use os cards abaixo ou o menu à esquerda para navegar.")
    c1,c2,c3 = st.columns(3)
    with c1:
        if st.button("📝 Nova Vistoria", use_container_width=True):
            goto("📝 Nova Vistoria")
    with c2:
        if st.button("🏢 Nova OM", use_container_width=True):
            goto("🏢 Nova OM")
    with c3:
        if st.button("📊 Resumos (Dashboards)", use_container_width=True):
            goto("📊 Resumos (Dashboards)")
    st.info("Dica: envie a planilha na sidebar, se quiser sobrepor a que está na pasta.")

# Nova Vistoria
elif st.session_state.page == "📝 Nova Vistoria":
    st.title("📝 Nova Vistoria")
    init_om_store(); init_vist_store()

    # OMs do cadastro (Sheets/CSV) + Excel
    oms = []
    om_to_dir = {}

    if not st.session_state.om_cad.empty:
        oms = (st.session_state.om_cad["Nome"].dropna().astype(str).str.strip().unique().tolist())
        for _, r in st.session_state.om_cad.iterrows():
            nome = str(r.get("Nome", "")).strip()
            diretoria = str(r.get("Diretoria", "")).strip()
            if nome:
                om_to_dir[nome] = diretoria

    if df_raw is not None and not df_raw.empty:
        cols = list(df_raw.columns)
        c_om_xls  = achar_multi(["OM APOIADA","OM APOIADORA","OM"], cols)
        c_dir_xls = achar_multi(["Diretoria Responsavel","Diretoria Responsável","Diretoria"], cols)
        if c_om_xls:
            for nome in df_raw[c_om_xls].dropna().astype(str).str.strip().unique().tolist():
                if nome not in oms:
                    oms.append(nome)
        if c_om_xls and c_dir_xls:
            pares = df_raw[[c_om_xls, c_dir_xls]].dropna().drop_duplicates()
            for _, r in pares.iterrows():
                om_to_dir[str(r[c_om_xls]).strip()] = str(r[c_dir_xls]).strip()

    oms = sorted(set(oms))

    with st.form("form_vistoria", clear_on_submit=True):
        objeto = st.text_input("Descrição / Objeto de Vistoria")
        om_escolhida = st.selectbox("Organização Militar Apoiada", ["(digitar)"] + oms, index=0)
        if om_escolhida == "(digitar)":
            om_texto = st.text_input("Informe a OM", placeholder="Ex.: 1º DSUP")
            diretoria_default = ""
        else:
            om_texto = om_escolhida
            diretoria_default = om_to_dir.get(om_escolhida, "")

        diretoria = st.text_input("Diretoria Responsável", value=diretoria_default, placeholder="Ex.: DECEx / COTER / DGP")
        urg = st.selectbox("Classificação de Urgência", ["Não Prioridade", "Prioridade", "Urgente"])
        situacao = st.selectbox("Situação atual", ["Não Atendida", "Em andamento", "Finalizada"])
        data_solic = st.date_input("Data da solicitação")

        salvar_csv = st.checkbox("Salvar também em CSV local (vistorias_local.csv)", value=not has_gsheets())
        enviar = st.form_submit_button("Salvar Vistoria")

        if enviar:
            if not objeto or not om_texto:
                st.error("Preencha **Objeto** e **OM**.")
            else:
                nova = {
                    "OBJETO DE VISTORIA": objeto.strip(),
                    "OM APOIADA": om_texto.strip(),
                    "Diretoria Responsavel": diretoria.strip(),
                    "Classificacao da Urgencia": urg,
                    "Situacao": situacao,
                    "DATA DA SOLICITACAO": pd.to_datetime(data_solic)
                }
                if has_gsheets():
                    try:
                        payload = nova.copy()
                        payload["DATA DA SOLICITACAO"] = (
                            str(nova["DATA DA SOLICITACAO"].date()) if pd.notna(nova["DATA DA SOLICITACAO"]) else ""
                        )
                        gs_append_row("Vistorias", VT_HEADERS, payload)
                        st.success("✅ Vistoria registrada no Google Sheets!")
                        # recarregar
                        gs_read_df.clear()
                        if "vist_cad" in st.session_state: del st.session_state["vist_cad"]
                        init_vist_store()
                    except Exception as e:
                        st.warning(f"Falha ao salvar no Google Sheets. Farei fallback local. Detalhes: {e}")

                # sessão + CSV (fallback/espelho)
                st.session_state.vist_cad = pd.concat(
                    [st.session_state.vist_cad, pd.DataFrame([nova])],
                    ignore_index=True
                )
                if salvar_csv:
                    try:
                        st.session_state.vist_cad.to_csv("vistorias_local.csv", index=False, encoding="utf-8-sig")
                    except Exception as e:
                        st.warning(f"Falha ao salvar CSV local. Detalhes: {e}")

                st.success("✅ Vistoria registrada!")

    st.subheader("Vistorias criadas (sessão/CSV/Sheets)")
    if st.session_state.vist_cad.empty:
        st.info("Ainda não há vistorias criadas pelo app.")
    else:
        st.dataframe(
            st.session_state.vist_cad.sort_values("DATA DA SOLICITACAO", ascending=False),
            use_container_width=True
        )
        st.download_button(
            "↓ Baixar CSV dessas vistorias",
            st.session_state.vist_cad.to_csv(index=False).encode("utf-8-sig"),
            "vistorias_local.csv",
            "text/csv"
        )

# Nova OM
elif st.session_state.page == "🏢 Nova OM":
    st.title("🏢 Cadastro de Organização Militar (OM)")
    init_om_store()

    with st.form("form_om", clear_on_submit=True):
        colA,colB,colC = st.columns([2,1,2])
        with colA:
            nome = st.text_input("Nome da OM", placeholder="Ex.: 1º DSUP")
        with colB:
            sigla = st.text_input("Sigla", placeholder="Ex.: DSUP")
        with colC:
            diretoria = st.text_input("Diretoria de Subordinação", placeholder="Ex.: DECEx / COTER / DGP")

        salvar_csv = st.checkbox("Salvar em arquivo local (om_cadastro.csv)", value=not has_gsheets())
        submit = st.form_submit_button("Salvar OM")

        if submit:
            if not nome or not sigla:
                st.error("Preencha pelo menos **Nome** e **Sigla**.")
            else:
                nova = {
                    "Nome": nome.strip(),
                    "Sigla": sigla.strip(),
                    "Diretoria": diretoria.strip(),
                    "Criado em": datetime.now().strftime("%Y-%m-%d %H:%M")
                }
                if has_gsheets():
                    try:
                        gs_upsert_om(nova["Nome"], nova["Sigla"], nova["Diretoria"])
                        st.success("✅ OM registrada no Google Sheets!")
                        gs_read_df.clear()
                        if "om_cad" in st.session_state: del st.session_state["om_cad"]
                        init_om_store()
                    except Exception as e:
                        st.warning(f"Falha ao salvar no Google Sheets. Farei fallback local. Detalhes: {e}")

                st.session_state.om_cad = pd.concat(
                    [st.session_state.om_cad, pd.DataFrame([nova])],
                    ignore_index=True
                )
                if salvar_csv:
                    try:
                        st.session_state.om_cad.to_csv("om_cadastro.csv", index=False, encoding="utf-8-sig")
                    except Exception as e:
                        st.warning(f"Falha ao salvar CSV local. Detalhes: {e}")
                st.success("✅ OM registrada!")

    st.subheader("OMs cadastradas")
    st.dataframe(st.session_state.om_cad.sort_values("Criado em", ascending=False), use_container_width=True)

# Novo Relatório (placeholder)
elif st.session_state.page == "📑 Novo Relatório":
    st.title("📑 Novo Relatório")
    st.write("Gerador de relatório pode ser implementado aqui.")

# Consulta
elif st.session_state.page == "🔍 Consulta":
    st.title("🔍 Consulta")
    if df_raw is None:
        st.warning("Envie um Excel na sidebar ou deixe o arquivo 'Acomp. de Vistorias CRO1 - 2025.xlsx' na pasta.")
    else:
        st.caption(f"Aba carregada: **{aba}**  •  Abas no arquivo: {abas}")
        st.dataframe(df_raw, use_container_width=True)

    st.markdown("### OMs cadastradas (sessão/CSV/Sheets)")
    init_om_store(); st.dataframe(st.session_state.om_cad, use_container_width=True)

    st.markdown("### Vistorias criadas (sessão/CSV/Sheets)")
    init_vist_store(); st.dataframe(st.session_state.vist_cad, use_container_width=True)

# Dashboards
elif st.session_state.page == "📊 Resumos (Dashboards)":
    st.title("📊 Resumos (Dashboards)")
    init_vist_store()

    # usa Excel + vistorias do app; se não houver Excel, usa só as do app
    if df_raw is None or df_raw.empty:
        df = st.session_state.vist_cad.copy()
        if df.empty:
            st.warning("Envie um Excel na sidebar ou cadastre vistorias na aba 'Nova Vistoria'.")
            st.stop()
        aba = "(somente vistorias do app)"
        abas = []
    else:
        df = df_raw.copy()
        if not st.session_state.vist_cad.empty:
            # adiciona colunas que faltem
            for c in st.session_state.vist_cad.columns:
                if c not in df.columns:
                    df[c] = pd.NA
            comuns = [c for c in df.columns if c in st.session_state.vist_cad.columns]
            df = pd.concat([df, st.session_state.vist_cad[comuns]], ignore_index=True)

    st.caption(f"Aba carregada: **{aba}**  •  Abas no arquivo: {abas}")

    # mapeamento resiliente
    c_obj = col_or_none(df, ["OBJETO DE VISTORIA", "OBJETO"])
    c_om  = col_or_none(df, ["OM APOIADA", "OM APOIADORA", "OM"])
    c_dir = col_or_none(df, ["Diretoria Responsavel", "Diretoria Responsável", "Diretoria"])
    c_urg = col_or_none(df, ["Classificacao da Urgencia","Classificação da Urgência","Urgencia"])
    c_sit = col_or_none(df, ["Situacao", "Situação"])
    c_data_solic = col_or_none(df, ["DATA DA SOLICITACAO", "DATA DA SOLICITAÇÃO"])
    c_data_vist  = col_or_none(df, ["DATA DA VISTORIA"])
    c_dias_total = col_or_none(df, ["QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO"])
    c_dias_exec  = col_or_none(df, ["QUANTIDADE DE DIAS PARA EXECUCAO", "QUANTIDADE DE DIAS PARA EXECUÇÃO"])
    c_status     = col_or_none(df, ["STATUS - ATUALIZACAO SEMANAL", "STATUS - ATUALIZAÇÃO SEMANAL", "Status"])

    # tipos
    for c in [c_data_solic, c_data_vist]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in [c_dias_total, c_dias_exec]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # filtros
    st.sidebar.subheader("Filtros")
    col_data_base = c_data_solic if c_data_solic in df.columns else c_data_vist
    if col_data_base and df[col_data_base].notna().any():
        min_dt = pd.to_datetime(df[col_data_base].min()).date()
        max_dt = pd.to_datetime(df[col_data_base].max()).date()
        periodo = st.sidebar.date_input("Período (pela data da solicitação)", value=(min_dt, max_dt))
    else:
        periodo = None

    def opts(series):
        try:
            return sorted(series.dropna().astype(str).unique().tolist())
        except Exception:
            return sorted(list({str(x) for x in series.dropna().tolist()}))

    dir_sel = st.sidebar.multiselect("Diretoria Responsável", opts(df[c_dir]) if c_dir in df.columns else [])
    sit_sel = st.sidebar.multiselect("Situação", opts(df[c_sit]) if c_sit in df.columns else [])
    urg_sel = st.sidebar.multiselect("Classificação de Urgência", opts(df[c_urg]) if c_urg in df.columns else [])
    om_sel  = st.sidebar.multiselect("OM Apoiadora", opts(df[c_om]) if c_om in df.columns else [])
    sla_dias = st.sidebar.number_input("SLA (dias para 'dentro do prazo')", 1, 365, value=30)

    df_f = df.copy()
    if periodo and col_data_base:
        ini, fim = periodo
        df_f = df_f[(df_f[col_data_base] >= pd.to_datetime(ini)) & (df_f[col_data_base] <= pd.to_datetime(fim))]
    if dir_sel and c_dir in df.columns:
        df_f = df_f[df_f[c_dir].astype(str).isin(dir_sel)]
    if sit_sel and c_sit in df.columns:
        df_f = df_f[df_f[c_sit].astype(str).isin(sit_sel)]
    if urg_sel and c_urg in df.columns:
        df_f = df_f[df_f[c_urg].astype(str).isin(urg_sel)]
    if om_sel and c_om in df.columns:
        df_f = df_f[df_f[c_om].astype(str).isin(om_sel)]

    # KPIs
    colk1, colk2, colk3, colk4, colk5 = st.columns(5)
    total_vist = len(df_f)
    finalizadas = df_f[c_sit].astype(str).str.upper().str.contains("FINALIZADA").sum() if c_sit in df_f.columns else None
    pct_final = (finalizadas / total_vist * 100) if (finalizadas is not None and total_vist > 0) else 0
    prazo_medio_total = df_f[c_dias_total].mean() if c_dias_total in df_f.columns else None
    prazo_medio_exec  = df_f[c_dias_exec].mean() if c_dias_exec   in df_f.columns else None
    pct_sla = None
    if c_dias_total in df_f.columns and total_vist > 0:
        dentro_sla = (df_f[c_dias_total] <= sla_dias).sum()
        pct_sla = dentro_sla / total_vist * 100

    with colk1: st.metric("Total de Vistorias", f"{total_vist:,}".replace(",", "."))
    with colk2: st.metric("Finalizadas (%)", f"{pct_final:,.1f}%")
    with colk3: st.metric("Prazo médio total (dias)", f"{prazo_medio_total:,.1f}" if prazo_medio_total is not None else "—")
    with colk4: st.metric("Prazo médio execução (dias)", f"{prazo_medio_exec:,.1f}" if prazo_medio_exec is not None else "—")
    with colk5: st.metric(f"% dentro do SLA (≤{sla_dias}d)", f"{pct_sla:,.1f}%" if pct_sla is not None else "—")

    st.divider()

    # gráficos
    if col_data_base and df_f[col_data_base].notna().any():
        tmp = (df_f.groupby(pd.Grouper(key=col_data_base, freq="MS")).size().reset_index(name="Vistorias"))
        st.plotly_chart(px.line(tmp, x=col_data_base, y="Vistorias", markers=True,
                                title="Evolução Mensal de Vistorias"), use_container_width=True)

    if c_dir in df_f.columns:
        tmp2 = df_f.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False)
        st.plotly_chart(px.bar(tmp2, x=c_dir, y="size", title="Vistorias por Diretoria Responsável"),
                        use_container_width=True)

    if c_sit in df_f.columns:
        tmp3 = df_f.groupby(c_sit, as_index=False).size()
        st.plotly_chart(px.pie(tmp3, names=c_sit, values="size", hole=0.4,
                               title="Distribuição por Situação"), use_container_width=True)

    if c_urg in df_f.columns:
        tmp4 = df_f.groupby(c_urg, as_index=False).size().sort_values("size", ascending=False)
        st.plotly_chart(px.bar(tmp4, x=c_urg, y="size", title="Vistorias por Classificação de Urgência"),
                        use_container_width=True)

    if c_dir in df_f.columns and c_dias_total in df_f.columns:
        base = df_f.dropna(subset=[c_dir, c_dias_total]).copy()
        base["Dentro SLA"] = base[c_dias_total] <= sla_dias
        tmp_sla = (base.groupby(c_dir)["Dentro SLA"].mean()*100).reset_index(name="pct_sla").sort_values("pct_sla")
        st.plotly_chart(px.bar(tmp_sla, x="pct_sla", y=c_dir, orientation="h",
                               title=f"% Dentro do SLA (≤{sla_dias}d) por Diretoria",
                               labels={"pct_sla": "% dentro do SLA"}), use_container_width=True)

    if col_data_base and c_sit in df_f.columns and df_f[col_data_base].notna().any():
        aux = df_f.copy()
        aux["Mes"] = aux[col_data_base].dt.to_period("M").dt.to_timestamp()
        piv = (aux.groupby(["Mes", c_sit]).size().reset_index(name="Qtd")
               .pivot(index="Mes", columns=c_sit, values="Qtd").fillna(0))
        st.plotly_chart(px.imshow(piv.T, aspect="auto", labels=dict(x="Mês", y="Situação", color="Qtd"),
                                  title="Heatmap — Mês x Situação"), use_container_width=True)

    card_title("Detalhamento (mais recentes)")
    ord_col = col_data_base if col_data_base else (c_data_vist if c_data_vist in df_f.columns else None)
    df_show = df_f.sort_values(ord_col, ascending=False).head(50) if ord_col else df_f.head(50)
    st.dataframe(df_show, use_container_width=True)
