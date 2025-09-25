# -*- coding: utf-8 -*-
# CRO1 — Dashboard & Editor (Google Sheets) | Tema claro

import warnings
warnings.filterwarnings("ignore")

from datetime import datetime
import unicodedata

import pandas as pd
import plotly.express as px
import streamlit as st

import gspread
from google.oauth2.service_account import Credentials
from streamlit_option_menu import option_menu

# =============== CONFIG GERAL ===============
st.set_page_config(page_title="CRO1 — Gestão de Vistorias", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# =============== ESTILO (CSS LIGHT) ===============
LIGHT_CSS = """
<style>
.block-container { padding-top: 1.1rem; max-width: 1400px; }
[data-testid="stSidebar"] { background: #ffffff; border-right: 1px solid #e5e7eb; }
section.main { background: #ffffff; }

.card {
  background: #ffffff; border: 1px solid #e5e7eb; border-radius: 14px; padding: 16px;
  box-shadow: 0 6px 16px rgba(0,0,0,.05);
}
.card-title { font-weight: 700; font-size: .95rem; color: #4b5563; letter-spacing: .2px; }
.kpi { font-size: 2rem; font-weight: 800; color: #111827; }
.kpi-sub { color: #6b7280; font-size: .85rem; margin-top: .25rem; }

.stButton>button { border-radius: 10px !important; padding: 9px 12px; }
.badge {
  display:inline-block; padding: 2px 10px; border-radius: 999px; font-size: .75rem;
  border:1px solid #e5e7eb; background:#f8fafc; color:#111827; margin-right:6px;
}
tbody tr:hover { background: #f9fafb !important; }
</style>
"""
st.markdown(LIGHT_CSS, unsafe_allow_html=True)

# =============== FUNÇÕES BASE (Sheets/cache) ===============
def has_gsheets() -> bool:
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and "spreadsheet_url" in st.secrets["gsheets"]
        and bool(st.secrets["gsheets"]["spreadsheet_url"])
    )

@st.cache_resource(show_spinner=False)
def _gs_client():
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

def _make_unique_headers(row):
    out, seen = [], {}
    for j, h in enumerate(row, start=1):
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
    hdr_idx = next((i for i,row in enumerate(values) if any(str(c).strip() for c in row)), 0)
    hdr = _make_unique_headers(values[hdr_idx])
    body = values[hdr_idx+1:]
    while body and not any(str(c).strip() for c in body[-1]):
        body.pop()
    df = pd.DataFrame(body, columns=hdr).replace("", pd.NA)
    # normalizações úteis
    for c in df.columns:
        up = c.upper()
        if "DATA" in up:
            df[c] = pd.to_datetime(df[c], errors="coerce")
        if "DIAS" in up or "QUANTIDADE" in up:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

@st.cache_data(ttl=120, show_spinner=False)
def read_tab_df(tab_name: str) -> pd.DataFrame:
    ws = _book().worksheet(tab_name)
    return _read_ws_loose(ws)

# --------- OMs (Validacao_de_Dados) com cache de 5 min ---------
@st.cache_data(ttl=300, show_spinner=False)
def load_oms_from_validation() -> pd.DataFrame:
    """
    Retorna DF com ['sigla','nome','diretoria'] a partir da aba 'Validacao_de_Dados'.
    """
    try:
        df = read_tab_df("Validacao_de_Dados")
    except Exception:
        return pd.DataFrame(columns=["sigla","nome","diretoria"])

    col_sigla = next((c for c in df.columns if _norm(c) in ("om","sigla")), None)
    col_nome  = next((c for c in df.columns if "organizacao" in _norm(c) or "nome" in _norm(c)), None)
    col_dir   = next((c for c in df.columns if "diretoria" in _norm(c)), None)

    data = pd.DataFrame(columns=["sigla","nome","diretoria"])
    if col_sigla: data["sigla"] = df[col_sigla].astype(str).str.strip()
    if col_nome:  data["nome"]  = df[col_nome].astype(str).str.strip()
    if col_dir:   data["diretoria"] = df[col_dir].astype(str).str.strip()

    data = data.dropna(subset=["sigla"]).drop_duplicates(subset=["sigla"])
    for c in ["nome","diretoria"]:
        if c not in data.columns: data[c] = ""
    return data[["sigla","nome","diretoria"]]

# =============== SIDEBAR / MENU ===============
with st.sidebar:
    st.markdown("### Sistema CRO1 — Gestão de Vistorias")
    st.write("🔌 Google Sheets:", "ON ✅" if has_gsheets() else "OFF ❌")

    if has_gsheets() and st.button("🔄 Atualizar dados (limpar cache)"):
        read_tab_df.clear()
        load_oms_from_validation.clear()
        st.toast("Cache limpo. Recarregando…")
        st.rerun()  # <<==== CORRIGIDO (antes era experimental_rerun)

    MENU = option_menu(
        "",
        ["📊 Dashboard", "🗂 Editor (Sheets)"],
        icons=["bar-chart","table"],
        default_index=0,
        styles={
            "nav-link": {"font-size":"15px", "text-align":"left", "margin":"2px"},
            "nav-link-selected": {"background-color":"#f3f4f6"},
        }
    )

# =============== FILTROS HIERÁRQUICOS ===============
def render_filters(df_base: pd.DataFrame):
    # detectar colunas
    cols = list(df_base.columns)
    def pick(*cands):
        for x in cands:
            for c in cols:
                if _norm(c) == _norm(x): return c
        for x in cands:
            alvo = _norm(x)
            for c in cols:
                if alvo in _norm(c): return c
        return None

    c_dir = pick("Diretoria Responsável","Diretoria Responsavel","Diretoria")
    c_om  = pick("OM APOIADA","OM APOIADORA","OM")
    c_sit = pick("Situação","Situacao")
    c_urg = pick("Classificação da Urgência","Classificacao da Urgencia","Urgencia")
    c_obj = pick("OBJETO DE VISTORIA","OBJETO")
    c_dtS = pick("DATA DA SOLICITAÇÃO","DATA DA SOLICITACAO")
    c_dtV = pick("DATA DA VISTORIA")
    base_date_col = c_dtS or c_dtV

    # Período
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">Período</div>', unsafe_allow_html=True)
    periodo = None
    if base_date_col and df_base[base_date_col].notna().any():
        dmin = pd.to_datetime(df_base[base_date_col]).min().date()
        dmax = pd.to_datetime(df_base[base_date_col]).max().date()
        periodo = st.date_input("Intervalo", (dmin, dmax))
    st.markdown('</div>', unsafe_allow_html=True)

    # Diretoria -> OM (com validação)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">Filtros por Estrutura</div>', unsafe_allow_html=True)
    oms_val = load_oms_from_validation()
    if not oms_val.empty:
        dirs = sorted([d for d in oms_val["diretoria"].dropna().unique().tolist() if d])
        dir_sel = st.selectbox("🏢 Diretoria Responsável", ["(Todas)"] + dirs, index=0)
        pool = oms_val if dir_sel == "(Todas)" else oms_val[oms_val["diretoria"].astype(str)==dir_sel]
        om_opts = (pool
                   .assign(label=lambda d: d["sigla"].fillna("") + " — " + d["nome"].fillna(""))
                   .sort_values("sigla"))
        lookup = dict(zip(om_opts["label"], om_opts["sigla"]))
        om_labels = st.multiselect("🏛️ OM Apoiadora (digite para buscar)", om_opts["label"].tolist())
        om_sel = [lookup[x] for x in om_labels]
    else:
        st.info("Aba **Validacao_de_Dados** ausente/vazia. Usando OMs do dataset.")
        om_sel = sorted(df_base[c_om].dropna().astype(str).unique().tolist()) if c_om else []
        om_sel = st.multiselect("🏛️ OM Apoiadora", om_sel)

    sit_sel = st.multiselect("📋 Situação", sorted(df_base[c_sit].dropna().astype(str).unique().tolist()) if c_sit else [])
    urg_sel = st.multiselect("⚡ Urgência", sorted(df_base[c_urg].dropna().astype(str).unique().tolist()) if c_urg else [])
    q = st.text_input("🔎 Buscar no objeto de vistoria", placeholder="Palavra-chave…")
    st.markdown('</div>', unsafe_allow_html=True)

    return dict(
        periodo=periodo, base_date_col=base_date_col,
        diretoria=(dir_sel if 'dir_sel' in locals() else None),
        oms=om_sel, c_dir=c_dir, c_om=c_om, c_sit=c_sit, c_urg=c_urg, c_obj=c_obj,
        sit_sel=sit_sel, urg_sel=urg_sel, query=q
    )

def apply_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    res = df.copy()
    # período
    if f["periodo"] and f["base_date_col"] in res.columns:
        ini, fim = f["periodo"]
        res = res[(pd.to_datetime(res[f["base_date_col"]]) >= pd.to_datetime(ini)) &
                  (pd.to_datetime(res[f["base_date_col"]]) <= pd.to_datetime(fim))]
    # diretoria
    if f.get("diretoria") and f["diretoria"] != "(Todas)" and f["c_dir"] in res.columns:
        res = res[res[f["c_dir"]].astype(str) == f["diretoria"]]
    # OMs
    if f["oms"] and f["c_om"] in res.columns:
        res = res[res[f["c_om"]].astype(str).isin(f["oms"])]
    # Situação / Urgência
    if f["sit_sel"] and f["c_sit"] in res.columns:
        res = res[res[f["c_sit"]].astype(str).isin(f["sit_sel"])]
    if f["urg_sel"] and f["c_urg"] in res.columns:
        res = res[res[f["c_urg"]].astype(str).isin(f["urg_sel"])]
    # Busca no objeto
    if f["query"] and f["c_obj"] in res.columns:
        q = f["query"].strip().lower()
        res = res[res[f["c_obj"]].astype(str).str.lower().str.contains(q)]
    return res

# =============== DASHBOARD ===============
def render_kpis(df: pd.DataFrame, fcols: dict):
    c_sit = fcols.get("c_sit")
    c_dias_total = next((c for c in df.columns if "TOTAL" in c.upper() and "DIAS" in c.upper()), None)
    c_dias_exec  = next((c for c in df.columns if "EXEC" in c.upper() and "DIAS" in c.upper()), None)

    total = len(df)
    finalizadas = df[c_sit].astype(str).str.upper().str.contains("FINALIZADA").sum() if c_sit in df.columns else 0
    pct_final = (finalizadas/total*100) if total else 0
    prazo_total = df[c_dias_total].mean() if c_dias_total in df.columns else None
    prazo_exec  = df[c_dias_exec].mean()  if c_dias_exec  in df.columns else None

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown('<div class="card"><div class="card-title">Vistorias</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="kpi">{total:,}</div><div class="kpi-sub">Total filtrado</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="card"><div class="card-title">Finalizadas</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="kpi">{pct_final:,.1f}%</div><div class="kpi-sub">do total</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="card"><div class="card-title">Prazo médio (total)</div>', unsafe_allow_html=True)
        val = f"{prazo_total:,.1f} dias" if prazo_total is not None else "—"
        st.markdown(f'<div class="kpi">{val}</div><div class="kpi-sub">até atendimento</div></div>', unsafe_allow_html=True)
    with col4:
        st.markdown('<div class="card"><div class="card-title">Prazo médio (execução)</div>', unsafe_allow_html=True)
        val = f"{prazo_exec:,.1f} dias" if prazo_exec is not None else "—"
        st.markdown(f'<div class="kpi">{val}</div><div class="kpi-sub">de execução</div></div>', unsafe_allow_html=True)

def render_charts(df: pd.DataFrame, fcols: dict):
    c_dir  = fcols.get("c_dir")
    c_sit  = fcols.get("c_sit")
    c_urg  = fcols.get("c_urg")
    c_dtS  = fcols.get("base_date_col")

    st.markdown("#### Visualizações")

    # Evolução mensal
    if c_dtS and df[c_dtS].notna().any():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        tmp = (df.groupby(pd.Grouper(key=c_dtS, freq="MS")).size().reset_index(name="Vistorias"))
        fig1 = px.line(tmp, x=c_dtS, y="Vistorias", markers=True, title="Evolução Mensal")
        fig1.update_layout(margin=dict(l=10,r=10,b=10,t=40))
        st.plotly_chart(fig1, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    cols = st.columns(2)
    with cols[0]:
        if c_dir in df.columns:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            tmp2 = df.groupby(c_dir, as_index=False).size().sort_values("size", ascending=False).head(15)
            fig2 = px.bar(tmp2, x=c_dir, y="size", title="Por Diretoria")
            fig2.update_layout(margin=dict(l=10,r=10,b=10,t=40))
            st.plotly_chart(fig2, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)
    with cols[1]:
        if c_urg in df.columns:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            tmp4 = df.groupby(c_urg, as_index=False).size().sort_values("size", ascending=False)
            fig4 = px.bar(tmp4, x=c_urg, y="size", title="Por Classificação de Urgência")
            fig4.update_layout(margin=dict(l=10,r=10,b=10,t=40))
            st.plotly_chart(fig4, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

    if c_sit in df.columns:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        tmp3 = df.groupby(c_sit, as_index=False).size()
        fig3 = px.pie(tmp3, names=c_sit, values="size", hole=0.4, title="Distribuição por Situação")
        fig3.update_layout(margin=dict(l=10,r=10,b=10,t=40))
        st.plotly_chart(fig3, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

def render_table(df: pd.DataFrame, fcols: dict):
    st.markdown("#### Itens recentes")
    base_col = fcols.get("base_date_col")
    if base_col in df.columns:
        df = df.sort_values(base_col, ascending=False)
    st.dataframe(df.head(80), use_container_width=True, height=360)

# =============== EDITOR ===============
def render_editor():
    st.title("🗂 Editor (Google Sheets)")
    if not has_gsheets():
        st.stop()
    sh = _book()
    tabs = [ws.title for ws in sh.worksheets()]
    st.success("Conectado ao Google Sheets ✅")
    st.caption(f"Planilha: {st.secrets['gsheets']['spreadsheet_url']}")

    colA, colB = st.columns([2,1])
    with colA:
        tab_name = st.selectbox("Aba:", tabs, index=0)
    with colB:
        if st.button("↻ Recarregar aba"):
            read_tab_df.clear()
            st.rerun()

    try:
        df_tab = read_tab_df(tab_name)
    except Exception as e:
        st.error(f"Falha ao ler a aba **{tab_name}**: {e}")
        st.stop()

    st.caption(f"Linhas: {len(df_tab)} • Colunas: {list(df_tab.columns)}")
    edited = st.data_editor(df_tab, use_container_width=True, height=520, num_rows="dynamic", key=f"edit_{tab_name}")

    if st.button("💾 Salvar alterações"):
        try:
            ws = _book().worksheet(tab_name)
            ws.clear()
            values = [list(map(str, edited.columns))] + edited.fillna("").astype(str).values.tolist()
            ws.update("A1", values, value_input_option="USER_ENTERED")
            read_tab_df.clear()
            st.success("Alterações salvas!")
        except Exception as e:
            st.error(f"Erro ao salvar: {e}")

# =============== ROTEAMENTO ===============
if MENU == "📊 Dashboard":
    st.title("CRO1 — Dashboard de Vistorias")

    if not has_gsheets():
        st.warning("Configure o Google Sheets em `.streamlit/secrets.toml`.")
        st.stop()

    sh = _book()
    all_tabs = [ws.title for ws in sh.worksheets()]
    base_tab = st.selectbox("Fonte do Dashboard (aba):", all_tabs, index=0)

    try:
        df_base = read_tab_df(base_tab)
    except Exception as e:
        st.error(f"Falha ao ler a aba **{base_tab}**: {e}")
        st.stop()

    # Filtros (sidebar)
    with st.sidebar:
        st.markdown("### Filtros")
        filters = render_filters(df_base)

    df_f = apply_filters(df_base, filters)

    st.markdown(
        f'<span class="badge">Aba: {base_tab}</span>'
        f'<span class="badge">Registros: {len(df_f)}</span>',
        unsafe_allow_html=True
    )

    render_kpis(df_f, filters)
    render_charts(df_f, filters)
    render_table(df_f, filters)

else:
    render_editor()
