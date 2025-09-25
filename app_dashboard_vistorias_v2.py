# -*- coding: utf-8 -*-
# CRO1 — Dashboard & Editor (Google Sheets) — v2.0 PRO
# Tudo-em-um: Fases 1, 2 e 3 integradas

import warnings
warnings.filterwarnings("ignore", message=".*outside the limits for dates.*", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", message=".*Data Validation extension is not supported.*", category=UserWarning, module="openpyxl")

from __future__ import annotations
import json, time, unicodedata
from datetime import datetime, timedelta
from typing import Dict, Any, List, Tuple

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials
from streamlit_option_menu import option_menu

# ===================== CONFIG BÁSICA =====================
st.set_page_config(page_title="CRO1 — Dashboard PRO", layout="wide")

# Ajuste: tema claro/escuro (aplica só nos gráficos Plotly)
TEMPLATE_LIGHT = "plotly"
TEMPLATE_DARK  = "plotly_dark"

# TTL padrão de cache (s)
DEFAULT_TTL = 300

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Abas padrão do Sheets
TAB_BASE_DEFAULT = "ACOMPANHAMENTO VISTORIAS"
TAB_VALID_DEFAULT = "Validacao_de_Dados"
TAB_CONFIG_APP = "Config_App"     # para favoritos/ajustes persistentes

# ===================== UTIL / CACHE / LOGS =====================
def _norm(s: str) -> str:
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower().strip()

def perf_log(label: str, ms: float):
    st.session_state.setdefault("_perf_logs", []).append({"etapa": label, "ms": round(ms, 2)})

class CacheMgr:
    def __init__(self, default_ttl: int = DEFAULT_TTL):
        self.default_ttl = default_ttl

    def cached(self, ttl: int | None = None, show_spinner=False):
        ttl = ttl or self.default_ttl
        def deco(func):
            cf = st.cache_data(ttl=ttl, show_spinner=show_spinner)(func)
            def wrapper(*args, **kwargs):
                t0 = time.perf_counter()
                out = cf(*args, **kwargs)
                perf_log(func.__name__, (time.perf_counter() - t0)*1000)
                return out
            wrapper.clear = cf.clear
            return wrapper
        return deco

cache = CacheMgr()

# ===================== CONEXÃO GOOGLE SHEETS =====================
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

def read_ws_loose(ws, header_row=None) -> pd.DataFrame:
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    if header_row is None:
        hdr_idx = next((i for i, row in enumerate(values) if any(str(c).strip() for c in row)), 0)
    else:
        hdr_idx = max(0, int(header_row) - 1)
    headers = _make_unique_headers(values[hdr_idx])
    body = values[hdr_idx + 1:]
    while body and not any(str(c).strip() for c in body[-1]):
        body.pop()
    df = pd.DataFrame(body, columns=headers).replace("", pd.NA)
    return df

@cache.cached(ttl=DEFAULT_TTL, show_spinner=False)
def read_tab_df(tab_name: str) -> pd.DataFrame:
    t0 = time.perf_counter()
    ws = _book().worksheet(tab_name)
    df = read_ws_loose(ws)
    for c in df.columns:
        if "DATA" in str(c).upper():
            df[c] = pd.to_datetime(df[c], errors="coerce")
    perf_log(f"read_tab_df:{tab_name}", (time.perf_counter() - t0) * 1000)
    return df

def overwrite_tab_from_df(tab_name: str, df: pd.DataFrame, keep_header=True):
    ws = None
    try:
        ws = _book().worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = _book().add_worksheet(title=tab_name, rows=max(2000, len(df)+10), cols=max(10, len(df.columns)))
    else:
        ws.clear()

    if keep_header:
        values = [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist()
    else:
        values = df.fillna("").astype(str).values.tolist()

    ws.update("A1", values, value_input_option="USER_ENTERED")
    read_tab_df.clear()

# ===================== CONFIG/VALIDAÇÃO/DADOS HIERÁRQUICOS =====================
def col_or_none(df: pd.DataFrame, opts: List[str]) -> str | None:
    # exata
    for o in opts:
        for c in df.columns:
            if _norm(c) == _norm(o):
                return c
    # contém
    for o in opts:
        alvo = _norm(o)
        for c in df.columns:
            if alvo in _norm(c):
                return c
    return None

@cache.cached(ttl=300, show_spinner=False)  # 5 min
def load_validation_oms(tab_name: str = TAB_VALID_DEFAULT) -> pd.DataFrame:
    val = read_tab_df(tab_name)
    c_sigla = col_or_none(val, ["Sigla", "OM", "SIGLA"])
    c_nome  = col_or_none(val, ["Organização Militar", "Organizacao Militar", "Nome", "OM Completa"])
    c_dir   = col_or_none(val, [
        "Diretoria Responsável", "Diretoria Responsavel",
        "Órgãos de Direção Setorial", "Orgaos de Direcao Setorial",
        "Diretoria"
    ])
    if not all([c_sigla, c_nome, c_dir]):
        return pd.DataFrame(columns=["sigla", "nome", "diretoria"])
    df = (val[[c_sigla, c_nome, c_dir]]
          .rename(columns={c_sigla: "sigla", c_nome: "nome", c_dir: "diretoria"})
          .dropna(subset=["sigla", "nome"])
          .copy())
    df["sigla"] = df["sigla"].astype(str).str.strip()
    df["nome"] = df["nome"].astype(str).str.strip()
    df["diretoria"] = df["diretoria"].astype(str).str.strip()
    df = df.drop_duplicates(subset=["sigla", "nome", "diretoria"], keep="first").reset_index(drop=True)
    return df

# ===================== FAVORITOS (persistência no Sheets) =====================
def _ensure_config_ws() -> gspread.Worksheet | None:
    try:
        ws = _book().worksheet(TAB_CONFIG_APP)
    except gspread.WorksheetNotFound:
        ws = _book().add_worksheet(title=TAB_CONFIG_APP, rows=200, cols=5)
        ws.update("A1", [["chave", "valor_json"]])
    return ws

@cache.cached(ttl=60)
def load_favorites() -> Dict[str, Any]:
    try:
        ws = _ensure_config_ws()
        rows = ws.get_all_records()
        fav = {}
        for r in rows:
            k = str(r.get("chave","")).strip()
            v = r.get("valor_json","")
            if k:
                try:
                    fav[k] = json.loads(v) if isinstance(v, str) and v else v
                except Exception:
                    pass
        return fav
    except Exception:
        return {}

def save_favorite(name: str, payload: Dict[str, Any]):
    ws = _ensure_config_ws()
    rows = ws.get_all_records()
    data = []
    found = False
    for r in rows:
        k = str(r.get("chave","")).strip()
        if k == name:
            data.append([name, json.dumps(payload, ensure_ascii=False)])
            found = True
        else:
            data.append([k, json.dumps(r.get("valor_json", {}), ensure_ascii=False) if isinstance(r.get("valor_json"), dict) else r.get("valor_json","")])
    if not found:
        data.append([name, json.dumps(payload, ensure_ascii=False)])
    ws.clear()
    ws.update("A1", [["chave","valor_json"]] + data)
    load_favorites.clear()

def delete_favorite(name: str):
    ws = _ensure_config_ws()
    rows = ws.get_all_records()
    data = []
    for r in rows:
        k = str(r.get("chave","")).strip()
        if k and k != name:
            data.append([k, json.dumps(r.get("valor_json", {}), ensure_ascii=False) if isinstance(r.get("valor_json"), dict) else r.get("valor_json","")])
    ws.clear()
    ws.update("A1", [["chave","valor_json"]] + data)
    load_favorites.clear()

# ===================== KPI/GRÁFICOS/ALERTAS =====================
def kpis_block(df_f: pd.DataFrame, c_sit: str | None, c_dias_total: str | None,
               c_dias_exec: str | None, sla_dias: int,
               df_prev: pd.DataFrame | None = None, template="plotly"):
    colk1, colk2, colk3, colk4, colk5, colk6 = st.columns(6)
    total_vist = len(df_f)
    finalizadas = df_f[c_sit].astype(str).str.upper().str.contains("FINALIZADA").sum() if c_sit in df_f.columns else None
    pct_final = (finalizadas / total_vist * 100) if (finalizadas is not None and total_vist > 0) else 0
    prazo_medio_total = df_f[c_dias_total].mean() if c_dias_total in df_f.columns else None
    prazo_medio_exec  = df_f[c_dias_exec].mean() if c_dias_exec   in df_f.columns else None
    pct_sla = None
    if c_dias_total in df_f.columns and total_vist > 0:
        dentro_sla = (df_f[c_dias_total] <= sla_dias).sum()
        pct_sla = dentro_sla / total_vist * 100

    # Comparativos (vs período anterior)
    delta_total = None
    delta_final = None
    if df_prev is not None:
        tot_prev = len(df_prev)
        if tot_prev:
            delta_total = (total_vist - tot_prev) / max(tot_prev, 1) * 100
        if c_sit in df_prev.columns:
            fin_prev = df_prev[c_sit].astype(str).str.upper().str.contains("FINALIZADA").sum()
            pct_prev = fin_prev / max(tot_prev,1) * 100 if tot_prev else 0
            delta_final = pct_final - pct_prev

    with colk1: st.metric("Total de Vistorias", f"{total_vist:,}".replace(",", "."), None if delta_total is None else f"{delta_total:+.1f}%")
    with colk2: st.metric("Finalizadas (%)", f"{pct_final:,.1f}%", None if delta_final is None else f"{delta_final:+.1f} pp")
    with colk3: st.metric("Prazo médio total (dias)", f"{prazo_medio_total:,.1f}" if prazo_medio_total is not None else "—")
    with colk4: st.metric("Prazo médio execução (dias)", f"{prazo_medio_exec:,.1f}" if prazo_medio_exec is not None else "—")
    with colk5: st.metric(f"% dentro do SLA (≤{sla_dias}d)", f"{pct_sla:,.1f}%" if pct_sla is not None else "—")
    with colk6:
        st.write("")  # espaçamento
        st.caption("Comparativos calculados vs período anterior selecionado.")

def line_monthly(df: pd.DataFrame, col_data_base: str, template="plotly"):
    if col_data_base and df[col_data_base].notna().any():
        tmp = (df.groupby(pd.Grouper(key=col_data_base, freq="MS"))
               .size().reset_index(name="Vistorias"))
        fig = px.line(tmp, x=col_data_base, y="Vistorias", markers=True, title="Evolução Mensal de Vistorias", template=template)
        st.plotly_chart(fig, use_container_width=True)

def bar_by(df: pd.DataFrame, col: str, title: str, template="plotly"):
    if col in df.columns:
        tmp = df.groupby(col, as_index=False).size().sort_values("size", ascending=False)
        fig = px.bar(tmp, x=col, y="size", title=title, template=template)
        st.plotly_chart(fig, use_container_width=True)

def pie_by(df: pd.DataFrame, col: str, title: str, template="plotly"):
    if col in df.columns:
        tmp = df.groupby(col, as_index=False).size()
        fig = px.pie(tmp, names=col, values="size", hole=0.4, title=title, template=template)
        st.plotly_chart(fig, use_container_width=True)

def sla_by_col(df: pd.DataFrame, by_col: str, c_dias_total: str, sla_dias: int, title: str, template="plotly"):
    if by_col in df.columns and c_dias_total in df.columns:
        base = df.dropna(subset=[by_col, c_dias_total]).copy()
        base["Dentro SLA"] = base[c_dias_total] <= sla_dias
        tmp_sla = (base.groupby(by_col)["Dentro SLA"].mean()*100).reset_index(name="pct_sla")
        fig = px.bar(tmp_sla.sort_values("pct_sla"), x="pct_sla", y=by_col, orientation="h",
                     title=title, labels={"pct_sla": "% dentro do SLA"}, template=template)
        st.plotly_chart(fig, use_container_width=True)

def heatmap_mes_situacao(df: pd.DataFrame, col_data_base: str, c_sit: str, template="plotly"):
    if col_data_base and c_sit in df.columns and df[col_data_base].notna().any():
        aux = df.copy()
        aux["Mes"] = aux[col_data_base].dt.to_period("M").dt.to_timestamp()
        piv = (aux.groupby(["Mes", c_sit]).size().reset_index(name="Qtd")
               .pivot(index="Mes", columns=c_sit, values="Qtd").fillna(0))
        fig = px.imshow(piv.T, aspect="auto",
                        labels=dict(x="Mês", y="Situação", color="Qtd"),
                        title="Heatmap — Mês x Situação", template=template)
        st.plotly_chart(fig, use_container_width=True)

def timeline_vistorias(df: pd.DataFrame, c_inicio: str, c_fim: str | None, c_rotulo: str, template="plotly"):
    if not c_inicio or c_inicio not in df.columns or df.empty:
        return
    base = df.copy()
    base["start"] = pd.to_datetime(base[c_inicio], errors="coerce")
    if c_fim and c_fim in df.columns:
        base["finish"] = pd.to_datetime(base[c_fim], errors="coerce")
    else:
        # se não existir, tenta inferir com dias de execução ou 1 dia
        if "QUANTIDADE DE DIAS PARA EXECUCAO" in base.columns:
            base["finish"] = base["start"] + pd.to_timedelta(pd.to_numeric(base["QUANTIDADE DE DIAS PARA EXECUCAO"], errors="coerce").fillna(1), unit="D")
        else:
            base["finish"] = base["start"] + pd.to_timedelta(1, unit="D")
    base = base.dropna(subset=["start", "finish"])
    if base.empty:
        return
    label = c_rotulo if (c_rotulo in base.columns) else None
    fig = px.timeline(base, x_start="start", x_end="finish", y=label, title="Timeline / Gantt das Vistorias", template=template)
    fig.update_yaxes(autorange="reversed")
    st.plotly_chart(fig, use_container_width=True)

def alerts_block(df: pd.DataFrame, c_dias_total: str | None, c_sit: str | None, sla_dias: int):
    if not c_dias_total or c_dias_total not in df.columns or df.empty:
        return
    base = df.copy()
    base["dias_total"] = pd.to_numeric(base[c_dias_total], errors="coerce")
    crit = base[(base["dias_total"] > sla_dias) & (~base[c_sit].astype(str).str.upper().str.contains("FINALIZADA") if c_sit in base.columns else True)]
    near = base[(base["dias_total"] >= 0.8 * sla_dias) & (base["dias_total"] <= sla_dias)]
    c1, c2 = st.columns(2)
    with c1:
        st.warning(f"⚠️ Fora do SLA: {len(crit)} registros")
        if not crit.empty:
            st.dataframe(crit.head(30), use_container_width=True)
    with c2:
        st.info(f"⏳ Em risco (≥80% do SLA): {len(near)} registros")
        if not near.empty:
            st.dataframe(near.head(30), use_container_width=True)

# ===================== UI FILTERS (hierárquico + busca + favoritos) =====================
def render_hierarchical_filters(df_oms: pd.DataFrame):
    # Diretoria
    if not df_oms.empty and df_oms["diretoria"].notna().any():
        diretorias = sorted([d for d in df_oms["diretoria"].dropna().astype(str).unique().tolist() if d.strip()])
    else:
        diretorias = []
    dir_opcoes = ["(Todas)"] + diretorias
    dir_sel = st.selectbox("🏢 Diretoria Responsável", dir_opcoes, index=0, key="f_dir")

    # universo
    if dir_sel and dir_sel != "(Todas)":
        df_oms_visiveis = df_oms[df_oms["diretoria"] == dir_sel].copy()
    else:
        df_oms_visiveis = df_oms.copy()

    busca_om = st.text_input("🔎 Buscar OM (sigla ou nome)", placeholder="Digite parte da sigla ou do nome...")

    def _match_busca(row) -> bool:
        if not busca_om:
            return True
        q = _norm(busca_om)
        return q in _norm(row["sigla"]) or q in _norm(row["nome"])

    if not df_oms_visiveis.empty:
        df_oms_visiveis = df_oms_visiveis[df_oms_visiveis.apply(_match_busca, axis=1)]

    om_labels = [f"{r.sigla} — {r.nome}" for r in df_oms_visiveis.itertuples(index=False)]
    label_to_sigla = {label: label.split(" — ", 1)[0] for label in om_labels}

    om_sel_labels = st.multiselect(
        "🏛️ OM Apoiadora",
        options=om_labels,
        default=[],
        placeholder="Digite para buscar (auto-complete)",
    )
    oms_siglas_sel = [label_to_sigla[l] for l in om_sel_labels]
    return dict(dir_sel=dir_sel, busca_om=busca_om, om_sel_labels=om_sel_labels,
                oms_siglas_sel=oms_siglas_sel, df_oms_visiveis=df_oms_visiveis)

def apply_all_filters(df: pd.DataFrame, filters: Dict[str, Any], df_oms: pd.DataFrame):
    dff = df.copy()
    col_data_base = filters.get("col_data_base")
    if filters.get("periodo") and col_data_base in dff.columns:
        ini, fim = filters["periodo"]
        dff = dff[(dff[col_data_base] >= pd.to_datetime(ini)) & (dff[col_data_base] <= pd.to_datetime(fim))]
    c_dir, c_sit, c_urg, c_om = filters.get("c_dir"), filters.get("c_sit"), filters.get("c_urg"), filters.get("c_om")
    if filters.get("dir_manual") and c_dir in dff.columns:
        dff = dff[dff[c_dir].astype(str).isin(filters["dir_manual"])]
    if filters.get("sit") and c_sit in dff.columns:
        dff = dff[dff[c_sit].astype(str).isin(filters["sit"])]
    if filters.get("urg") and c_urg in dff.columns:
        dff = dff[dff[c_urg].astype(str).isin(filters["urg"])]

    oms_siglas_sel = filters.get("om_siglas", [])
    if c_om in dff.columns and oms_siglas_sel:
        col_series = dff[c_om].astype(str).fillna("").str.strip()
        sigla_to_nome = {r.sigla: r.nome for r in df_oms.itertuples(index=False)} if not df_oms.empty else {}
        nomes_sel = {sigla_to_nome.get(s, "") for s in oms_siglas_sel if sigla_to_nome.get(s, "")}
        mask = col_series.isin(oms_siglas_sel) | col_series.isin(list(nomes_sel))
        dff = dff[mask]
    return dff

# ===================== SIDEBAR =====================
with st.sidebar:
    st.write("🔌 Google Sheets:", "ON ✅" if has_gsheets() else "OFF ❌")
    if not has_gsheets():
        st.error("Secrets não detectado. Configure `.streamlit/secrets.toml` e [gsheets][spreadsheet_url].")

    theme = st.toggle("🌗 Tema escuro (gráficos)", value=False)
    PLOT_TEMPLATE = TEMPLATE_DARK if theme else TEMPLATE_LIGHT

    if st.button("🧹 Limpar caches"):
        st.cache_data.clear()
        st.session_state["_perf_logs"] = []
        st.success("Caches limpos.")
        st.rerun()

    MENU = option_menu(
        "CRO1 — Sistema",
        ["📊 Dashboard Principal", "🗂️ Editor da Planilha", "🧩 Favoritos", "⚙️ Configurações"],
        icons=["bar-chart", "table", "star", "gear"],
        default_index=0,
        menu_icon="grid"
    )

# ===================== EDITOR =====================
if MENU == "🗂️ Editor da Planilha":
    st.title("🗂️ Editor da Planilha (Google Sheets)")
    if not has_gsheets():
        st.stop()
    sh = _book()
    tabs = [ws.title for ws in sh.worksheets()]
    st.success("Google Sheets conectado ✅")
    st.caption(f"Planilha: {st.secrets['gsheets']['spreadsheet_url']}")
    tab_name = st.selectbox("Escolha a aba para visualizar/editar:", tabs, index=0)
    if st.button("↻ Recarregar (limpar cache)"):

        st.cache_data.clear()
        st.rerun()
    try:
        df_tab = read_tab_df(tab_name)
    except Exception as e:
        st.error(f"Falha ao ler a aba **{tab_name}**: {e}")
        st.stop()
    st.caption(f"Linhas: {len(df_tab)} • Colunas: {list(df_tab.columns)}")
    edited_df = st.data_editor(df_tab, use_container_width=True, num_rows="dynamic",
                               key=f"editor_{tab_name}", height=520)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Salvar alterações na aba"):
            try:
                out = edited_df.copy()
                for c in out.columns:
                    if "DATA" in str(c).upper():
                        out[c] = pd.to_datetime(out[c], errors="coerce").dt.strftime("%Y-%m-%d")
                overwrite_tab_from_df(tab_name, out, keep_header=True)
                st.success(f"Alterações salvas em **{tab_name}**.")
            except Exception as e:
                st.error(f"Falha ao salvar: {e}")
    with c2:
        st.download_button("⬇️ Baixar CSV desta aba",
                           edited_df.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"{tab_name}.csv", mime="text/csv")

# ===================== FAVORITOS (salvar/carregar) =====================
if MENU == "🧩 Favoritos":
    st.title("🧩 Favoritos de Filtros")
    if not has_gsheets():
        st.warning("Conecte o Google Sheets para persistir favoritos.")
        st.stop()
    favs = load_favorites()
    if not favs:
        st.info("Você ainda não salvou favoritos. Volte ao Dashboard, aplique os filtros e salve aqui.")
    else:
        nomes = sorted(list(favs.keys()))
        pick = st.selectbox("Escolha um favorito:", ["(selecione)"]+nomes)
        if pick and pick != "(selecione)":
            st.code(json.dumps(favs[pick], ensure_ascii=False, indent=2))
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("❌ Apagar este favorito"):
                    delete_favorite(pick)
                    st.success("Favorito removido.")
                    st.rerun()
            with col_b:
                st.caption("Para aplicar no dashboard, copie/cole os parâmetros manualmente no momento.")

# ===================== DASHBOARD =====================
if MENU == "📊 Dashboard Principal":
    st.title("📊 Dashboard Principal")
    if not has_gsheets():
        st.warning("Ative o Google Sheets para carregar dashboards.")
        st.stop()

    # Escolha de aba base
    base_tab = st.selectbox("Escolha a aba (fonte dos gráficos/KPIs):",
                            [ws.title for ws in _book().worksheets()],
                            index=0, key="dashboard_tab")
    try:
        df = read_tab_df(base_tab)
    except Exception as e:
        st.error(f"Falha ao ler a aba **{base_tab}**: {e}")
        st.stop()
    if df.empty:
        st.warning("A aba está vazia.")
        st.stop()

    # Mapear colunas do DF base
    c_obj = col_or_none(df, ["OBJETO DE VISTORIA", "OBJETO"])
    c_om  = col_or_none(df, ["OM APOIADA", "OM APOIADORA", "OM"])
    c_dir = col_or_none(df, ["Diretoria Responsavel", "Diretoria Responsável", "Diretoria"])
    c_urg = col_or_none(df, ["Classificacao da Urgencia", "Classificação da Urgência", "Urgencia"])
    c_sit = col_or_none(df, ["Situacao", "Situação"])
    c_data_solic = col_or_none(df, ["DATA DA SOLICITACAO", "DATA DA SOLICITAÇÃO"])
    c_data_vist  = col_or_none(df, ["DATA DA VISTORIA"])
    c_dias_total = col_or_none(df, ["QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO"])
    c_dias_exec  = col_or_none(df, ["QUANTIDADE DE DIAS PARA EXECUCAO", "QUANTIDADE DE DIAS PARA EXECUÇÃO"])

    for c in [c_data_solic, c_data_vist]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in [c_dias_total, c_dias_exec]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    st.caption(f"Aba: **{base_tab}** • Linhas: {len(df)}")

    # ====== Filtros com OMs da aba de validação ======
    df_oms = load_validation_oms(TAB_VALID_DEFAULT)
    col_btn, col_info = st.sidebar.columns([1, 1.4])
    with col_btn:
        if st.button("🔁 Atualizar lista de OMs"):
            load_validation_oms.clear()
            st.rerun()
    with col_info:
        st.caption("Lista de OMs com cache de 5 min")

    with st.sidebar:
        st.subheader("Filtros — Diretoria → OM")
        h = render_hierarchical_filters(df_oms)

    # Período e demais filtros
    st.sidebar.subheader("Filtros adicionais")
    col_data_base = c_data_solic if c_data_solic in df.columns else c_data_vist
    if col_data_base and df[col_data_base].notna().any():
        min_dt = pd.to_datetime(df[col_data_base].min()).date()
        max_dt = pd.to_datetime(df[col_data_base].max()).date()
        periodo = st.sidebar.date_input("Período (pela data da solicitação)", value=(min_dt, max_dt))
    else:
        periodo = None

    def opts(series):
        try: return sorted(series.dropna().astype(str).unique().tolist())
        except Exception: return sorted(list({str(x) for x in series.dropna().tolist()}))

    dir_sel_manual = st.sidebar.multiselect("Diretoria (planilha)", opts(df[c_dir]) if c_dir in df.columns else [])
    sit_sel = st.sidebar.multiselect("Situação", opts(df[c_sit]) if c_sit in df.columns else [])
    urg_sel = st.sidebar.multiselect("Classificação de Urgência", opts(df[c_urg]) if c_urg in df.columns else [])
    sla_dias = st.sidebar.number_input("SLA (dias p/ 'dentro do prazo')", 1, 365, value=30)

    # ====== Favoritos (salvar/aplicar) ======
    st.sidebar.subheader("⭐ Favoritos")
    favs = load_favorites() if has_gsheets() else {}
    # Salvar
    fav_name = st.sidebar.text_input("Nome do favorito")
    if st.sidebar.button("💾 Salvar favorito"):
        payload = dict(
            base_tab=base_tab,
            dir_sel=h["dir_sel"],
            om_siglas=h["oms_siglas_sel"],
            dir_manual=dir_sel_manual,
            sit=sit_sel,
            urg=urg_sel,
            periodo=[str(periodo[0]), str(periodo[1])] if periodo else None,
            sla_dias=sla_dias,
        )
        if has_gsheets() and fav_name.strip():
            save_favorite(fav_name.strip(), payload)
            st.sidebar.success("Favorito salvo!")
        else:
            st.sidebar.warning("Informe um nome e confirme conexão com Sheets.")

    # Aplicar
    if favs:
        chosen = st.sidebar.selectbox("Aplicar favorito:", ["(nenhum)"] + sorted(list(favs.keys())))
        if chosen and chosen != "(nenhum)":
            data = favs[chosen]
            try:
                if data.get("base_tab") and data["base_tab"] in [ws.title for ws in _book().worksheets()]:
                    if data["base_tab"] != base_tab:
                        st.warning("Este favorito é de outra aba. Troque a aba para aplicar totalmente.")
                if isinstance(data.get("sla_dias"), int):
                    sla_dias = data["sla_dias"]
                # Obs.: aplicar totalmente em UI exigiria st.session_state; mantemos como referência.
                st.sidebar.info("Favorito carregado (consulte parâmetros exibidos).")
                st.sidebar.code(json.dumps(data, ensure_ascii=False, indent=2))
            except Exception:
                st.sidebar.warning("Não foi possível aplicar completamente este favorito.")

    # ====== Aplicação dos filtros
    df_f = apply_all_filters(df, {
        "periodo": periodo,
        "dir_manual": dir_sel_manual,
        "sit": sit_sel,
        "urg": urg_sel,
        "om_siglas": h["oms_siglas_sel"],
        "c_dir": c_dir, "c_sit": c_sit, "c_urg": c_urg, "c_om": c_om,
        "col_data_base": col_data_base,
    }, df_oms)

    # ====== Período anterior (comparativos)
    df_prev = None
    if periodo and col_data_base:
        ini, fim = periodo
        delta = (pd.to_datetime(fim) - pd.to_datetime(ini)).days or 30
        prev_ini = pd.to_datetime(ini) - pd.Timedelta(days=delta)
        prev_fim = pd.to_datetime(ini) - pd.Timedelta(days=1)
        df_prev = apply_all_filters(df, {
            "periodo": (prev_ini.date(), prev_fim.date()),
            "dir_manual": dir_sel_manual,
            "sit": sit_sel,
            "urg": urg_sel,
            "om_siglas": h["oms_siglas_sel"],
            "c_dir": c_dir, "c_sit": c_sit, "c_urg": c_urg, "c_om": c_om,
            "col_data_base": col_data_base,
        }, df_oms)

    # ====== KPIs (com comparativos)
    kpis_block(df_f, c_sit, c_dias_total, c_dias_exec, sla_dias, df_prev=df_prev, template=PLOT_TEMPLATE)
    st.divider()

    # ====== Visualizações avançadas
    line_monthly(df_f, col_data_base, template=PLOT_TEMPLATE)
    bar_by(df_f, c_dir, "Vistorias por Diretoria Responsável", template=PLOT_TEMPLATE)
    pie_by(df_f, c_sit, "Distribuição por Situação", template=PLOT_TEMPLATE)
    bar_by(df_f, c_urg, "Vistorias por Classificação de Urgência", template=PLOT_TEMPLATE)
    sla_by_col(df_f, c_dir, c_dias_total, sla_dias, "% Dentro do SLA (≤ SLA) por Diretoria", template=PLOT_TEMPLATE)
    # SLA por OM (se a coluna OM estiver coerente com siglas/nomes)
    sla_by_col(df_f, c_om, c_dias_total, sla_dias, "% Dentro do SLA (≤ SLA) por OM", template=PLOT_TEMPLATE)
    heatmap_mes_situacao(df_f, col_data_base, c_sit, template=PLOT_TEMPLATE)
    timeline_vistorias(df_f, c_data_solic or c_data_vist, c_data_vist, c_obj or c_om or c_dir, template=PLOT_TEMPLATE)

    # ====== Alertas
    st.subheader("🔔 Alertas")
    alerts_block(df_f, c_dias_total, c_sit, sla_dias)

    # ====== Relatório (Exportar HTML p/ PDF do navegador)
    st.subheader("📄 Relatório")
    html_report = f"""
    <html>
      <head><meta charset="utf-8"><title>Relatório CRO1</title></head>
      <body>
        <h2>Relatório — CRO1</h2>
        <p><b>Aba base:</b> {base_tab} | <b>Gerado em:</b> {datetime.now():%Y-%m-%d %H:%M}</p>
        <p>Filtros: Diretoria UI: {h['dir_sel']} | OMs: {', '.join(h['oms_siglas_sel']) or '(todas)'} |
        Situação: {', '.join(sit_sel) or '(todas)'} | Urgência: {', '.join(urg_sel) or '(todas)'} |
        Período: {str(periodo[0]) + ' a ' + str(periodo[1]) if periodo else '(não aplicado)'} |
        SLA: {sla_dias} dias</p>
      </body>
    </html>
    """.encode("utf-8")
    st.download_button("⬇️ Baixar Relatório (HTML)", html_report, file_name="relatorio_cro1.html", mime="text/html")

    # ====== Tabela final
    st.subheader("Detalhamento (mais recentes)")
    ord_col = col_data_base if col_data_base else (c_data_vist if c_data_vist in df_f.columns else None)
    df_show = df_f.sort_values(ord_col, ascending=False).head(100) if ord_col else df_f.head(100)
    st.dataframe(df_show, use_container_width=True)

# ===================== CONFIGURAÇÕES =====================
if MENU == "⚙️ Configurações":
    st.title("⚙️ Configurações & Monitoramento")
    st.write("TTL padrão do cache:", DEFAULT_TTL, "s")
    st.write("Abas padrão:", {"base": TAB_BASE_DEFAULT, "validacao": TAB_VALID_DEFAULT, "config": TAB_CONFIG_APP})
    st.subheader("Logs de performance (sessão)")
    logs = st.session_state.get("_perf_logs", [])
    if not logs:
        st.info("Sem logs nesta sessão ainda.")
    else:
        st.dataframe(pd.DataFrame(logs))
