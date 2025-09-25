import warnings
warnings.filterwarnings("ignore", message=".*outside the limits for dates.*", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", message=".*Data Validation extension is not supported and will be removed.*", category=UserWarning, module="openpyxl")

import unicodedata
from datetime import datetime
import re

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials
from streamlit_option_menu import option_menu

# =========================================================
# CONFIGURAÇÃO GERAL
# =========================================================
st.set_page_config(
    page_title="CRO1 — Editor & Dashboards (Sheets)",
    layout="wide",
    initial_sidebar_state="expanded"
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# =========================================================
# CONEXÃO GOOGLE SHEETS
# =========================================================
def has_gsheets() -> bool:
    return (
        "gcp_service_account" in st.secrets
        and "gsheets" in st.secrets
        and "spreadsheet_url" in st.secrets["gsheets"]
        and bool(st.secrets["gsheets"]["spreadsheet_url"])
    )

@st.cache_resource(show_spinner=False)
def get_gs_client():
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def get_workbook():
    return get_gs_client().open_by_url(st.secrets["gsheets"]["spreadsheet_url"])

# =========================================================
# LEITURA TOLERANTE A CABEÇALHO
# =========================================================
def make_unique_headers(raw_headers):
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

def read_worksheet_safe(ws, header_row=None) -> pd.DataFrame:
    try:
        values = ws.get_all_values()
        if not values:
            return pd.DataFrame()

        if header_row is None:
            hdr_idx = next((i for i, row in enumerate(values) if any(str(c).strip() for c in row)), 0)
        else:
            hdr_idx = max(0, int(header_row) - 1)

        headers = make_unique_headers(values[hdr_idx])
        body = values[hdr_idx + 1:]

        while body and not any(str(c).strip() for c in body[-1]):
            body.pop()

        df = pd.DataFrame(body, columns=headers).replace("", pd.NA)
        return df
    except Exception as e:
        st.error(f"Erro ao ler worksheet: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=60, show_spinner=False)
def read_tab_df(tab_name: str) -> pd.DataFrame:
    try:
        ws = get_workbook().worksheet(tab_name)
        df = read_worksheet_safe(ws)
        for col in df.columns:
            if "DATA" in col.upper():
                df[col] = pd.to_datetime(df[col], errors="coerce")
        return df
    except Exception as e:
        st.error(f"Erro ao ler aba {tab_name}: {e}")
        return pd.DataFrame()

def overwrite_tab_from_df(tab_name: str, df: pd.DataFrame, keep_header=True) -> bool:
    try:
        sh = get_workbook()
        try:
            ws = sh.worksheet(tab_name)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(
                title=tab_name,
                rows=max(2000, len(df) + 10),
                cols=max(10, len(df.columns))
            )
        ws.clear()
        values = [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist() if keep_header \
                 else df.fillna("").astype(str).values.tolist()
        ws.update("A1", values, value_input_option="USER_ENTERED")
        read_tab_df.clear()
        return True
    except Exception as e:
        st.error(f"Erro ao salvar aba {tab_name}: {e}")
        return False

# =========================================================
# HELPERS
# =========================================================
def normalize_text(s: str) -> str:
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower().strip()

def find_column(df: pd.DataFrame, options) -> str | None:
    cols = list(df.columns)
    for opt in options:
        for col in cols:
            if normalize_text(col) == normalize_text(opt):
                return col
    for opt in options:
        target = normalize_text(opt)
        for col in cols:
            if target in normalize_text(col):
                return col
    return None

def create_card(title: str):
    st.markdown(
        f"<div style='padding:12px 16px;border-radius:12px;background:linear-gradient(135deg, #667eea 0%, #764ba2 100%);"
        f"border:none;color:white;font-weight:700;font-size:20px;text-align:center;margin:10px 0;'>{title}</div>",
        unsafe_allow_html=True
    )

def get_filter_options(series):
    try:
        return sorted(series.dropna().astype(str).unique().tolist())
    except Exception:
        return sorted(list({str(x) for x in series.dropna().tolist()}))

# =========================================================
# VALIDACAO_DE_DADOS → OMs com Diretoria (cache 5 min)
# =========================================================
@st.cache_data(ttl=300, show_spinner=False)
def load_oms_validation_data() -> pd.DataFrame:
    dfv = read_tab_df("Validacao_de_Dados")
    if dfv.empty:
        return pd.DataFrame()

    # procurar colunas preferenciais
    pref_sigla = ["Sigla", "OM", "Sigla OM"]
    pref_nome  = ["Organização Militar", "Organizacao Militar", "Nome", "OM Nome"]
    pref_dir   = ["Diretoria Responsável", "Diretoria Responsavel", "Diretoria"]

    c_sigla = find_column(dfv, pref_sigla)
    c_nome  = find_column(dfv, pref_nome)
    c_dir   = find_column(dfv, pref_dir)

    # fallback por posição (B,C,D,M) somente se nada foi encontrado
    cols = list(dfv.columns)
    if not c_sigla and len(cols) > 1: c_sigla = cols[1]
    if not c_nome  and len(cols) > 2: c_nome  = cols[2]
    if not c_dir   and len(cols) > 3: c_dir   = cols[3]

    if not c_sigla:
        return pd.DataFrame()

    df_oms = (
        dfv[[c_sigla] + ([c_nome] if c_nome else []) + ([c_dir] if c_dir else [])]
        .rename(columns={c_sigla: "sigla", c_nome or c_sigla: "nome_completo", c_dir or c_sigla: "diretoria"})
        .fillna({"nome_completo": "" , "diretoria": ""})
    )

    df_oms["sigla"] = df_oms["sigla"].astype(str).str.strip()
    df_oms["nome_completo"] = df_oms["nome_completo"].astype(str).str.strip()
    df_oms["diretoria"] = df_oms["diretoria"].astype(str).str.strip()
    df_oms = df_oms[df_oms["sigla"].str.len() >= 2].drop_duplicates(subset=["sigla"])
    df_oms["display_name"] = df_oms.apply(
        lambda r: f"{r['sigla']} - {r['nome_completo']}" if r["nome_completo"] and r["nome_completo"] != r["sigla"] else r["sigla"],
        axis=1
    )
    df_oms["search_text"] = (df_oms["sigla"] + " " + df_oms["nome_completo"]).str.upper()
    return df_oms.sort_values(["diretoria", "sigla"])

def create_om_filter_component(df_oms: pd.DataFrame, key_suffix: str = ""):
    if df_oms.empty:
        st.sidebar.warning("⚠️ Lista de OMs não disponível")
        return [], []

    diretorias = ["Todas"] + sorted(df_oms["diretoria"].dropna().unique().tolist())
    diretoria_sel = st.sidebar.selectbox("🏢 Diretoria Responsável", diretorias, key=f"dir_filter_{key_suffix}")

    if diretoria_sel == "Todas":
        base = df_oms
    else:
        base = df_oms[df_oms["diretoria"] == diretoria_sel]

    busca = st.sidebar.text_input("🔎 Buscar OM (sigla ou nome)", key=f"om_search_{key_suffix}", placeholder="Digite para buscar...")
    if busca:
        up = busca.upper()
        base = base[base["search_text"].str.contains(up, na=False)]

    opcoes = base["display_name"].tolist()
    selecionadas = st.sidebar.multiselect(
        f"🏛️ OM Apoiadora ({len(opcoes)} encontradas)",
        opcoes,
        key=f"om_multi_{key_suffix}"
    )

    siglas = []
    if selecionadas:
        m = base.set_index("display_name")["sigla"]
        for item in selecionadas:
            if item in m.index:
                siglas.append(m.loc[item])

    return siglas, ([diretoria_sel] if diretoria_sel != "Todas" else [])

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown("### 🔌 Status da Conexão")
    if has_gsheets():
        st.success("Google Sheets: ✅ Conectado")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 Cache Geral", use_container_width=True):
                read_tab_df.clear(); load_oms_validation_data.clear()
                st.success("Cache limpo!")
                st.rerun()
        with c2:
            if st.button("📋 Atualizar OMs", use_container_width=True):
                load_oms_validation_data.clear()
                st.success("Lista de OMs atualizada!")
                st.rerun()

        try:
            _oms = load_oms_validation_data()
            st.info(f"📋 {len(_oms)} OMs carregadas" if not _oms.empty else "⚠️ Lista de OMs vazia")
        except Exception as e:
            st.error(f"❌ Erro ao carregar OMs: {e}")
    else:
        st.error("Google Sheets: ❌ Desconectado")
        st.warning("Configure o arquivo `.streamlit/secrets.toml`")

    st.markdown("---")
    MENU = option_menu(
        "🚀 CRO1 Sistema",
        ["🗂️ Editor de Planilha", "📊 Dashboards"],
        icons=["table", "bar-chart-fill"],
        default_index=0,
        menu_icon="grid-3x3-gap-fill",
        styles={
            "container": {"padding": "5px", "background-color": "#fafafa"},
            "icon": {"color": "#667eea", "font-size": "18px"},
            "nav-link": {"font-size": "16px", "text-align": "left", "margin": "0px"},
            "nav-link-selected": {"background-color": "#667eea"},
        }
    )

# =========================================================
# 1) EDITOR
# =========================================================
if MENU == "🗂️ Editor de Planilha":
    create_card("🗂️ Editor de Planilha Google Sheets")
    if not has_gsheets():
        st.error("❌ Google Sheets não configurado. Verifique o arquivo secrets.toml")
        st.stop()

    try:
        workbook = get_workbook()
        tabs = [ws.title for ws in workbook.worksheets()]
        st.success("✅ Google Sheets conectado!")
        st.caption(f"📋 **Planilha:** {st.secrets['gsheets']['spreadsheet_url']}")

        c1, c2 = st.columns([3, 1])
        with c1:
            tab_name = st.selectbox("📂 Escolha a aba para visualizar/editar:", tabs, index=0)
        with c2:
            if st.button("↻ Recarregar", use_container_width=True, key="btn_recarregar_editor"):
                read_tab_df.clear(); st.rerun()

        df_tab = read_tab_df(tab_name)
        if df_tab.empty:
            st.warning("⚠️ A aba selecionada está vazia.")
        else:
            st.info(f"📊 **Linhas:** {len(df_tab):,} • **Colunas:** {len(df_tab.columns)}")
            edited_df = st.data_editor(
                df_tab, use_container_width=True, num_rows="dynamic",
                key=f"editor_{tab_name}", height=500, hide_index=True
            )
            c1, c2, c3 = st.columns([2, 2, 2])
            with c1:
                if st.button("💾 Salvar Alterações", use_container_width=True):
                    with st.spinner("Salvando..."):
                        df_save = edited_df.copy()
                        for col in df_save.columns:
                            if "DATA" in col.upper():
                                df_save[col] = pd.to_datetime(df_save[col], errors="coerce")
                                df_save[col] = df_save[col].dt.strftime("%Y-%m-%d")
                        df_save = df_save.fillna("")
                        if overwrite_tab_from_df(tab_name, df_save):
                            st.success("✅ Alterações salvas!")
                            read_tab_df.clear(); st.rerun()
                        else:
                            st.error("❌ Erro ao salvar alterações.")
            with c2:
                if st.button("➕ Adicionar Linha", use_container_width=True):
                    st.warning("Use o editor para inserir linhas (num_rows='dynamic' já permite).")
            with c3:
                st.caption("Para excluir, baixe CSV, edite e reenvie (ou implemente coluna 'Excluir?').")
    except Exception as e:
        st.error(f"❌ Erro no Editor de Planilha: {e}")

# =========================================================
# 2) DASHBOARDS
# =========================================================
if MENU == "📊 Dashboards":
    create_card("📊 Dashboards de Vistorias")
    if not has_gsheets():
        st.error("❌ Google Sheets não configurado. Verifique o arquivo secrets.toml")
        st.stop()

    try:
        workbook = get_workbook()
        tabs = [ws.title for ws in workbook.worksheets()]
        base_tab = st.selectbox("📂 Escolha a aba para análise:", tabs, index=0, key="dashboard_tab")
        df = read_tab_df(base_tab)
        if df.empty:
            st.warning("⚠️ A aba selecionada está vazia."); st.stop()

        st.info(f"📊 Analisando **{base_tab}** • **{len(df):,}** registros")

        col_mappings = {
            'objeto':           find_column(df, ["OBJETO DE VISTORIA", "OBJETO"]),
            'om':               find_column(df, ["OM APOIADA", "OM APOIADORA", "OM"]),
            'diretoria':        find_column(df, ["Diretoria Responsavel", "Diretoria Responsável", "Diretoria"]),
            'urgencia':         find_column(df, ["Classificacao da Urgencia", "Classificação da Urgência", "Urgencia"]),
            'situacao':         find_column(df, ["Situacao", "Situação"]),
            'data_solicitacao': find_column(df, ["DATA DA SOLICITACAO", "DATA DA SOLICITAÇÃO"]),
            'data_vistoria':    find_column(df, ["DATA DA VISTORIA"]),
            'dias_total':       find_column(df, ["QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO"]),
            'dias_execucao':    find_column(df, ["QUANTIDADE DE DIAS PARA EXECUCAO", "QUANTIDADE DE DIAS PARA EXECUÇÃO"]),
            'status':           find_column(df, ["STATUS - ATUALIZACAO SEMANAL", "STATUS - ATUALIZAÇÃO SEMANAL", "Status", "VT EXECUTADA POR"])
        }

        with st.expander("🔍 Debug - Colunas Mapeadas", expanded=False):
            c1, c2 = st.columns(2)
            with c1: st.write("**Colunas disponíveis:**"); st.write(list(df.columns))
            with c2: st.write("**Mapeamento encontrado:**"); st.write({k:v for k,v in col_mappings.items() if v})

        # Conversões
        for col in [col_mappings.get('data_solicitacao'), col_mappings.get('data_vistoria')]:
            if col and col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
        for col in [col_mappings.get('dias_total'), col_mappings.get('dias_execucao')]:
            if col and col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        # Filtros
        st.sidebar.markdown("### 🔍 Filtros")
        df_oms = load_oms_validation_data()

        col_data_base = col_mappings.get('data_solicitacao') or col_mappings.get('data_vistoria')
        periodo = None
        if col_data_base and col_data_base in df.columns and df[col_data_base].notna().any():
            min_dt = pd.to_datetime(df[col_data_base].min()).date()
            max_dt = pd.to_datetime(df[col_data_base].max()).date()
            periodo = st.sidebar.date_input("📅 Período", value=(min_dt, max_dt), min_value=min_dt, max_value=max_dt)

        if not df_oms.empty:
            oms_selecionadas, diretorias_selecionadas = create_om_filter_component(df_oms, "dashboard")
        else:
            st.sidebar.warning("⚠️ Lista de OMs não disponível - usando filtro manual")
            oms_selecionadas, diretorias_selecionadas = [], []

        filtros = {}
        for key, label in [('situacao',"📋 Situação"), ('urgencia',"⚡ Urgência")]:
            col_name = col_mappings.get(key)
            if col_name and col_name in df.columns:
                options = get_filter_options(df[col_name])
                if options:
                    filtros[key] = st.sidebar.multiselect(label, options, key=f"filter_{key}")

        sla_dias = st.sidebar.number_input("⏱️ SLA (dias)", min_value=1, max_value=365, value=30)

        # Aplicar filtros
        df_f = df.copy()
        if periodo and len(periodo) == 2 and col_data_base:
            ini, fim = periodo
            df_f = df_f[(df_f[col_data_base] >= pd.to_datetime(ini)) & (df_f[col_data_base] <= pd.to_datetime(fim))]
        col_om = col_mappings.get('om')
        if oms_selecionadas and col_om and col_om in df_f.columns:
            mask_om = df_f[col_om].astype(str).str.upper().isin([om.upper() for om in oms_selecionadas])
            if not mask_om.any():
                pattern = '|'.join([re.escape(om.upper()) for om in oms_selecionadas])
                mask_om = df_f[col_om].astype(str).str.upper().str.contains(pattern, na=False, regex=True)
            df_f = df_f[mask_om]
        col_dir = col_mappings.get('diretoria')
        if diretorias_selecionadas and col_dir and col_dir in df_f.columns:
            df_f = df_f[df_f[col_dir].astype(str).isin(diretorias_selecionadas)]
        for k, col_name in col_mappings.items():
            if k in ['om','diretoria']: 
                continue
            sel = filtros.get(k)
            if sel and col_name and col_name in df_f.columns:
                df_f = df_f[df_f[col_name].astype(str).isin(sel)]

        # KPIs
        st.markdown("### 📈 Indicadores Principais")
        c1, c2, c3, c4, c5 = st.columns(5)
        total = len(df_f)
        finalizadas = 0; pct_final = 0
        if col_mappings.get('situacao') and col_mappings['situacao'] in df_f.columns:
            finalizadas = df_f[col_mappings['situacao']].astype(str).str.upper().str.contains('FINALIZADA', na=False).sum()
            pct_final = (finalizadas/total*100) if total>0 else 0
        prazo_total = df_f[col_mappings['dias_total']].mean() if col_mappings.get('dias_total') in df_f.columns else None
        prazo_exec  = df_f[col_mappings['dias_execucao']].mean() if col_mappings.get('dias_execucao') in df_f.columns else None
        pct_sla = None
        if col_mappings.get('dias_total') in df_f.columns and total>0:
            dentro = (df_f[col_mappings['dias_total']] <= sla_dias).sum()
            pct_sla = dentro/total*100

        with c1: st.metric("📊 Total Vistorias", f"{total:,}".replace(",", "."))
        with c2: st.metric("✅ Finalizadas", f"{finalizadas:,} ({pct_final:.1f}%)")
        with c3: st.metric("⏱️ Prazo Médio Total", f"{prazo_total:.1f} dias" if prazo_total is not None else "—")
        with c4: st.metric("🚀 Prazo Médio Exec.", f"{prazo_exec:.1f} dias" if prazo_exec is not None else "—")
        with c5: st.metric(f"🎯 SLA ≤{sla_dias}d", f"{pct_sla:.1f}%" if pct_sla is not None else "—")

        if oms_selecionadas or diretorias_selecionadas or any(filtros.values()) or (periodo and len(periodo)==2):
            applied = []
            if oms_selecionadas: applied.append(f"{len(oms_selecionadas)} OM(s)")
            if diretorias_selecionadas: applied.append(f"{len(diretorias_selecionadas)} Diretoria(s)")
            if any(filtros.values()): applied.extend([f"{len(v)} {k}" for k,v in filtros.items() if v])
            st.info(f"🔍 **Filtros aplicados:** {', '.join(applied)}")

        st.markdown("---")
        st.markdown("### 📊 Análises Gráficas")

        # Evolução
        if col_data_base and col_data_base in df_f.columns and df_f[col_data_base].notna().any():
            monthly = df_f.groupby(pd.Grouper(key=col_data_base, freq="MS")).size().reset_index(name="Vistorias")
            fig1 = px.line(monthly, x=col_data_base, y="Vistorias", markers=True, title="📈 Evolução Mensal de Vistorias", template="plotly_white")
            fig1.update_layout(height=400)
            st.plotly_chart(fig1, use_container_width=True)

        # Diretoria
        if col_dir and col_dir in df_f.columns:
            diretoria_data = (
                df_f[col_dir].dropna().value_counts()
                .reset_index(name="Quantidade")
                .rename(columns={"index": col_dir})
                .sort_values("Quantidade", ascending=True)
            )
            if not diretoria_data.empty:
                fig_dir = px.bar(diretoria_data, x="Quantidade", y=col_dir, orientation="h",
                                 title="🏢 Vistorias por Diretoria Responsável", template="plotly_white",
                                 color="Quantidade", color_continuous_scale="Blues")
                fig_dir.update_layout(height=400, showlegend=False)
                st.plotly_chart(fig_dir, use_container_width=True)

        # Situação
        col_sit = col_mappings.get('situacao')
        if col_sit and col_sit in df_f.columns:
            situacao_data = df_f[col_sit].dropna().value_counts().reset_index(name="Quantidade").rename(columns={"index": col_sit})
            if not situacao_data.empty:
                fig_sit = px.pie(situacao_data, names=col_sit, values="Quantidade",
                                 title="📋 Distribuição por Situação", hole=0.4,
                                 color_discrete_sequence=['#FF6B6B','#4ECDC4','#45B7D1','#96CEB4','#FECA57','#FF9FF3','#54A0FF'])
                fig_sit.update_traces(textposition="inside", textinfo="percent+label",
                                      hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>')
                fig_sit.update_layout(height=400)
                st.plotly_chart(fig_sit, use_container_width=True)

        # Urgência
        col_urg = col_mappings.get('urgencia')
        if col_urg and col_urg in df_f.columns:
            urgencia_data = (
                df_f[col_urg].dropna().value_counts()
                .reset_index(name="Quantidade").rename(columns={"index": col_urg})
                .sort_values("Quantidade", ascending=False)
            )
            if not urgencia_data.empty:
                fig_urg = px.bar(urgencia_data, x=col_urg, y="Quantidade",
                                 title="⚡ Vistorias por Classificação de Urgência", template="plotly_white",
                                 color="Quantidade", color_continuous_scale="Reds")
                fig_urg.update_layout(height=400, showlegend=False)
                fig_urg.update_xaxes(tickangle=45)
                st.plotly_chart(fig_urg, use_container_width=True)

        # Por OM (quando filtradas)
        if oms_selecionadas and col_om and col_om in df_f.columns:
            om_data = (
                df_f[col_om].dropna().value_counts()
                .reset_index(name="Quantidade").rename(columns={"index": col_om})
                .sort_values("Quantidade", ascending=True)
            )
            if not om_data.empty:
                fig_om = px.bar(om_data, x="Quantidade", y=col_om, orientation="h",
                                title=f"🏛️ Vistorias por OM Selecionada ({len(oms_selecionadas)} filtradas)",
                                template="plotly_white", color="Quantidade", color_continuous_scale="Greens")
                fig_om.update_layout(height=400, showlegend=False)
                st.plotly_chart(fig_om, use_container_width=True)

        # Heatmap temporal
        if col_data_base and col_sit and (col_data_base in df_f.columns) and (col_sit in df_f.columns) and df_f[col_data_base].notna().any():
            aux = df_f[[col_data_base, col_sit]].dropna().copy()
            aux["Mes"] = aux[col_data_base].dt.to_period("M").dt.to_timestamp()
            heat = (aux.groupby(["Mes", col_sit]).size().reset_index(name="Qtd")
                    .pivot(index=col_sit, columns="Mes", values="Qtd").fillna(0))
            if not heat.empty:
                fig_hm = px.imshow(heat, aspect="auto", title="🔥 Heatmap - Vistorias por Situação ao Longo do Tempo",
                                   labels=dict(x="Mês", y="Situação", color="Quantidade"),
                                   x=[d.strftime("%Y-%m") for d in heat.columns], y=heat.index,
                                   color_continuous_scale="Viridis")
                fig_hm.update_layout(height=500)
                st.plotly_chart(fig_hm, use_container_width=True)

        st.markdown("### 📋 Dados Detalhados")
        st.dataframe(df_f, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Erro no Dashboard: {e}")
