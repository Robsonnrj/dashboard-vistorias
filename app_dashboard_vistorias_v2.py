# -*- coding: utf-8 -*-
# CRO1 — Editor + Dashboards (Google Sheets) - Versão Otimizada

import warnings
warnings.filterwarnings("ignore", message=".*outside the limits for dates.*", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", message=".*Data Validation extension is not supported and will be removed.*", category=UserWarning, module="openpyxl")

import unicodedata
from datetime import datetime
from pathlib import Path
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
    """Verifica se as configurações do Google Sheets estão disponíveis."""
    return (
        "gcp_service_account" in st.secrets 
        and "gsheets" in st.secrets 
        and "spreadsheet_url" in st.secrets["gsheets"]
        and bool(st.secrets["gsheets"]["spreadsheet_url"])
    )

@st.cache_resource(show_spinner=False)
def get_gs_client():
    """Cliente gspread autenticado via service account do secrets.toml"""
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def get_workbook():
    """Spreadsheet (arquivo) aberto pela URL do secrets.toml"""
    return get_gs_client().open_by_url(st.secrets["gsheets"]["spreadsheet_url"])

def make_unique_headers(raw_headers):
    """Gera nomes únicos: vazio -> col_1; duplicados -> nome_2, nome_3, ..."""
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
    """
    Lê a worksheet tolerando cabeçalho repetido/mesclado/vazio.
    - Se header_row não for dado, usa a primeira linha com algum conteúdo.
    - Garante nomes únicos nas colunas.
    """
    try:
        values = ws.get_all_values()
        if not values:
            return pd.DataFrame()

        # Descobre a linha do cabeçalho
        if header_row is None:
            hdr_idx = next(
                (i for i, row in enumerate(values) if any(str(c).strip() for c in row)), 
                0
            )
        else:
            hdr_idx = max(0, int(header_row) - 1)

        headers = make_unique_headers(values[hdr_idx])
        body = values[hdr_idx + 1:]

        # Remove linhas finais 100% vazias
        while body and not any(str(c).strip() for c in body[-1]):
            body.pop()

        df = pd.DataFrame(body, columns=headers).replace("", pd.NA)
        return df
    
    except Exception as e:
        st.error(f"Erro ao ler worksheet: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=60, show_spinner=False)
def read_tab_df(tab_name: str) -> pd.DataFrame:
    """Lê uma aba do Sheets como DataFrame (infere header da linha 1)."""
    try:
        ws = get_workbook().worksheet(tab_name)
        df = read_worksheet_safe(ws)
        
        # Normaliza datas
        for col in df.columns:
            if "DATA" in col.upper():
                df[col] = pd.to_datetime(df[col], errors="coerce")
        
        return df
    
    except Exception as e:
        st.error(f"Erro ao ler aba {tab_name}: {e}")
        return pd.DataFrame()

def overwrite_tab_from_df(tab_name: str, df: pd.DataFrame, keep_header=True):
    """Sobrescreve a aba com o DataFrame."""
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
        
        # Limpa toda a aba
        ws.clear()
        
        if keep_header:
            values = [list(map(str, df.columns))] + df.fillna("").astype(str).values.tolist()
        else:
            values = df.fillna("").astype(str).values.tolist()
        
        # Atualiza com os novos dados
        ws.update("A1", values, value_input_option="USER_ENTERED")
        
        # Invalida cache de leitura
        read_tab_df.clear()
        return True
        
    except Exception as e:
        st.error(f"Erro ao salvar aba {tab_name}: {e}")
        return False

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================

def normalize_text(s: str) -> str:
    """Normaliza texto removendo acentos e convertendo para minúsculas."""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower().strip()

def find_column(df: pd.DataFrame, options: list[str]) -> str | None:
    """Encontra coluna por nome exato ou contendo o texto."""
    cols = list(df.columns)
    
    # Busca exata primeiro
    for opt in options:
        for col in cols:
            if normalize_text(col) == normalize_text(opt):
                return col
    
    # Busca por conteúdo
    for opt in options:
        target = normalize_text(opt)
        for col in cols:
            if target in normalize_text(col):
                return col
    
    return None

def create_card(title: str):
    """Cria um card estilizado para títulos."""
    st.markdown(
        f"<div style='padding:12px 16px;border-radius:12px;background:linear-gradient(135deg, #667eea 0%, #764ba2 100%);"
        f"border:none;color:white;font-weight:700;font-size:20px;text-align:center;margin:10px 0;'>"
        f"{title}</div>",
        unsafe_allow_html=True
    )

def get_filter_options(series):
    """Obtém opções únicas de uma série para filtros."""
    try:
        return sorted(series.dropna().astype(str).unique().tolist())
    except Exception:
        return sorted(list({str(x) for x in series.dropna().tolist()}))

# =========================================================
# SIDEBAR (STATUS + MENU)
# =========================================================

with st.sidebar:
    st.markdown("### 🔌 Status da Conexão")
    if has_gsheets():
        st.success("Google Sheets: ✅ Conectado")
        
        if st.button("🔄 Limpar cache e recarregar", use_container_width=True):
            read_tab_df.clear()
            st.success("Cache limpo!")
            st.rerun()
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
# 1) EDITOR DE PLANILHA
# =========================================================

if MENU == "🗂️ Editor de Planilha":
    create_card("🗂️ Editor de Planilha Google Sheets")
    
    if not has_gsheets():
        st.error("❌ Google Sheets não configurado. Verifique o arquivo secrets.toml")
        st.stop()

    try:
        workbook = get_workbook()
        tabs = [ws.title for ws in workbook.worksheets()]
        
        st.success("✅ Google Sheets conectado com sucesso!")
        st.caption(f"📋 **Planilha:** {st.secrets['gsheets']['spreadsheet_url']}")
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            tab_name = st.selectbox(
                "📂 Escolha a aba para visualizar/editar:",
                tabs,
                index=0
            )
        
        with col2:
            if st.button("↻ Recarregar", use_container_width=True):
                read_tab_df.clear()
                st.rerun()

        # Carrega dados da aba
        df_tab = read_tab_df(tab_name)
        
        if df_tab.empty:
            st.warning("⚠️ A aba selecionada está vazia.")
        else:
            st.info(f"📊 **Linhas:** {len(df_tab):,} • **Colunas:** {len(df_tab.columns)}")
            
            # Editor interativo
            edited_df = st.data_editor(
                df_tab,
                use_container_width=True,
                num_rows="dynamic",
                key=f"editor_{tab_name}",
                height=500,
                hide_index=True
            )
            
            # Botões de ação
            col1, col2, col3 = st.columns([2, 2, 2])
            
            with col1:
                if st.button("💾 Salvar Alterações", use_container_width=True):
                    with st.spinner("Salvando..."):
                        # Converte datas para formato ISO
                        df_save = edited_df.copy()
                        for col in df_save.columns:
                            if "DATA" in col.upper():
                                df_save[col] = pd.to_datetime(df_save[col], errors="coerce")
                                df_save[col] = df_save[col].dt.strftime("%Y-%m-%d")
                        
                        if overwrite_tab_from_df(tab_name, df_save, keep_header=True):
                            st.success(f"✅ Alterações salvas em **{tab_name}**")
                        else:
                            st.error("❌ Falha ao salvar alterações")
            
            with col2:
                csv = edited_df.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "⬇️ Baixar CSV",
                    csv,
                    file_name=f"{tab_name}_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            with col3:
                excel_buffer = pd.ExcelWriter("temp.xlsx", engine='openpyxl')
                edited_df.to_excel(excel_buffer, sheet_name=tab_name, index=False)
                excel_buffer.close()
                
                st.download_button(
                    "⬇️ Baixar Excel",
                    excel_buffer.getvalue() if hasattr(excel_buffer, 'getvalue') else b'',
                    file_name=f"{tab_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    except Exception as e:
        st.error(f"❌ Erro ao conectar com Google Sheets: {e}")

# =========================================================
# 2) DASHBOARDS
# =========================================================

if MENU == "📊 Dashboards":
    create_card("📊 Dashboards e Análises")
    
    if not has_gsheets():
        st.error("❌ Google Sheets não configurado. Verifique o arquivo secrets.toml")
        st.stop()

    try:
        workbook = get_workbook()
        tabs = [ws.title for ws in workbook.worksheets()]
        
        # Seleção da aba base
        base_tab = st.selectbox(
            "📂 Escolha a aba para análise:",
            tabs,
            index=0,
            key="dashboard_tab",
        )
        
        df = read_tab_df(base_tab)
        
        if df.empty:
            st.warning("⚠️ A aba selecionada está vazia.")
            st.stop()
        
        st.info(f"📊 Analisando **{base_tab}** • **{len(df):,}** registros")

        # Mapeamento de colunas
        col_mappings = {
            'objeto': find_column(df, ["OBJETO DE VISTORIA", "OBJETO"]),
            'om': find_column(df, ["OM APOIADA", "OM APOIADORA", "OM"]),
            'diretoria': find_column(df, ["Diretoria Responsavel", "Diretoria Responsável", "Diretoria"]),
            'urgencia': find_column(df, ["Classificacao da Urgencia", "Classificação da Urgência", "Urgencia"]),
            'situacao': find_column(df, ["Situacao", "Situação"]),
            'data_solicitacao': find_column(df, ["DATA DA SOLICITACAO", "DATA DA SOLICITAÇÃO"]),
            'data_vistoria': find_column(df, ["DATA DA VISTORIA"]),
            'dias_total': find_column(df, ["QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO"]),
            'dias_execucao': find_column(df, ["QUANTIDADE DE DIAS PARA EXECUCAO", "QUANTIDADE DE DIAS PARA EXECUÇÃO"]),
            'status': find_column(df, ["STATUS - ATUALIZACAO SEMANAL", "Status"])
        }

        # Conversão de tipos
        for col in [col_mappings['data_solicitacao'], col_mappings['data_vistoria']]:
            if col and col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")

        for col in [col_mappings['dias_total'], col_mappings['dias_execucao']]:
            if col and col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")

        # Sidebar com filtros
        st.sidebar.markdown("### 🔍 Filtros")
        
        # Filtro de período
        col_data_base = col_mappings['data_solicitacao'] if col_mappings['data_solicitacao'] else col_mappings['data_vistoria']
        periodo = None
        
        if col_data_base and col_data_base in df.columns and df[col_data_base].notna().any():
            min_dt = pd.to_datetime(df[col_data_base].min()).date()
            max_dt = pd.to_datetime(df[col_data_base].max()).date()
            periodo = st.sidebar.date_input(
                "📅 Período",
                value=(min_dt, max_dt),
                min_value=min_dt,
                max_value=max_dt
            )

        # Outros filtros
        filtros = {}
        
        if col_mappings['diretoria'] and col_mappings['diretoria'] in df.columns:
            filtros['diretoria'] = st.sidebar.multiselect(
                "🏢 Diretoria Responsável", 
                get_filter_options(df[col_mappings['diretoria']])
            )

        if col_mappings['situacao'] and col_mappings['situacao'] in df.columns:
            filtros['situacao'] = st.sidebar.multiselect(
                "📋 Situação", 
                get_filter_options(df[col_mappings['situacao']])
            )

        if col_mappings['urgencia'] and col_mappings['urgencia'] in df.columns:
            filtros['urgencia'] = st.sidebar.multiselect(
                "⚡ Urgência", 
                get_filter_options(df[col_mappings['urgencia']])
            )

        if col_mappings['om'] and col_mappings['om'] in df.columns:
            filtros['om'] = st.sidebar.multiselect(
                "🏛️ OM Apoiadora", 
                get_filter_options(df[col_mappings['om']])
            )

        sla_dias = st.sidebar.number_input(
            "⏱️ SLA (dias)", 
            min_value=1, 
            max_value=365, 
            value=30,
            help="Prazo considerado para análise de SLA"
        )

        # Aplicar filtros
        df_filtered = df.copy()

        if periodo and col_data_base:
            ini, fim = periodo
            df_filtered = df_filtered[
                (df_filtered[col_data_base] >= pd.to_datetime(ini)) & 
                (df_filtered[col_data_base] <= pd.to_datetime(fim))
            ]

        # Aplicar filtros das seleções
        filter_mapping = {
            'diretoria': col_mappings['diretoria'],
            'situacao': col_mappings['situacao'],
            'urgencia': col_mappings['urgencia'],
            'om': col_mappings['om']
        }

        for filter_key, col_name in filter_mapping.items():
            if filtros.get(filter_key) and col_name and col_name in df_filtered.columns:
                df_filtered = df_filtered[df_filtered[col_name].astype(str).isin(filtros[filter_key])]

        # KPIs
        st.markdown("### 📈 Indicadores Principais")
        
        col1, col2, col3, col4, col5 = st.columns(5)
        
        total_vistorias = len(df_filtered)
        
        # Finalizadas
        finalizadas = 0
        pct_finalizadas = 0
        if col_mappings['situacao'] and col_mappings['situacao'] in df_filtered.columns:
            finalizadas = df_filtered[col_mappings['situacao']].astype(str).str.upper().str.contains('FINALIZADA', na=False).sum()
            pct_finalizadas = (finalizadas / total_vistorias * 100) if total_vistorias > 0 else 0

        # Prazos médios
        prazo_medio_total = None
        prazo_medio_exec = None
        
        if col_mappings['dias_total'] and col_mappings['dias_total'] in df_filtered.columns:
            prazo_medio_total = df_filtered[col_mappings['dias_total']].mean()
            
        if col_mappings['dias_execucao'] and col_mappings['dias_execucao'] in df_filtered.columns:
            prazo_medio_exec = df_filtered[col_mappings['dias_execucao']].mean()

        # SLA
        pct_sla = None
        if col_mappings['dias_total'] and col_mappings['dias_total'] in df_filtered.columns and total_vistorias > 0:
            dentro_sla = (df_filtered[col_mappings['dias_total']] <= sla_dias).sum()
            pct_sla = dentro_sla / total_vistorias * 100

        with col1:
            st.metric("📊 Total Vistorias", f"{total_vistorias:,}".replace(",", "."))
        
        with col2:
            st.metric("✅ Finalizadas", f"{finalizadas:,} ({pct_finalizadas:.1f}%)")
        
        with col3:
            st.metric("⏱️ Prazo Médio Total", f"{prazo_medio_total:.1f} dias" if prazo_medio_total is not None else "—")
        
        with col4:
            st.metric("🚀 Prazo Médio Exec.", f"{prazo_medio_exec:.1f} dias" if prazo_medio_exec is not None else "—")
        
        with col5:
            st.metric(f"🎯 SLA ≤{sla_dias}d", f"{pct_sla:.1f}%" if pct_sla is not None else "—")

        st.markdown("---")

        # Gráficos
        st.markdown("### 📊 Análises Gráficas")
        
        # Evolução temporal
        if col_data_base and col_data_base in df_filtered.columns and df_filtered[col_data_base].notna().any():
            monthly_data = (
                df_filtered.groupby(pd.Grouper(key=col_data_base, freq='MS'))
                .size()
                .reset_index(name='Vistorias')
            )
            
            fig_evolucao = px.line(
                monthly_data, 
                x=col_data_base, 
                y='Vistorias',
                markers=True,
                title="📈 Evolução Mensal de Vistorias",
                template="plotly_white"
            )
            fig_evolucao.update_layout(height=400)
            st.plotly_chart(fig_evolucao, use_container_width=True)

        # Gráficos por categoria
        chart_configs = [
            (col_mappings['diretoria'], "🏢 Vistorias por Diretoria", "bar"),
            (col_mappings['situacao'], "📋 Distribuição por Situação", "pie"),
            (col_mappings['urgencia'], "⚡ Vistorias por Urgência", "bar"),
        ]

        for col_name, title, chart_type in chart_configs:
            if col_name and col_name in df_filtered.columns:
                chart_data = df_filtered.groupby(col_name, as_index=False).size().sort_values('size', ascending=False)
                
                if chart_type == "bar":
                    fig = px.bar(chart_data, x=col_name, y='size', title=title, template="plotly_white")
                elif chart_type == "pie":
                    fig = px.pie(chart_data, names=col_name, values='size', title=title, hole=0.4)
                
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)

        # Análise de SLA por Diretoria
        if (col_mappings['diretoria'] and col_mappings['diretoria'] in df_filtered.columns and 
            col_mappings['dias_total'] and col_mappings['dias_total'] in df_filtered.columns):
            
            sla_data = df_filtered.dropna(subset=[col_mappings['diretoria'], col_mappings['dias_total']]).copy()
            sla_data['Dentro_SLA'] = sla_data[col_mappings['dias_total']] <= sla_dias
            sla_summary = (
                sla_data.groupby(col_mappings['diretoria'])['Dentro_SLA']
                .mean() * 100
            ).reset_index(name='pct_sla').sort_values('pct_sla')
            
            fig_sla = px.bar(
                sla_summary, 
                x='pct_sla', 
                y=col_mappings['diretoria'],
                orientation='h',
                title=f"🎯 % Dentro do SLA (≤{sla_dias} dias) por Diretoria",
                labels={'pct_sla': '% dentro do SLA'},
                template="plotly_white"
            )
            fig_sla.update_layout(height=400)
            st.plotly_chart(fig_sla, use_container_width=True)

        # Detalhamento dos dados
        st.markdown("### 📋 Detalhamento dos Dados")
        
        # Ordenar por data mais recente se possível
        if col_data_base and col_data_base in df_filtered.columns:
            df_show = df_filtered.sort_values(col_data_base, ascending=False).head(100)
        else:
            df_show = df_filtered.head(100)
        
        st.dataframe(df_show, use_container_width=True, height=400)
        
        # Download dos dados filtrados
        col1, col2 = st.columns(2)
        
        with col1:
            csv_filtered = df_filtered.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Baixar Dados Filtrados (CSV)",
                csv_filtered,
                file_name=f"{base_tab}_filtrado_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"❌ Erro ao carregar dashboards: {e}")

# =========================================================
# RODAPÉ
# =========================================================

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666; font-size: 12px;'>"
    "🚀 CRO1 Sistema - Desenvolvido com Streamlit • "
    f"Última atualização: {datetime.now().strftime('%d/%m/%Y às %H:%M')}"
    "</div>",
    unsafe_allow_html=True
)
