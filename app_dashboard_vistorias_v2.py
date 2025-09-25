# CRO1 Dashboard - Sistema Transformado v2.0
# main.py - Aplicação Principal

import warnings
warnings.filterwarnings("ignore", message=".*outside the limits for dates.*", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", message=".*Data Validation extension is not supported.*", category=UserWarning, module="openpyxl")

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from streamlit_option_menu import option_menu
from datetime import datetime
import sys
import os
from pathlib import Path

# Adiciona o diretório atual ao path para imports locais
current_dir = Path(__file__).parent
sys.path.append(str(current_dir))

# Configuração da página
st.set_page_config(
    page_title="CRO1 Dashboard v2.0",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado
def load_custom_css():
    """Carrega CSS personalizado para o tema"""
    css = """
    <style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        margin-bottom: 1rem;
    }
    
    .filter-section {
        background: #f8f9fa;
        padding: 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        border: 1px solid #dee2e6;
    }
    
    .status-badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 0.5rem;
        font-size: 0.85rem;
        font-weight: 600;
        margin: 0.25rem;
    }
    
    .status-success { 
        background-color: #4ECDC4; 
        color: white; 
    }
    .status-warning { 
        background-color: #FECA57; 
        color: black; 
    }
    .status-danger { 
        background-color: #FF6B6B; 
        color: white; 
    }
    .status-info {
        background-color: #667eea;
        color: white;
    }
    
    .stButton > button {
        border-radius: 8px;
        border: none;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

# Função para verificar conexão Google Sheets
def check_gsheets_connection():
    """Verifica se as configurações do Google Sheets estão disponíveis"""
    try:
        return (
            "gcp_service_account" in st.secrets
            and "gsheets" in st.secrets
            and "spreadsheet_url" in st.secrets["gsheets"]
            and bool(st.secrets["gsheets"]["spreadsheet_url"])
        )
    except Exception:
        return False

# Função para criar card estilizado
def create_card(title: str, content: str = ""):
    """Cria um card estilizado"""
    st.markdown(
        f"""
        <div class='main-header'>
            <h1 style='margin: 0; font-size: 2.5rem;'>🚀 {title}</h1>
            <p style='margin: 0.5rem 0 0 0; opacity: 0.9; font-size: 1.1rem;'>{content}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

def main():
    """Função principal da aplicação"""
    
    # Carrega CSS personalizado
    load_custom_css()
    
    # Header principal
    create_card(
        "CRO1 Dashboard v2.0", 
        "Sistema Modular de Análise de Vistorias - Transformação Completa"
    )
    
    # Sidebar com status e menu
    render_sidebar()
    
    # Menu principal
    menu_selection = render_main_menu()
    
    # Roteamento baseado na seleção do menu
    route_handler(menu_selection)

def render_sidebar():
    """Renderiza a sidebar com status e controles"""
    
    with st.sidebar:
        st.markdown("### 🔌 Status do Sistema")
        
        # Status da conexão
        if check_gsheets_connection():
            st.markdown(
                '<div class="status-badge status-success">✅ Google Sheets Conectado</div>',
                unsafe_allow_html=True
            )
            
            # Botões de controle
            st.markdown("### ⚡ Controles")
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("🔄 Cache", use_container_width=True, help="Limpar cache do sistema"):
                    clear_cache()
                    st.success("Cache limpo!")
                    st.rerun()
            
            with col2:
                if st.button("📊 Status", use_container_width=True, help="Ver estatísticas"):
                    show_stats()
                    
        else:
            st.markdown(
                '<div class="status-badge status-danger">❌ Google Sheets Desconectado</div>',
                unsafe_allow_html=True
            )
            st.warning("Configure o arquivo `.streamlit/secrets.toml`")
            
            # Instruções de configuração
            with st.expander("📋 Como Configurar"):
                st.markdown("""
                **1. Crie o arquivo `.streamlit/secrets.toml`:**
                ```
                [gcp_service_account]
                type = "service_account"
                project_id = "your-project-id"
                private_key_id = "your-private-key-id"
                private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
                client_email = "your-service-account@project.iam.gserviceaccount.com"
                client_id = "your-client-id"
                auth_uri = "https://accounts.google.com/o/oauth2/auth"
                token_uri = "https://oauth2.googleapis.com/token"

                [gsheets]
                spreadsheet_url = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID"
                ```
                
                **2. Configure as permissões do Google Sheets**
                
                **3. Reinicie a aplicação**
                """)

def render_main_menu():
    """Renderiza o menu principal"""
    
    with st.sidebar:
        st.markdown("---")
        
        menu = option_menu(
            "🚀 Menu Principal",
            [
                "📊 Dashboard Principal", 
                "📋 Editor de Dados",
                "⚡ Análise de Urgências", 
                "📈 Tendências e Métricas",
                "🔍 Relatórios Avançados",
                "⚙️ Configurações"
            ],
            icons=[
                "bar-chart-fill", "table", "lightning-fill", 
                "graph-up", "file-earmark-text", "gear-fill"
            ],
            default_index=0,
            menu_icon="grid-3x3-gap-fill",
            styles={
                "container": {"padding": "5px", "background-color": "#fafafa"},
                "icon": {"color": "#667eea", "font-size": "18px"},
                "nav-link": {
                    "font-size": "16px", 
                    "text-align": "left", 
                    "margin": "0px",
                    "padding": "10px 15px"
                },
                "nav-link-selected": {
                    "background-color": "#667eea",
                    "color": "white"
                },
            }
        )
    
    return menu

def clear_cache():
    """Limpa todos os caches"""
    try:
        st.cache_data.clear()
        st.cache_resource.clear()
    except Exception as e:
        st.error(f"Erro ao limpar cache: {e}")

def show_stats():
    """Mostra estatísticas do sistema"""
    with st.sidebar.expander("📊 Estatísticas do Sistema"):
        st.write("**Cache Status**: Ativo")
        st.write("**Sessão**: Ativa")
        st.write("**Última atualização**: ", datetime.now().strftime("%H:%M:%S"))

def route_handler(menu_selection):
    """Gerencia o roteamento entre as diferentes páginas"""
    
    if menu_selection == "📊 Dashboard Principal":
        render_dashboard_page()
        
    elif menu_selection == "📋 Editor de Dados":
        render_editor_page()
        
    elif menu_selection == "⚡ Análise de Urgências":
        render_urgency_page()
        
    elif menu_selection == "📈 Tendências e Métricas":
        render_trends_page()
        
    elif menu_selection == "🔍 Relatórios Avançados":
        render_reports_page()
        
    elif menu_selection == "⚙️ Configurações":
        render_settings_page()

# Importa as funções das páginas (colocar os arquivos na pasta pages/)
try:
    from pages.dashboard import render_dashboard_page
    from pages.editor import render_editor_page
    from pages.urgency import render_urgency_page
    from pages.trends import render_trends_page
    from pages.reports import render_reports_page
    from pages.settings import render_settings_page
except ImportError:
    # Fallback se as páginas não existirem
    def render_dashboard_page():
        st.error("❌ Arquivo pages/dashboard.py não encontrado")
    def render_editor_page():
        st.error("❌ Arquivo pages/editor.py não encontrado")
    def render_urgency_page():
        st.error("❌ Arquivo pages/urgency.py não encontrado")
    def render_trends_page():
        st.error("❌ Arquivo pages/trends.py não encontrado")
    def render_reports_page():
        st.error("❌ Arquivo pages/reports.py não encontrado")
    def render_settings_page():
        st.error("❌ Arquivo pages/settings.py não encontrado")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"❌ Erro na aplicação: {str(e)}")
        st.exception(e)
    
    # Rodapé
    st.markdown("---")
    st.markdown(
        f"""
        <div style='text-align: center; color: #666; font-size: 14px; padding: 1rem;'>
            🚀 <strong>CRO1 Dashboard v2.0</strong> - Sistema Modular de Gestão de Vistorias<br>
            <small>Transformação Completa • Última atualização: {datetime.now().strftime('%d/%m/%Y às %H:%M')}</small>
        </div>
        """,
        unsafe_allow_html=True
    )
# pages/dashboard.py
# Dashboard Principal - CRO1 v2.0

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import numpy as np
import unicodedata
import io

# Configurações do Google Sheets
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

@st.cache_resource
def get_gs_client():
    """Cliente gspread autenticado"""
    try:
        info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Erro na autenticação: {e}")
        return None

@st.cache_resource  
def get_workbook():
    """Workbook do Google Sheets"""
    try:
        client = get_gs_client()
        if client:
            return client.open_by_url(st.secrets["gsheets"]["spreadsheet_url"])
    except Exception as e:
        st.error(f"Erro ao abrir planilha: {e}")
        return None

def make_unique_headers(raw_headers):
    """Gera nomes únicos para cabeçalhos"""
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

@st.cache_data(ttl=300)
def load_worksheet_data(tab_name: str) -> pd.DataFrame:
    """Carrega dados de uma worksheet"""
    try:
        workbook = get_workbook()
        if not workbook:
            return pd.DataFrame()
            
        worksheet = workbook.worksheet(tab_name)
        values = worksheet.get_all_values()
        
        if not values:
            return pd.DataFrame()

        # Identifica cabeçalho
        header_row = 0
        for i, row in enumerate(values):
            if any(str(c).strip() for c in row):
                header_row = i
                break

        headers = make_unique_headers(values[header_row])
        data_rows = values[header_row + 1:]

        # Remove linhas vazias
        while data_rows and not any(str(c).strip() for c in data_rows[-1]):
            data_rows.pop()

        if not data_rows:
            return pd.DataFrame(columns=headers)

        df = pd.DataFrame(data_rows, columns=headers).replace("", pd.NA)
        
        # Converte tipos
        for col in df.columns:
            if "DATA" in col.upper():
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
            elif any(keyword in col.upper() for keyword in ["DIAS", "QUANTIDADE", "NUMERO"]):
                df[col] = pd.to_numeric(df[col], errors="coerce")
        
        return df
        
    except Exception as e:
        st.error(f"Erro ao carregar {tab_name}: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=900)
def load_oms_validation_data() -> pd.DataFrame:
    """Carrega dados de OMs da aba de validação"""
    try:
        df_validation = load_worksheet_data("Validacao_de_Dados")
        
        if df_validation.empty:
            return pd.DataFrame()
        
        # Mapeia colunas automaticamente
        column_mapping = {}
        for col in df_validation.columns:
            col_upper = col.upper()
            if 'OM' in col_upper and not column_mapping.get('sigla'):
                column_mapping['sigla'] = col
            elif 'ORGANIZACAO MILITAR' in col_upper or 'ORGANIZAÇÃO MILITAR' in col_upper:
                column_mapping['nome_completo'] = col
            elif 'DIRETORIA' in col_upper:
                column_mapping['diretoria'] = col

        # Se não mapeou automaticamente, usa posição
        cols = list(df_validation.columns)
        if not column_mapping.get('sigla') and len(cols) > 1:
            column_mapping['sigla'] = cols[1]
        if not column_mapping.get('nome_completo') and len(cols) > 2:
            column_mapping['nome_completo'] = cols[2]
        if not column_mapping.get('diretoria') and len(cols) > 3:
            column_mapping['diretoria'] = cols[3]

        # Processa dados
        oms_data = []
        for _, row in df_validation.iterrows():
            sigla = str(row.get(column_mapping.get('sigla', ''), '')).strip()
            nome_completo = str(row.get(column_mapping.get('nome_completo', ''), '')).strip()
            diretoria = str(row.get(column_mapping.get('diretoria', ''), '')).strip()
            
            if not sigla or sigla == 'nan' or len(sigla) < 2:
                continue
                
            om_entry = {
                'sigla': sigla,
                'nome_completo': nome_completo if nome_completo != 'nan' else sigla,
                'diretoria': diretoria if diretoria != 'nan' else 'Não Especificada',
                'display_name': f"{sigla} - {nome_completo}" if nome_completo != 'nan' and nome_completo != sigla else sigla,
                'search_text': f"{sigla} {nome_completo}".upper()
            }
            oms_data.append(om_entry)

        if not oms_data:
            return pd.DataFrame()

        df_oms = pd.DataFrame(oms_data)
        df_oms = df_oms.drop_duplicates(subset=['sigla'], keep='first')
        df_oms = df_oms.sort_values(['diretoria', 'sigla'])
        
        return df_oms
        
    except Exception as e:
        st.error(f"Erro ao carregar dados de OMs: {e}")
        return pd.DataFrame()

def find_column(df: pd.DataFrame, patterns: list) -> str:
    """Encontra coluna por padrões"""
    for pattern in patterns:
        for col in df.columns:
            if pattern.upper() in col.upper():
                return col
    return None

def create_om_filter_component(df_oms: pd.DataFrame, key_suffix: str = ""):
    """Cria filtro hierárquico de OMs"""
    
    if df_oms.empty:
        st.sidebar.warning("⚠️ Lista de OMs não disponível")
        return [], []

    # Filtro de Diretoria
    diretorias_disponiveis = ['Todas'] + sorted(df_oms['diretoria'].unique().tolist())
    diretoria_selecionada = st.sidebar.selectbox(
        "🏢 Diretoria Responsável",
        diretorias_disponiveis,
        key=f"dir_filter_{key_suffix}"
    )
    
    # Filtra OMs por diretoria
    if diretoria_selecionada == 'Todas':
        oms_filtradas = df_oms
    else:
        oms_filtradas = df_oms[df_oms['diretoria'] == diretoria_selecionada]
    
    # Campo de busca
    search_term = st.sidebar.text_input(
        "🔍 Buscar OM",
        key=f"om_search_{key_suffix}",
        placeholder="Digite sigla ou nome..."
    )
    
    # Filtra por busca
    if search_term:
        search_upper = search_term.upper()
        mask = oms_filtradas['search_text'].str.contains(search_upper, na=False, regex=False)
        oms_para_selecao = oms_filtradas[mask]
    else:
        oms_para_selecao = oms_filtradas
    
    # Multiselect
    opcoes_om = oms_para_selecao['display_name'].tolist()
    
    if opcoes_om:
        oms_selecionadas = st.sidebar.multiselect(
            f"🏛️ OM Apoiadora ({len(opcoes_om)} encontradas)",
            opcoes_om,
            key=f"om_multi_{key_suffix}",
            help=f"Filtradas da {diretoria_selecionada}"
        )
        
        # Converte para siglas
        siglas_selecionadas = []
        if oms_selecionadas:
            for om_display in oms_selecionadas:
                sigla = oms_para_selecao[oms_para_selecao['display_name'] == om_display]['sigla'].iloc[0]
                siglas_selecionadas.append(sigla)
        
        return siglas_selecionadas, [diretoria_selecionada] if diretoria_selecionada != 'Todas' else []
    else:
        st.sidebar.info("ℹ️ Nenhuma OM encontrada")
        return [], [diretoria_selecionada] if diretoria_selecionada != 'Todas' else []

def render_dashboard_page():
    """Renderiza página principal do dashboard"""
    
    st.markdown("## 📊 Dashboard Principal")
    st.markdown("Análise completa de vistorias com KPIs em tempo real e filtros avançados")
    
    # Verifica conexão
    if not get_workbook():
        st.error("❌ Erro de conexão com Google Sheets. Verifique a configuração.")
        return
    
    # Obtém lista de abas
    try:
        workbook = get_workbook()
        tabs = [ws.title for ws in workbook.worksheets()]
        
        if not tabs:
            st.warning("Nenhuma aba encontrada na planilha")
            return
            
    except Exception as e:
        st.error(f"Erro ao acessar planilha: {e}")
        return

    # Seleção da aba
    col1, col2 = st.columns([3, 1])
    with col1:
        base_tab = st.selectbox(
            "📂 Selecione a aba para análise:",
            tabs,
            key="dashboard_tab"
        )
    
    with col2:
        if st.button("🔄 Atualizar", use_container_width=True):
            load_worksheet_data.clear()
            load_oms_validation_data.clear()
            st.rerun()

    # Carrega dados principais
    df = load_worksheet_data(base_tab)
    
    if df.empty:
        st.warning("⚠️ Nenhum dado encontrado na aba selecionada")
        return

    st.success(f"✅ {len(df):,} registros carregados de **{base_tab}**")

    # Mapeamento de colunas
    column_mappings = {
        'objeto': find_column(df, ["OBJETO DE VISTORIA", "OBJETO"]),
        'om': find_column(df, ["OM APOIADA", "OM APOIADORA", "OM"]),
        'diretoria': find_column(df, ["DIRETORIA RESPONSAVEL", "DIRETORIA"]),
        'urgencia': find_column(df, ["CLASSIFICACAO DA URGENCIA", "URGENCIA"]),
        'situacao': find_column(df, ["SITUACAO"]),
        'data_solicitacao': find_column(df, ["DATA DA SOLICITACAO"]),
        'data_vistoria': find_column(df, ["DATA DA VISTORIA"]),
        'dias_total': find_column(df, ["QUANTIDADE DE DIAS PARA TOTAL"]),
        'dias_execucao': find_column(df, ["QUANTIDADE DE DIAS PARA EXECUCAO"]),
        'status': find_column(df, ["STATUS", "VT EXECUTADA POR"])
    }

    # Sidebar com filtros
    st.sidebar.markdown("### 🔍 Filtros Avançados")
    
    # Carrega dados de OMs
    df_oms = load_oms_validation_data()
    
    # Filtros hierárquicos de OMs
    oms_selecionadas, diretorias_selecionadas = create_om_filter_component(df_oms, "dashboard")

    # SLA
    sla_dias = st.sidebar.number_input(
        "⏱️ SLA (dias)",
        min_value=1,
        max_value=365,
        value=30
    )

    # Aplica filtros
    df_filtered = df.copy()

    # Filtro de OMs
    if oms_selecionadas and column_mappings.get('om'):
        col_om = column_mappings['om']
        pattern = '|'.join([om.upper() for om in oms_selecionadas])
        mask = df_filtered[col_om].astype(str).str.upper().str.contains(pattern, na=False, regex=True)
        df_filtered = df_filtered[mask]

    # KPIs principais
    st.markdown("### 📈 Indicadores Principais")
    
    total_vistorias = len(df_filtered)
    
    # Calcula métricas
    finalizadas = 0
    pct_finalizadas = 0
    if column_mappings.get('situacao'):
        col_sit = column_mappings['situacao']
        finalizadas = df_filtered[col_sit].astype(str).str.upper().str.contains('FINALIZADA', na=False).sum()
        pct_finalizadas = (finalizadas / total_vistorias * 100) if total_vistorias > 0 else 0

    prazo_medio_total = None
    if column_mappings.get('dias_total'):
        prazo_medio_total = pd.to_numeric(df_filtered[column_mappings['dias_total']], errors='coerce').mean()

    pct_sla = None
    if column_mappings.get('dias_total') and total_vistorias > 0:
        dias_numeric = pd.to_numeric(df_filtered[column_mappings['dias_total']], errors='coerce')
        dentro_sla = (dias_numeric <= sla_dias).sum()
        pct_sla = dentro_sla / total_vistorias * 100

    # Renderiza KPIs
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("📊 Total", f"{total_vistorias:,}".replace(",", "."))
    
    with col2:
        st.metric("✅ Finalizadas", f"{finalizadas:,} ({pct_finalizadas:.1f}%)")
    
    with col3:
        if prazo_medio_total is not None:
            st.metric("⏱️ Prazo Médio", f"{prazo_medio_total:.1f} dias")
        else:
            st.metric("⏱️ Prazo Médio", "—")
    
    with col4:
        if pct_sla is not None:
            st.metric(f"🎯 SLA ≤{sla_dias}d", f"{pct_sla:.1f}%")
        else:
            st.metric(f"🎯 SLA ≤{sla_dias}d", "—")

    # Gráficos
    st.markdown("### 📊 Análises Gráficas")

    # Evolução temporal
    if column_mappings.get('data_solicitacao') and column_mappings['data_solicitacao'] in df_filtered.columns:
        monthly_data = (
            df_filtered.groupby(pd.Grouper(key=column_mappings['data_solicitacao'], freq='MS'))
            .size()
            .reset_index(name='Vistorias')
        )
        
        if not monthly_data.empty:
            fig_evolucao = px.line(
                monthly_data,
                x=column_mappings['data_solicitacao'],
                y='Vistorias',
                markers=True,
                title="📈 Evolução Mensal de Vistorias",
                color_discrete_sequence=["#667eea"]
            )
            fig_evolucao.update_layout(height=400, template="plotly_white")
            st.plotly_chart(fig_evolucao, use_container_width=True)

    # Gráficos lado a lado
    col1, col2 = st.columns(2)
    
    with col1:
        # Distribuição por situação
        if column_mappings.get('situacao'):
            situacao_data = df_filtered[column_mappings['situacao']].value_counts().head(8)
            
            fig_sit = px.pie(
                values=situacao_data.values,
                names=situacao_data.index,
                title="📋 Distribuição por Situação",
                hole=0.4,
                color_discrete_sequence=["#667eea", "#4ECDC4", "#FECA57", "#FF6B6B"]
            )
            fig_sit.update_layout(height=400)
            st.plotly_chart(fig_sit, use_container_width=True)

    with col2:
        # Vistorias por diretoria
        if column_mappings.get('diretoria'):
            dir_data = df_filtered[column_mappings['diretoria']].value_counts().head(10)
            
            fig_dir = px.bar(
                x=dir_data.values,
                y=dir_data.index,
                orientation='h',
                title="🏢 Vistorias por Diretoria",
                color=dir_data.values,
                color_continuous_scale="Blues"
            )
            fig_dir.update_layout(height=400, showlegend=False, template="plotly_white")
            st.plotly_chart(fig_dir, use_container_width=True)

    # Detalhamento dos dados
    st.markdown("### 📋 Detalhamento dos Dados")
    
    # Tabela de dados
    df_show = df_filtered.head(100)
    st.dataframe(df_show, use_container_width=True, height=400, hide_index=True)

    # Downloads
    col1, col2 = st.columns(2)
    
    with col1:
        csv_data = df_filtered.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Baixar CSV Filtrado",
            csv_data,
            file_name=f"cro1_dashboard_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with col2:
        # Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_filtered.to_excel(writer, index=False, sheet_name="CRO1_Dados")
        
        st.download_button(
            "⬇️ Baixar Excel Filtrado", 
            output.getvalue(),
            file_name=f"cro1_dashboard_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
