import streamlit as st
from streamlit_option_menu import option_menu
from core.config import has_gsheets, TAB_SOLICITACOES, TAB_VALIDACAO
from core.data_loader import clear_caches

# Configuração da página
st.set_page_config(page_title="CRO1 — Seção de Vistorias", layout="wide")

# Inicializa o estado da sessão para o menu
if "main_menu" not in st.session_state:
    st.session_state["main_menu"] = "📊 Dashboard"  # valor padrão

# Barra lateral fixa
with st.sidebar:
    st.write("Google Sheets:", "ON ✅" if has_gsheets() else "OFF ❌")
    st.markdown("### Abas usadas pelo sistema")
    st.caption(f"Aba de Solicitações / Base: **{TAB_SOLICITACOES}**")
    st.caption(f"Aba de Validação (OMs): **{TAB_VALIDACAO}**")
    if st.button("🔄 Atualizar dados (limpar cache)"):
        clear_caches()
        st.toast("Cache limpo. Recarregando…")
        st.experimental_rerun()

# Menu lateral / topo
MENU = option_menu(
    None,
    ["📊 Dashboard", "📝 Cadastro", "🧾 Status/Auditoria"],
    icons=["bar-chart", "file-plus", "list-check"],
    default_index=["📊 Dashboard", "📝 Cadastro", "🧾 Status/Auditoria"].index(st.session_state["main_menu"]),
    key="main_menu",
)

# Roteamento das páginas por menu
if MENU == "📊 Dashboard":
    from features import dashboard
    dashboard.page()

elif MENU == "📝 Cadastro":
    from features import cadastro

    # Exibe a página de cadastro e permite redirecionar para dashboard
    cadastro.page()

elif MENU == "🧾 Status/Auditoria":
    from features import status
    status.page()
