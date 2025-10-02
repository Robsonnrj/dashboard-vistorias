# ... imports no topo
import streamlit as st
from streamlit_option_menu import option_menu
from core.config import has_gsheets, TAB_SOLICITACOES, TAB_VALIDACAO
from core.data_loader import clear_caches

st.set_page_config(page_title="CRO1 — Seção de Vistorias", layout="wide")

with st.sidebar:
    st.write("Google Sheets:", "ON ✅" if has_gsheets() else "OFF ❌")
    st.markdown("### Abas usadas pelo sistema")
    st.caption(f"Aba de Solicitações / Base: **{TAB_SOLICITACOES}**")
    st.caption(f"Aba de Validação (OMs): **{TAB_VALIDACAO}**")
    if st.button("🔄 Atualizar dados (limpar cache)"):
        clear_caches()
        st.toast("Cache limpo. Recarregando…")
        st.rerun()  # <- troquei experimental_rerun() por rerun()
   
    None,
    ["📊 Dashboard", "📝 Cadastro", "🧾 Status/Auditoria"],
    icons=["bar-chart", "file-plus", "list-check"],
    default_index=0,
)

if MENU == "📊 Dashboard":
    from features import dashboard
    dashboard.page()
elif MENU == "📝 Cadastro":
    from features import cadastro
    cadastro.page()
else:
    from features import status
    status.page()
