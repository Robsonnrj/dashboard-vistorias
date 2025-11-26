# core/layout.py
import streamlit as st

def hide_multipage_nav():
    """Esconde o menu de páginas padrão do Streamlit (Home, pages etc)."""
    st.markdown(
        """
        <style>
        /* Esconde apenas o nav de páginas, mas mantém a sidebar
           (para você usar filtros no dashboard, por exemplo) */
        [data-testid="stSidebarNav"] {
            display: none;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
