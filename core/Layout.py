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

def top_nav(active: str):
    """
    Menu superior para navegar entre as páginas.
    `active` = nome da página atual (só para estilizar o botão).
    """
    pages = [
        ("Home",                  "Home"),
        ("Cadastro de vistorias", "pages/Cadastro_de_vistorias.py"),
        ("Dashboard operacional", "pages/Dashboard_operacional.py"),
        ("Relatórios",            "pages/Relatorios.py"),
        ("Status / Andamento",    "pages/Status_Andamento.py"),
    ]

    cols = st.columns(len(pages))
    for (label, target), col in zip(pages, cols):
        with col:
            is_active = (label == active)
            btn_label = f"• {label}" if is_active else label
            if st.button(
                btn_label,
                key=f"nav-{label}",
                type="primary" if is_active else "secondary",
                use_container_width=True,
            ):
                # troca de página
                st.switch_page(target)
