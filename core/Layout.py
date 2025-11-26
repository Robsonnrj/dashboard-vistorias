# core/Layout.py
import streamlit as st

def hide_multipage_nav():
    """
    Esconde o menu padrão de páginas do Streamlit (Home, etc.)
    que aparece na sidebar.
    """
    st.markdown(
        """
        <style>
        /* Esconde apenas o nav de páginas padrão */
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
    `active` = rótulo da página atual (só para destacar).
    Usa st.page_link em vez de st.switch_page (mais estável).
    """

    pages = [
        ("Home",                  "Home.py",                          "🏠"),
        ("Cadastro de vistorias", "pages/Cadastro_de_vistorias.py",   "🗂️"),
        ("Dashboard operacional", "pages/Dashboard_operacional.py",   "📊"),
        ("Relatorios",            "pages/Relatorios.py",              "📑"),
        ("Status / Andamento",    "pages/Status_Andamento.py",        "🔄"),
    ]

    st.markdown(
        '<div style="margin-bottom:0.5rem;"></div>',
        unsafe_allow_html=True,
    )

    cols = st.columns(len(pages))

    for (label, target, icon), col in zip(pages, cols):
        with col:
            if label == active:
                # Página atual: botão “marcado” e desabilitado
                st.button(
                    f"{icon} {label}",
                    disabled=True,
                    use_container_width=True,
                )
            else:
                # Outras páginas: link de navegação
                st.page_link(
                    target,
                    label=f"{icon} {label}",
                )
