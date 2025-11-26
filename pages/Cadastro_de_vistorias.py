from __future__ import annotations
import streamlit as st

try:
    from core.config import SHEET_TABS  # type: ignore
except Exception:
    SHEET_TABS = {
        "solicitacoes": "Solicitacoes",
        "historicos":   "Historicos",
        "relatorios":   "Relatorios",
    }

st.session_state.setdefault("tabs_map", dict(SHEET_TABS))

from core.layout import hide_multipage_nav, top_nav
from features.cadastro import page as cadastro_page  # função que desenha o formulário


def page():
    hide_multipage_nav()
    top_nav("Cadastro de vistorias")
    cadastro_page()   # chama a função que desenha o formulário


# 👇 ISSO AQUI É O QUE ESTAVA FALTANDO
if __name__ == "__main__":
    page()
