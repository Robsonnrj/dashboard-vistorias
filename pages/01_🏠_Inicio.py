# 01_🏠_Inicio.py
from __future__ import annotations
import streamlit as st

# 1) Config da página (sidebar recolhida só aqui)
st.set_page_config(page_title="Início", page_icon="🏠", layout="wide", initial_sidebar_state="collapsed")

# 2) CSS básico para deixar só ícones grandões
st.markdown("""
<style>
/* esconder a borda do expander etc. e dar foco só nos ícones */
.block-container { padding-top: 2rem; padding-bottom: 3rem; }
/* Botão "card" redondo grandão */
.icon-btn > button[kind="secondary"] {
  height: 140px; width: 140px;
  border-radius: 28px;
  font-size: 64px; line-height: 1; 
  display: inline-flex; align-items: center; justify-content: center;
}
/* legenda embaixo do ícone */
.icon-caption {
  margin-top: 8px; font-weight: 600; font-size: 0.95rem; text-align: center;
}
/* remove espaço extra entre colunas em telas grandes */
@media (min-width: 992px){
  .row { gap: 28px; }
}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align:center'>Navegação</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center'>Clique em um ícone para abrir a seção</p>", unsafe_allow_html=True)
st.write("")

# 3) Defina aqui seus “ícones do menu” e a página-alvo (pelo nome da página multipage)
ITEMS = [
    # icon, label, TARGET_PAGE (o nome que aparece no menu multipage)
    ("🗂️", "Cadastro de Vistorias", "Cadastro de Vistorias"),
    ("📊", "Dashboard Operacional", "Dashboard Operacional"),
    ("📑", "Relatórios", "Relatórios"),
    ("🔄", "Status / Andamento", "Status / Andamento"),
    ("🕵️", "Auditoria", "Auditoria"),
]

# 4) Helper para desenhar um “card” de ícone
def icon_card(icon: str, label: str, target_page_name: str):
    # botão “vazio” com tooltip = label
    clicked = st.button(icon, key=f"btn_{label}", help=label, type="secondary")
    st.markdown(f"<div class='icon-caption'>{label}</div>", unsafe_allow_html=True)
    if clicked:
        try:
            # navega pelo nome da página (exatamente como aparece no menu do Streamlit)
            st.switch_page(target_page_name)
        except Exception:
            # fallback: se preferir navegar por arquivo, troque pelo caminho relativo:
            # st.switch_page('pages/02_Dashboard_Operacional.py')
            st.toast(f"Não encontrei a página “{target_page_name}”. Confira o nome no menu lateral.", icon="⚠️")

# 5) Grid responsivo (3 por linha em desktop)
cols_per_row = 5 if len(ITEMS) >= 5 else len(ITEMS)
rows = (len(ITEMS) + cols_per_row - 1) // cols_per_row

idx = 0
for _ in range(rows):
    cols = st.columns(cols_per_row, gap="large")
    for c in cols:
        if idx >= len(ITEMS): 
            break
        with c.container():
            st.container().markdown("<div style='text-align:center'>", unsafe_allow_html=True)
            with st.container():
                st.markdown("<div style='display:flex;justify-content:center'>", unsafe_allow_html=True)
                with st.container():
                    st.container().class_name = "icon-btn"  # para o CSS pegar
                # truque: aplicar a classe acima
                st.markdown("<div class='icon-btn'>", unsafe_allow_html=True)
                icon_card(*ITEMS[idx])
                st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        idx += 1

# 6) Opcional: esconder completamente a sidebar nesta página
# (o set_page_config já inicia colapsada; se quiser forçar esconder em telas pequenas, deixe assim mesmo)
