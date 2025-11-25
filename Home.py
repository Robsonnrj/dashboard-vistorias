from __future__ import annotations
from string import Template
import streamlit as st

# =========================================================
# Import / Fallback de config
# =========================================================
try:
    from core.config import SHEET_TABS  # type: ignore
except Exception:
    SHEET_TABS = {
        "solicitacoes": "Solicitacoes",
        "historicos":   "Historicos",
        "relatorios":   "Relatorios",
    }

# =========================================================
# Configuração da Página
# =========================================================
st.set_page_config(
    page_title="CRO/1 — Sistema de Vistorias",
    page_icon="https://raw.githubusercontent.com/Robsonnrj/dashboard-vistorias/main/house_3661264.png",
    layout="wide",
    initial_sidebar_state="collapsed",
)
st.session_state.setdefault("tabs_map", dict(SHEET_TABS))

# =========================================================
# Router — detecta ?nav= e faz switch_page
# =========================================================
def _nav_param():
    try:
        val = st.query_params.get("nav", None)
        if isinstance(val, list):
            val = val[0]
        return val
    except Exception:
        return None

_nav = _nav_param()
if _nav:
    try:
        st.query_params.clear()
    except Exception:
        pass
    st.switch_page(_nav)

# =========================================================
# Paleta de Cores — Azul Institucional
# =========================================================
PRIMARY_NAVY   = "#0F2A3A"
ACCENT_BLUE    = "#1E40AF"
ACCENT_BLUE_LT = "#E8F0FF"
TEXT_MUTED     = "#64748B"
CARD_BORDER    = "#E8EEF5"
BG             = "#F7F9FC"

# =========================================================
# CSS Global
# =========================================================
PALETTE = {
    "BG": BG,
    "PRIMARY_NAVY": PRIMARY_NAVY,
    "TEXT_MUTED": TEXT_MUTED,
    "CARD_BORDER": CARD_BORDER,
    "ACCENT_BLUE": ACCENT_BLUE,
    "ACCENT_BLUE_LT": ACCENT_BLUE_LT,
}

_css_tpl = Template("""
<style>
[data-testid="stAppViewContainer"] {
  background: $BG;
}
.block-container {padding-top:2.2rem; padding-bottom:3rem;}

/* Cabeçalho */
h1.page-title { color:$PRIMARY_NAVY; margin:0; }
.page-sub { color:$TEXT_MUTED; font-size:18px; margin-top:6px; }
.header-strip {
  position:absolute; right:60px; top:86px;
  background:$PRIMARY_NAVY; color:#E6F2EA;
  padding:10px 20px; border-radius:12px; font-size:14px; letter-spacing:.2px;
}
.hr { height:2px; background:#E6ECF2; margin:20px 0 10px; border-radius:2px; }

/* Grid e Cards */
#cards-grid {
  display:grid; grid-template-columns:repeat(5, minmax(220px,1fr));
  gap:28px; margin-top:22px;
}
#cards-grid .st-card{
  display:block; text-decoration:none; color:$PRIMARY_NAVY;
  background:#fff; border:1.5px solid $CARD_BORDER; border-radius:20px;
  padding:26px 22px 74px; box-shadow:0 10px 30px rgba(15,42,58,.08);
  transition:transform .15s ease, box-shadow .2s ease, border-color .2s ease;
  text-align:center;
}
#cards-grid .st-card:hover{
  transform:translateY(-2px); border-color:#D7E3F5;
  box-shadow:0 16px 36px rgba(30,64,175,.16);
}
.card-ico{
  display:inline-grid; place-items:center;
  width:68px; height:68px; border-radius:50%;
  background:$ACCENT_BLUE; color:#fff; font-size:30px;
  box-shadow:0 6px 16px rgba(30,64,175,.30);
  margin:2px auto 10px auto;
}
.card-title{ font-size:20px; font-weight:700; margin:8px 0 6px 0; }
.card-hint{ font-size:14px; color:$TEXT_MUTED; min-height:42px; }
.card-chip{
  display:inline-block; margin-top:14px; padding:6px 18px; border-radius:20px;
  background:$ACCENT_BLUE_LT; color:$ACCENT_BLUE;
  border:1px solid #C9DAFF; font-size:14px; font-weight:600; letter-spacing:.2px;
}

/* Responsivo */
@media (max-width:1400px){ #cards-grid{grid-template-columns:repeat(4,1fr);} }
@media (max-width:1100px){ #cards-grid{grid-template-columns:repeat(3,1fr);} }
@media (max-width:800px){  #cards-grid{grid-template-columns:repeat(2,1fr);} }
@media (max-width:520px){  #cards-grid{grid-template-columns:1fr;} }
</style>
""")
st.markdown(_css_tpl.substitute(PALETTE), unsafe_allow_html=True)

# =========================================================
# Cabeçalho
# =========================================================
st.markdown("""
<div style="position:relative;">
  <div>
    <h1 class="page-title">Navegação</h1>
    <div class="page-sub">Clique em um ícone para abrir a seção</div>
  </div>
  <div class="header-strip">CRO/1 — Vistorias Técnicas</div>
</div>
<div class="hr"></div>
""", unsafe_allow_html=True)

# =========================================================
# Lista de Cards
# =========================================================
# =========================================================
# Lista de Cards
# =========================================================
cards = [
    {"k":"1","title":"Cadastro de Vistorias","icon":"🗂️","hint":"Criar, editar e validar registros","path":"pages/Cadastro_de_vistorias.py"},
    {"k":"2","title":"Dashboard Operacional","icon":"📊","hint":"KPIs, prazos e mapa de calor","path":"pages/Dashboard_operacional.py"},
    {"k":"3","title":"Relatórios","icon":"📑","hint":"Emitir PDFs e planilhas","path":"pages/Relatorios.py"},
    {"k":"4","title":"Status / Andamento","icon":"🔄","hint":"Triagem e acompanhamento","path":"pages/Status_Andamento.py"},
]

# =========================================================
# Cards estilizados lado a lado (5 por linha)
# =========================================================
st.markdown(
    """
    <style>
    .card-box {
        background: white;
        border: 1.5px solid #E8EEF5;
        border-radius: 20px;
        padding: 26px 22px 28px 22px;
        box-shadow: 0 10px 30px rgba(15,42,58,0.08);
        text-align: center;
        transition: all .2s ease;
    }
    .card-box:hover {
        transform: translateY(-3px);
        border-color: #D7E3F5;
        box-shadow: 0 16px 36px rgba(30,64,175,.16);
    }
    .card-ico {
        display: inline-grid;
        place-items: center;
        width: 68px;
        height: 68px;
        border-radius: 50%;
        background: #1E40AF;
        color: white;
        font-size: 30px;
        box-shadow: 0 6px 16px rgba(30,64,175,.30);
        margin: 0 auto 10px auto;
    }
    .card-title { color: #0F2A3A; font-size: 20px; margin: 6px 0 8px 0; font-weight: 700; }
    .card-hint { color: #64748B; font-size: 14px; min-height: 42px; margin-bottom: 10px; }
    .stButton>button {
        border: 1px solid #C9DAFF;
        background: #E8F0FF;
        color: #1E40AF;
        font-weight: 600;
        border-radius: 20px;
        padding: 6px 18px;
        transition: all .2s ease;
    }
    .stButton>button:hover {
        background: #D7E3F5;
        color: #0F2A3A;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Define grid responsivo (5 colunas)
cols_per_row = 5
clicked_path = None

for i in range(0, len(cards), cols_per_row):
    cols = st.columns(cols_per_row)
    for col, c in zip(cols, cards[i:i + cols_per_row]):
        with col:
            st.markdown(
                f"""
                <div class="card-box">
                    <div class="card-ico">{c['icon']}</div>
                    <div class="card-title">{c['title']}</div>
                    <div class="card-hint">{c['hint']}</div>
                </div>
                """,
                unsafe_allow_html=True
            )
            if st.button("Abrir", key=f"card-{c['k']}", use_container_width=True):
                clicked_path = c["path"]

# =========================================================
# Redirecionamento se clicou
# =========================================================
if clicked_path:
    st.switch_page(clicked_path)

st.caption("Dica: use as teclas **1–5** para abrir as seções rapidamente.")
