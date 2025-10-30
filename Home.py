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

# =========================================================
# Configuração da Página
# =========================================================
st.set_page_config(
    page_title="CRO/1 — Sistema de Vistorias",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="collapsed",
)
st.session_state.setdefault("tabs_map", dict(SHEET_TABS))

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
# CSS (sem f-string — usa %)
# =========================================================
PALETTE = {
    "BG": BG,
    "PRIMARY_NAVY": PRIMARY_NAVY,
    "TEXT_MUTED": TEXT_MUTED,
    "CARD_BORDER": CARD_BORDER,
    "ACCENT_BLUE": ACCENT_BLUE,
    "ACCENT_BLUE_LT": ACCENT_BLUE_LT,
}

st.markdown("""
<style>
/* Fundo */
[data-testid="stAppViewContainer"] {
  background: %(BG)s;
}
.block-container {padding-top:2.2rem; padding-bottom:3rem;}

/* Cabeçalho */
h1.page-title { color:%(PRIMARY_NAVY)s; margin:0; }
.page-sub { color:%(TEXT_MUTED)s; font-size:18px; margin-top:6px; }
.header-strip {
  position:absolute; right:60px; top:86px;
  background:%(PRIMARY_NAVY)s; color:#E6F2EA;
  padding:10px 20px; border-radius:12px; font-size:14px; letter-spacing:.2px;
}
.hr { height:2px; background:#E6ECF2; margin:20px 0 10px; border-radius:2px; }

/* Grid e Cards */
#cards-grid {
  display:grid; grid-template-columns:repeat(5, minmax(220px,1fr));
  gap:28px; margin-top:22px;
}
#cards-grid .stButton > button {
  width:100%; height:220px; text-align:center;
  border-radius:20px; border:1.5px solid %(CARD_BORDER)s;
  background:#FFFFFF; color:%(PRIMARY_NAVY)s;
  box-shadow:0 10px 30px rgba(15,42,58,.08);
  transition:transform .15s ease, box-shadow .2s ease, border-color .2s ease, background .2s ease;
  padding:18px 16px;
  line-height:1.25;
  white-space:pre-wrap;
  font-size:18px; font-weight:600;
}
#cards-grid .stButton > button:hover {
  transform: translateY(-2px);
  border-color:#D7E3F5;
  box-shadow:0 16px 36px rgba(30,64,175,.16);
  background:#FFFFFF;
}

/* Ícone circular */
.card-ico {
  display:inline-grid; place-items:center;
  width:68px; height:68px; border-radius:50%%;
  background:%(ACCENT_BLUE)s; color:#fff; font-size:30px;
  box-shadow:0 6px 16px rgba(30,64,175,.30);
  margin:2px auto 10px auto;
}

/* Subtexto e Chip */
.card-hint { display:block; margin-top:6px; color:%(TEXT_MUTED)s; font-weight:400; font-size:14px; }
.card-chip  { display:inline-block; margin-top:14px; padding:6px 18px; border-radius:20px;
              background:%(ACCENT_BLUE_LT)s; color:%(ACCENT_BLUE)s; border:1px solid #C9DAFF;
              font-size:14px; font-weight:600; letter-spacing:.2px; }

/* Responsivo */
@media (max-width:1400px){ #cards-grid{grid-template-columns:repeat(4,1fr);} }
@media (max-width:1100px){ #cards-grid{grid-template-columns:repeat(3,1fr);} }
@media (max-width:800px){  #cards-grid{grid-template-columns:repeat(2,1fr);} }
@media (max-width:520px){  #cards-grid{grid-template-columns:1fr;} }
</style>
""" % PALETTE, unsafe_allow_html=True)

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
# Lista de Cards (páginas)
# =========================================================
cards = [
    {"k":"1","title":"Cadastro de Vistorias","icon":"🗂️","hint":"Criar, editar e validar registros","path":"pages/Cadastro_de_vistorias.py"},
    {"k":"2","title":"Dashboard Operacional","icon":"📊","hint":"KPIs, prazos e mapa de calor","path":"pages/Dashboard_operacional.py"},
    {"k":"3","title":"Relatórios","icon":"📑","hint":"Emitir PDFs e planilhas","path":"pages/Relatorios.py"},
    {"k":"4","title":"Status / Andamento","icon":"🔄","hint":"Triagem e acompanhamento","path":"pages/Status_Andamento.py"},
    {"k":"5","title":"Auditoria","icon":"🕵️","hint":"Logs e rastreabilidade","path":"pages/Auditoria.py"},
]

# =========================================================
# Grid de Cards (com botões nativos)
# =========================================================
st.markdown('<div id="cards-grid">', unsafe_allow_html=True)

clicked_path = None
for c in cards:
    label = f"""
<span class="card-ico">{c['icon']}</span>
{c['title']}
<span class="card-hint">{c['hint']}</span>
<span class="card-chip">Abrir</span>
""".strip()
    if st.button(label, key=f"card-{c['k']}", use_container_width=True, help=c["title"]):
        clicked_path = c["path"]

st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# Redirecionamento se clicou
# =========================================================
if clicked_path:
    st.switch_page(clicked_path)

# =========================================================
# Atalhos 1–5 (acessam cada card)
# =========================================================
st.markdown("""
<script>
document.addEventListener('keydown', function(e){
  const map = {'1':'card-1','2':'card-2','3':'card-3','4':'card-4','5':'card-5'};
  if(map[e.key]){
    const list = Array.from(document.querySelectorAll('div#cards-grid .stButton button'));
    const idx = parseInt(e.key, 10) - 1;
    if (list[idx]) list[idx].click();
  }
});
</script>
""", unsafe_allow_html=True)

st.caption("Dica: use as teclas **1–5** para abrir as seções rapidamente.")
