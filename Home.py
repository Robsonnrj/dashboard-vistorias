from __future__ import annotations
import streamlit.components.v1 as components
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
# Config da página
# =========================================================
st.set_page_config(
    page_title="CRO/1 — Sistema de Vistorias",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="collapsed",
)
st.session_state.setdefault("tabs_map", dict(SHEET_TABS))

# =========================================================
# Paleta — Azul institucional
# =========================================================
PRIMARY_NAVY   = "#0F2A3A"  # faixa/cabeçalho
ACCENT_BLUE    = "#1E40AF"  # ícones/botões
ACCENT_BLUE_LT = "#E8F0FF"  # chip "Abrir"
TEXT_MUTED     = "#64748B"
CARD_BG        = "#FFFFFF"
CARD_BORDER    = "#E8EEF5"
DIVIDER        = "#E6ECF2"
BG             = "#F7F9FC"

# =========================================================
# Estilo
# =========================================================
st.markdown(f"""
<style>
/* fundo app */
[data-testid="stAppViewContainer"] {{
  background: {BG};
}}
.block-container{{padding-top:2.2rem;padding-bottom:3rem;}}

/* título */
h1.page-title {{
  color:{PRIMARY_NAVY}; margin:0;
}}
.page-sub {{
  color:{TEXT_MUTED}; font-size:18px; margin-top:6px;
}}

/* faixa institucional */
.header-strip {{
  position:absolute; right:60px; top:86px;
  background:{PRIMARY_NAVY}; color:#E6F2EA;
  padding:10px 20px; border-radius:12px; font-size:14px;
  letter-spacing:.2px;
}}
.hr {{ height:2px; background:{DIVIDER}; margin:20px 0 10px 0; border-radius:2px; }}

/* grid dos cards */
.cards {{
  display:grid;
  grid-template-columns: repeat(5, minmax(220px,1fr));
  gap:28px; margin-top:22px;
}}
.card {{
  position:relative;
  background:{CARD_BG};
  border:1.5px solid {CARD_BORDER};
  border-radius:20px;
  padding:26px 22px 74px 22px;
  box-shadow:0 10px 30px rgba(15,42,58,0.08);
  transition: transform .15s ease, box-shadow .2s ease, border-color .2s ease;
  text-decoration:none !important; display:block;
}}
.card:hover {{
  transform: translateY(-2px);
  border-color:#D7E3F5;
  box-shadow:0 16px 36px rgba(30,64,175,0.16);
}}
.icon-circle {{
  width:68px; height:68px; background:{ACCENT_BLUE}; color:#fff;
  display:grid; place-items:center; border-radius:50%;
  font-size:30px; margin:2px auto 12px auto;
  box-shadow:0 6px 16px rgba(30,64,175,0.30);
}}
.card h3 {{
  color:{PRIMARY_NAVY}; font-size:20px; line-height:1.25;
  text-align:center; margin:6px 0 8px 0;
}}
.card .hint {{
  color:{TEXT_MUTED}; text-align:center; font-size:14px; min-height:42px;
}}
.card .chip {{
  position:absolute; left:50%; transform:translateX(-50%); bottom:18px;
  background:{ACCENT_BLUE_LT}; border:1px solid #C9DAFF;
  color:{ACCENT_BLUE}; padding:6px 18px; border-radius:20px;
  font-weight:600; font-size:14px; letter-spacing:.2px;
}}

@media (max-width: 1400px) {{ .cards {{ grid-template-columns: repeat(4, 1fr); }} }}
@media (max-width: 1100px) {{ .cards {{ grid-template-columns: repeat(3, 1fr); }} }}
@media (max-width: 800px)  {{ .cards {{ grid-template-columns: repeat(2, 1fr); }} }}
@media (max-width: 520px)  {{ .cards {{ grid-template-columns: 1fr; }} }}
</style>
""", unsafe_allow_html=True)

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
# Definição dos cards (com seus paths reais)
# =========================================================

cards = [
    {"k":"1","title":"Cadastro de Vistorias","hint":"Criar, editar e validar registros","path":"pages/Cadastro_de_vistorias.py","svg":"""
      <svg width="28" height="28" viewBox="0 0 24 24" fill="none"><path d="M3 7h7l2 2h9v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z" stroke="white" stroke-width="1.8" stroke-linejoin="round"/><path d="M3 7V5a2 2 0 0 1 2-2h5l2 2h5a2 2 0 0 1 2 2" stroke="white" stroke-width="1.8" stroke-linecap="round"/></svg>
    """},
    {"k":"2","title":"Dashboard Operacional","hint":"KPIs, prazos e mapa de calor","path":"pages/Dashboard_operacional.py","svg":"""
      <svg width="28" height="28" viewBox="0 0 24 24" fill="none"><path d="M4 13v7M10 4v16M16 9v11M22 2v18" stroke="white" stroke-width="1.8" stroke-linecap="round"/></svg>
    """},
    {"k":"3","title":"Relatórios","hint":"Emitir PDFs e planilhas","path":"pages/Relatorios.py","svg":"""
      <svg width="28" height="28" viewBox="0 0 24 24" fill="none"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z" stroke="white" stroke-width="1.8"/><path d="M14 2v6h6" stroke="white" stroke-width="1.8"/><path d="M8 13h8M8 17h8" stroke="white" stroke-width="1.6" stroke-linecap="round"/></svg>
    """},
    {"k":"4","title":"Status / Andamento","hint":"Triagem e acompanhamento","path":"pages/Status_Andamento.py","svg":"""
      <svg width="28" height="28" viewBox="0 0 24 24" fill="none"><path d="M21 12a9 9 0 1 1-3.2-6.9" stroke="white" stroke-width="1.8"/><path d="M21 3v6h-6" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
    """},
    {"k":"5","title":"Auditoria","hint":"Logs e rastreabilidade","path":"pages/Auditoria.py","svg":"""
      <svg width="28" height="28" viewBox="0 0 24 24" fill="none"><path d="M12 22s8-4 8-10V6l-8-4-8 4v6c0 6 8 10 8 10z" stroke="white" stroke-width="1.8"/><path d="M9 12l2 2 4-4" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
    """},
]

ACCENT_BLUE    = "#1E40AF"
ACCENT_BLUE_LT = "#E8F0FF"
PRIMARY_NAVY   = "#0F2A3A"
TEXT_MUTED     = "#64748B"
CARD_BORDER    = "#E8EEF5"

cards_html = """
<style>
.menu-grid{
  display:grid; grid-template-columns:repeat(5,minmax(220px,1fr)); gap:28px; margin-top:22px;
}
.menu-card{
  position:relative; background:#fff; border:1.5px solid %(CARD_BORDER)s;
  border-radius:20px; padding:26px 22px 74px; text-decoration:none;
  box-shadow:0 10px 30px rgba(15,42,58,.08);
  transition:transform .15s ease, box-shadow .2s ease, border-color .2s ease;
  display:block; color:inherit;
}
.menu-card:hover{ transform:translateY(-2px); border-color:#D7E3F5; box-shadow:0 16px 36px rgba(30,64,175,.16); }
.icon-circle{
  width:68px; height:68px; border-radius:50%%; background:%(ACCENT_BLUE)s; display:grid; place-items:center;
  margin:2px auto 12px; box-shadow:0 6px 16px rgba(30,64,175,.30);
}
.menu-card h3{ color:%(PRIMARY_NAVY)s; font-size:20px; line-height:1.25; text-align:center; margin:6px 0 8px; }
.menu-card .hint{ color:%(TEXT_MUTED)s; text-align:center; font-size:14px; min-height:42px; }
.menu-card .chip{
  position:absolute; left:50%%; transform:translateX(-50%%); bottom:18px;
  background:%(ACCENT_BLUE_LT)s; border:1px solid #C9DAFF; color:%(ACCENT_BLUE)s;
  padding:6px 18px; border-radius:20px; font-weight:600; font-size:14px;
}
@media (max-width:1400px){ .menu-grid{grid-template-columns:repeat(4,1fr);} }
@media (max-width:1100px){ .menu-grid{grid-template-columns:repeat(3,1fr);} }
@media (max-width:800px){ .menu-grid{grid-template-columns:repeat(2,1fr);} }
@media (max-width:520px){ .menu-grid{grid-template-columns:1fr;} }
</style>
<div class="menu-grid">
""" % {
    "ACCENT_BLUE": ACCENT_BLUE,
    "ACCENT_BLUE_LT": ACCENT_BLUE_LT,
    "PRIMARY_NAVY": PRIMARY_NAVY,
    "TEXT_MUTED": TEXT_MUTED,
    "CARD_BORDER": CARD_BORDER,
}

for c in cards:
    cards_html += f"""
    <a class="menu-card" href="?nav={c['path']}" id="card-{c['k']}" aria-label="{c['title']}">
      <div class="icon-circle">{c['svg']}</div>
      <h3>{c['title']}</h3>
      <div class="hint">{c['hint']}</div>
      <div class="chip">Abrir</div>
    </a>
    """

cards_html += """
</div>
<script>
document.addEventListener('keydown', function(e){
  const map={'1':'card-1','2':'card-2','3':'card-3','4':'card-4','5':'card-5'};
  if(map[e.key]){ const a=document.getElementById(map[e.key]); if(a) a.click(); }
});
</script>
"""

# Renderiza os cards em um iframe (não escapa o HTML)
components.html(cards_html, height=560, scrolling=False)
# ===== FIM DO TRECHO SUBSTITUÍDO =====
