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
    {"k":"1", "title":"Cadastro de Vistorias",  "icon":"🗂️", "hint":"Criar, editar e validar registros", "path":"pages/Cadastro_de_vistorias.py"},
    {"k":"2", "title":"Dashboard Operacional", "icon":"📊", "hint":"KPIs, prazos e mapa de calor",       "path":"pages/Dashboard_operacional.py"},
    {"k":"3", "title":"Relatórios",            "icon":"📑", "hint":"Emitir PDFs e planilhas",            "path":"pages/Relatorios.py"},
    {"k":"4", "title":"Status / Andamento",    "icon":"🔄", "hint":"Triagem e acompanhamento",          "path":"pages/Status_Andamento.py"},
    {"k":"5", "title":"Auditoria",             "icon":"🕵️", "hint":"Logs e rastreabilidade",            "path":"pages/Auditoria.py"},
]

# =========================================================
# Router via query string (?nav=pages/...py)
# — Clicar no card (ou atalho do teclado) seta nav e navega
# =========================================================
nav = st.query_params.get("nav", [None])
nav = nav[0] if isinstance(nav, list) else nav
if nav:
    # Navega para a página real
    st.switch_page(nav)

# =========================================================
# Renderização dos cards (HTML clicável)
# =========================================================
html = ['<div class="cards">']
for c in cards:
    href = f"?nav={c['path']}"
    html.append(f"""
    <a class="card" href="{href}" id="card-{c['k']}" data-key="{c['k']}">
      <div class="icon-circle">{c['icon']}</div>
      <h3>{c['title']}</h3>
      <div class="hint">{c['hint']}</div>
      <div class="chip">Abrir</div>
    </a>
    """)
html.append("</div>")
st.markdown("\n".join(html), unsafe_allow_html=True)

# =========================================================
# Atalhos de teclado 1–5 (aciona os cards)
# =========================================================
st.markdown("""
<script>
document.addEventListener('keydown', function(e) {{
  const map = {{'1':'card-1','2':'card-2','3':'card-3','4':'card-4','5':'card-5'}};
  if (map[e.key]) {{
    const el = document.getElementById(map[e.key]);
    if (el) el.click();
  }}
}});
</script>
""", unsafe_allow_html=True)

# =========================================================
# Rodapé
# =========================================================
st.caption("Dica: use as teclas **1–5** para abrir as seções rapidamente.")
