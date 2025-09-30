# features/relatorios.py

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    HAVE_REPORTLAB = True
except ImportError:
    HAVE_REPORTLAB = False

import streamlit as st
import pandas as pd
from core.data_loader import read_df

def gerar_pdf_relatorio_naom(dados_relatorio: dict, caminho_pdf: str):
    if not HAVE_REPORTLAB:
        raise RuntimeError(
            "O módulo 'reportlab' não está instalado. "
            "Adicione 'reportlab' ao requirements.txt e reimplante."
        )
    c = canvas.Canvas(caminho_pdf, pagesize=A4)
    w, h = A4

    # Cabeçalho simples — substitua pelo seu template NAOM
    c.setFont("Helvetica-Bold", 14)
    c.drawString(25*mm, h - 25*mm, "Relatório de Vistoria - NAOM")
    c.setFont("Helvetica", 10)
    y = h - 35*mm
    for k, v in dados_relatorio.items():
        c.drawString(25*mm, y, f"{k}: {v}")
        y -= 6*mm
        if y < 20*mm:
            c.showPage()
            y = h - 25*mm
    c.showPage()
    c.save()

def ui_relatorios():
    st.header("🧾 Relatórios de Vistoria (NAOM)")
    if not HAVE_REPORTLAB:
        st.warning("Pacote 'reportlab' ausente. Adicione ao requirements.txt para habilitar o PDF.")
        st.stop()

    # Exemplo mínimo de interface
    col1, col2 = st.columns(2)
    with col1:
        numero = st.text_input("Número do Relatório NAOM")
        om = st.text_input("OM")
        fiscal = st.text_input("Fiscal Responsável")
        data = st.date_input("Data do Relatório")
    with col2:
        objetivo = st.text_area("Objetivo")
        conclusoes = st.text_area("Conclusões")
        recomendacoes = st.text_area("Recomendações")

    if st.button("Gerar PDF"):
        info = {
            "Número": numero,
            "OM": om,
            "Fiscal": fiscal,
            "Data": str(data),
            "Objetivo": objetivo,
            "Conclusões": conclusoes,
            "Recomendações": recomendacoes,
        }
        caminho = "relatorio_vistoria_naom.pdf"
        try:
            gerar_pdf_relatorio_naom(info, caminho)
            with open(caminho, "rb") as f:
                st.download_button("⬇️ Baixar PDF", f, file_name=caminho, mime="application/pdf")
            st.success("PDF gerado com sucesso.")
        except Exception as e:
            st.error(f"Falha ao gerar PDF: {e}")
