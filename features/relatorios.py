# -*- coding: utf-8 -*-
"""
features/relatorios.py
Geração de Relatórios NAOM (PDF) via reportlab.
Exibe page() para integração com o roteador.
"""

from __future__ import annotations

import streamlit as st
import pandas as pd

# carrega dados se precisar buscar algo da base
from core.data_loader import read_df
from core.layout import hide_multipage_nav, top_nav
from features.relatorios import page as relatorios_page

def page():
    hide_multipage_nav()
    top_nav("Relatorios")

    relatorios_page()

# dependência opcional
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    HAVE_REPORTLAB = True
except Exception:
    HAVE_REPORTLAB = False


def gerar_pdf_relatorio_naom(dados_relatorio: dict, caminho_pdf: str):
    if not HAVE_REPORTLAB:
        raise RuntimeError(
            "O módulo 'reportlab' não está instalado. "
            "Adicione 'reportlab' ao requirements.txt e faça novo deploy."
        )

    c = canvas.Canvas(caminho_pdf, pagesize=A4)
    w, h = A4

    # Cabeçalho simples — ajuste conforme seu template
    c.setFont("Helvetica-Bold", 14)
    c.drawString(25 * mm, h - 25 * mm, "Relatório de Vistoria - NAOM")

    c.setFont("Helvetica", 10)
    y = h - 35 * mm
    for k, v in dados_relatorio.items():
        c.drawString(25 * mm, y, f"{k}: {v}")
        y -= 6 * mm
        if y < 20 * mm:
            c.showPage()
            c.setFont("Helvetica", 10)
            y = h - 25 * mm

    c.showPage()
    c.save()


def ui_relatorios():
    st.header("🧾 Relatórios de Vistoria (NAOM)")

    if not HAVE_REPORTLAB:
        st.warning(
            "Pacote **reportlab** ausente. "
            "Adicione `reportlab` ao requirements.txt para habilitar a geração de PDF."
        )
        return

    # Campos básicos (adicione os que quiser)
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
                st.download_button(
                    "⬇️ Baixar PDF",
                    f,
                    file_name=caminho,
                    mime="application/pdf",
                    use_container_width=True,
                )
            st.success("PDF gerado com sucesso.")
        except Exception as e:
            st.error(f"Falha ao gerar PDF: {e}")


def page():
    """Wrapper chamado pelo roteador do app."""
    ui_relatorios()
