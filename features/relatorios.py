# -*- coding: utf-8 -*-
import io
from datetime import datetime
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from core.data_loader import read_df, append_row
from core.models import RegistroRelatorio

def _render_pdf_buffer(dados: dict) -> bytes:
    # PDF mínimo (padrão NAOM simplificado)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, h-50, "Relatório de Vistoria — Padrão NAOM")
    c.setFont("Helvetica", 10)
    y = h-90
    for k, v in dados.items():
        c.drawString(40, y, f"{k}: {v}")
        y -= 16
        if y < 80:
            c.showPage()
            y = h-60
    c.showPage()
    c.save()
    return buf.getvalue()

def page():
    st.header("📄 VIS-005 — Geração de Relatório (NAOM)")

    df = read_df("solicitacoes")
    if df.empty:
        st.info("Sem solicitações para gerar relatório.")
        return

    numero = st.selectbox("Solicitação", df["numero"].tolist())
    titulo = st.text_input("Título do Relatório", value=f"Relatório de Vistoria {numero}")
    gerado_por = st.text_input("Assinatura/Responsável", value="Engº Militar")
    gerar = st.button("Gerar PDF")

    if gerar:
        reg = df[df["numero"] == numero].iloc[0].to_dict()
        dados = {
            "Número": numero,
            "Título": titulo,
            "OM": f'{reg.get("om_solicitante","")} — {reg.get("om_nome","")}',
            "Diretoria": reg.get("diretoria",""),
            "Local": reg.get("local",""),
            "Tipo de Vistoria": reg.get("tipo_vistoria",""),
            "Urgência": reg.get("urgencia",""),
            "Motivo": reg.get("motivo",""),
            "Data Limite": reg.get("data_limite",""),
            "Gerado por": gerado_por,
            "Gerado em": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }

        pdf_bytes = _render_pdf_buffer(dados)
        nome_pdf = f"Relatorio_{numero.replace('/','_')}.pdf"

        # salva metadado na aba Relatorios
        append_row("relatorios", RegistroRelatorio(
            numero=numero, titulo=titulo, arquivo_pdf=nome_pdf, gerado_por=gerado_por
        ).to_row())

        st.success("Relatório gerado.")
        st.download_button("⬇️ Baixar PDF", data=pdf_bytes, file_name=nome_pdf, mime="application/pdf")
