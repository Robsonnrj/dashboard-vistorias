# -*- coding: utf-8 -*-
"""
Dashboard Operacional — Vistorias CRO/1
Versão adaptada para estrutura modular do projeto IME / CRO1
Autor: Robson Nunes Rodrigues Junior
Data: Atualizado em 13/10/2025
"""

from __future__ import annotations
import io
from datetime import datetime
import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st

# Importações do core do projeto
from core.data_loader import read_df
from core.config import TAB_SOLICITACOES

# =========================================================
# Funções auxiliares
# =========================================================
def _nf(s: str) -> str:
    return (s or "").casefold().strip()

def _pick(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Busca coluna por nome aproximado (case insensitive)."""
    if df is None or df.empty:
        return None
    for c in candidates:
        for col in df.columns:
            if _nf(col) == _nf(c) or _nf(c) in _nf(col):
                return col
    return None

def _to_date(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, (datetime, pd.Timestamp)):
        return pd.to_datetime(x)
    try:
        return pd.to_datetime(x, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT

def _safe_num(x):
    try:
        if pd.isna(x):
            return np.nan
        if isinstance(x, str):
            return float(x.replace("R$", "").replace(".", "").replace(",", ".").strip())
        return float(x)
    except Exception:
        return np.nan

# =========================================================
# Função principal do Dashboard
# =========================================================
def main():
    st.set_page_config(
        page_title="Dashboard CRO/1 — Vistorias",
        layout="wide",
        page_icon="📊"
    )

    st.title("📊 Dashboard Operacional — CRO/1 (Vistorias Técnicas)")

    # -----------------------------------------------------
    # Leitura da base
    # -----------------------------------------------------
    try:
        df_raw = read_df(TAB_SOLICITACOES)
    except Exception as e:
        st.error(f"Erro ao carregar a base de dados: {e}")
        return

    if df_raw is None or df_raw.empty:
        st.warning("Não há registros na base de solicitações (TAB_SOLICITACOES).")
        return

    df = df_raw.copy()

    # -----------------------------------------------------
    # Identificação automática de colunas
    # -----------------------------------------------------
    COL = {
        "om": _pick(df, ["OM", "OM beneficiada", "OM apoiada"]),
        "diretoria": _pick(df, ["Diretoria", "Dir responsável"]),
        "especialidade": _pick(df, ["Especialidade", "Engenharia"]),
        "prioridade": _pick(df, ["Prioridade", "Classificação"]),
        "status": _pick(df, ["Status", "Situação"]),
        "dt_solic": _pick(df, ["Data da solicitação"]),
        "dt_real_visita": _pick(df, ["Data da vistoria", "Data de realização"]),
        "dt_conc": _pick(df, ["Data da conclusão", "Conclusão"]),
        "orcamento": _pick(df, ["Orçamento", "Custo estimado", "Valor"]),
    }

    for k in ["dt_solic", "dt_real_visita", "dt_conc"]:
        if COL[k]:
            df[COL[k]] = df[COL[k]].map(_to_date)

    if COL["orcamento"]:
        df[COL["orcamento"]] = df[COL["orcamento"]].map(_safe_num)

    # Campo de emergência inferido
    df["Classificação"] = np.where(
        df[COL["prioridade"]].astype(str).str.contains("Urgente|Emerg", case=False, na=False),
        "Emergencial", "Não Emergencial"
    )

    # =========================================================
    # Filtros laterais
    # =========================================================
    with st.sidebar:
        st.header("Filtros")
        if COL["dt_solic"]:
            min_d = df[COL["dt_solic"]].min()
            max_d = df[COL["dt_solic"]].max()
            periodo = st.date_input("Período", (min_d.date(), max_d.date()))
        else:
            periodo = (datetime(2025, 1, 1), datetime.today())

        om_sel = st.multiselect("OM", sorted(df[COL["om"]].dropna().unique()))
        esp_sel = st.multiselect("Especialidade", sorted(df[COL["especialidade"]].dropna().unique()))
        stat_sel = st.multiselect("Status", sorted(df[COL["status"]].dropna().unique()))

    mask = pd.Series(True, index=df.index)
    if COL["dt_solic"]:
        mask &= df[COL["dt_solic"]].between(pd.to_datetime(periodo[0]), pd.to_datetime(periodo[1]))
    if om_sel:
        mask &= df[COL["om"]].isin(om_sel)
    if esp_sel:
        mask &= df[COL["especialidade"]].isin(esp_sel)
    if stat_sel:
        mask &= df[COL["status"]].isin(stat_sel)

    dff = df[mask].copy()

    # =========================================================
    # Indicadores principais (KPIs)
    # =========================================================
    st.subheader("Indicadores de Desempenho")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total de Vistorias", len(dff))
    col2.metric("% Emergenciais", f"{100 * dff['Classificação'].eq('Emergencial').mean():.1f}%")
    if COL["orcamento"]:
        col3.metric("Orçamento Total", f"R$ {dff[COL['orcamento']].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    else:
        col3.metric("Orçamento Total", "—")
    if COL["dt_solic"] and COL["dt_conc"]:
        dff["Tempo Execução (dias)"] = (dff[COL["dt_conc"]] - dff[COL["dt_solic"]]).dt.days
        media = dff["Tempo Execução (dias)"].mean()
        col4.metric("Prazo Médio de Execução", f"{media:.1f} dias")
    else:
        col4.metric("Prazo Médio de Execução", "—")

    # =========================================================
    # Gráficos principais
    # =========================================================
    st.subheader("Visualizações Analíticas")

    # 1. Gráfico temporal (função)
    if COL["dt_solic"]:
        df_mes = dff.groupby(dff[COL["dt_solic"]].dt.to_period("M").dt.to_timestamp()).size().reset_index(name="qtd")
        fig1 = px.line(df_mes, x=COL["dt_solic"], y="qtd", title="Evolução Mensal das Vistorias", markers=True)
        fig1.update_traces(line_shape="spline")
        st.plotly_chart(fig1, use_container_width=True)

    # 2. Distribuição por Status
    if COL["status"]:
        fig2 = px.pie(dff, names=COL["status"], title="Distribuição por Status", hole=0.45)
        st.plotly_chart(fig2, use_container_width=True)

    # 3. Top OMs
    if COL["om"]:
        top_om = dff.groupby(COL["om"]).size().nlargest(10).reset_index(name="Qtd")
        fig3 = px.bar(top_om, x="Qtd", y=COL["om"], orientation="h", title="Top 10 OMs — Quantidade de Vistorias")
        st.plotly_chart(fig3, use_container_width=True)

    # 4. Orçamento por Classificação
    if COL["orcamento"]:
        by_class = dff.groupby("Classificação")[COL["orcamento"]].sum().reset_index()
        fig4 = px.bar(by_class, x="Classificação", y=COL["orcamento"], text_auto=".2s", title="Orçamento por Classificação")
        st.plotly_chart(fig4, use_container_width=True)

    # =========================================================
    # Tabela detalhada
    # =========================================================
    st.subheader("Tabela Detalhada de Vistorias")
    st.dataframe(dff, use_container_width=True, hide_index=True)

    # =========================================================
    # Exportação
    # =========================================================
    st.subheader("Exportar Relatório Resumido")

    if st.button("Gerar DOCX"):
        from docx import Document
        doc = Document()
        doc.add_heading("Relatório Resumido — Vistorias CRO/1", 0)
        doc.add_paragraph(f"Gerado em {datetime.now():%d/%m/%Y %H:%M}")

        doc.add_heading("Indicadores Principais", level=2)
        doc.add_paragraph(f"Total de Vistorias: {len(dff)}")
        doc.add_paragraph(f"% Emergenciais: {100 * dff['Classificação'].eq('Emergencial').mean():.1f}%")
        if COL["orcamento"]:
            doc.add_paragraph(f"Orçamento Total: R$ {dff[COL['orcamento']].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        if "Tempo Execução (dias)" in dff:
            doc.add_paragraph(f"Prazo Médio de Execução: {dff['Tempo Execução (dias)'].mean():.1f} dias")

        doc.add_heading("Tabela de Vistorias", level=2)
        t = doc.add_table(rows=1, cols=len(dff.columns))
        hdr_cells = t.rows[0].cells
        for j, c in enumerate(dff.columns):
            hdr_cells[j].text = c
        for _, row in dff.head(25).iterrows():
            row_cells = t.add_row().cells
            for j, c in enumerate(dff.columns):
                row_cells[j].text = str(row[c]) if pd.notna(row[c]) else ""

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        st.download_button(
            "📄 Baixar Relatório DOCX",
            data=buffer,
            file_name="Relatorio_Vistorias_CRO1.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# =========================================================
# Execução direta
# =========================================================
if __name__ == "__main__":
    main()
