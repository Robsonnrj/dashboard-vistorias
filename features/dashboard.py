# -*- coding: utf-8 -*-
"""
features/dashboard.py
Dashboard Operacional — Vistorias CRO/1
Este arquivo implementa a feature do dashboard como função page(),
com filtros interativos, KPIs, gráficos, tabela e exportação DOCX.
"""

from datetime import datetime
import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st
from core.data_loader import read_df
from core.config import TAB_SOLICITACOES
from core.utils import pick_col

# Utilitário para converter entradas em datas robustas (dayfirst acertando casos brasileiros)
def _to_date(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, (datetime, pd.Timestamp)):
        return pd.to_datetime(x)
    try:
        return pd.to_datetime(x, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT

# Utilitário para converter valores orçamentários para float, aceitando string de moeda BR
def _safe_num(x):
    try:
        if pd.isna(x):
            return np.nan
        if isinstance(x, str):
            # Remove prefixo moeda e normaliza para float
            return float(x.replace("R$", "").replace(".", "").replace(",", ".").strip())
        return float(x)
    except Exception:
        return np.nan

# ----------------------- Main page Feature -----------------------

def page():
    st.caption("Use os filtros ao lado para refinar os indicadores.")
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    # ---- Carregamento principal da aba de dados ----
    try:
        df = read_df(TAB_SOLICITACOES)
    except Exception as e:
        st.error(f"Erro ao carregar a base de dados: {e}")
        return

    if df is None or df.empty:
        st.warning("Não há registros na base de solicitações.")
        return

    # ---- Mapeamento tolerante das colunas principais usando helper central ----
    COL = {
        "om": pick_col(df, ["OM", "OM beneficiada", "OM apoiada", "Organização Militar"]),
        "diretoria": pick_col(df, ["Diretoria", "Dir responsável", "DIR", "Direção"]),
        "especialidade": pick_col(df, [
            "Especialidade", "Especialidade envolvida", "Tipo/Especialidade",
            "Engenharia", "Área técnica", "Filtro - qual a especialidade da VT"
        ]),
        "prioridade": pick_col(df, [
            "Prioridade", "Classificação", "Tratativa da vistoria", "Classe da demanda",
            "Normal/Prioridade/Urgente/Urgentíssimo"
        ]),
        "status": pick_col(df, ["Status da Vistoria", "Status", "Situação", "Andamento"]),
        "dt_solic": pick_col(df, ["Data da solicitação", "Dt Solicitação", "Solicitado em"]),
        "dt_real_visita": pick_col(df, [
            "Data da realização da vistoria", "Data da vistoria", "Realização da visita", "Data visita"
        ]),
        "dt_conc": pick_col(df, [
            "Data da conclusão da VT", "Data da conclusão", "Conclusão da VT", "Conclusão"
        ]),
        "orcamento": pick_col(df, ["Orçamento estimado", "Valor estimado", "Custo", "PFR", "Total R$", "Orçamento"]),
    }

    # ---- Normalização dos dados de datas e valores ----
    for k in ["dt_solic", "dt_real_visita", "dt_conc"]:
        if COL[k]:
            df[COL[k]] = df[COL[k]].map(_to_date)

    if COL["orcamento"]:
        df[COL["orcamento"]] = df[COL["orcamento"]].map(_safe_num)

    # ---- Emergencialidade: usado para KPI e gráfico ----
    if COL["prioridade"]:
        df["Classificação"] = np.where(
            df[COL["prioridade"]].astype(str).str.contains("urg|emerg", case=False, na=False),
            "Emergencial", "Não Emergencial"
        )
    else:
        df["Classificação"] = "Não Informado"

    # ---- Sidebar: filtros dinâmicos ----
    with st.sidebar:
        st.header("Filtros")
        # Filtro por período baseado na data de solicitação
        if COL["dt_solic"]:
            min_d = pd.to_datetime(df[COL["dt_solic"]]).min()
            max_d = pd.to_datetime(df[COL["dt_solic"]]).max()
            periodo = st.date_input(
                "Período",
                (
                    min_d.date() if pd.notna(min_d) else datetime(2025,1,1).date(),
                    max_d.date() if pd.notna(max_d) else datetime.today().date(),
                ),
            )
        else:
            periodo = (datetime(2025,1,1).date(), datetime.today().date())

        # Helpers para montar opções de filtros (OM, Especialidade, Status)
        def _opts(key):
            col = COL.get(key)
            if col and col in df.columns:
                return sorted(df[col].dropna().astype(str).unique())
            return []
        om_sel  = st.multiselect("OM", _opts("om"))
        esp_sel = st.multiselect("Especialidade", _opts("especialidade"))
        stat_sel= st.multiselect("Status", _opts("status"))

    # ---- Aplicação dos filtros ----
    mask = pd.Series(True, index=df.index)
    if COL["dt_solic"] and COL["dt_solic"] in df.columns:
        mask &= df[COL["dt_solic"]].between(pd.to_datetime(periodo[0]), pd.to_datetime(periodo[1]))
    if om_sel and COL["om"]:
        mask &= df[COL["om"]].astype(str).isin(om_sel)
    if esp_sel and COL["especialidade"]:
        mask &= df[COL["especialidade"]].astype(str).isin(esp_sel)
    if stat_sel and COL["status"]:
        mask &= df[COL["status"]].astype(str).isin(stat_sel)

    dff = df[mask].copy()

    # ---- Prazo de execução (diferença entre datas, em dias) ----
    if COL["dt_solic"] and COL["dt_conc"]:
        dff["Tempo Execução (dias)"] = (dff[COL["dt_conc"]] - dff[COL["dt_solic"]]).dt.days

    # ---- KPIs ----
    st.subheader("Indicadores de Desempenho")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de Vistorias", len(dff))
    if "Classificação" in dff.columns:
        base_emerg = dff["Classificação"].isin(["Emergencial", "Não Emergencial"])
        pct_emerg = 100 * dff.loc[base_emerg, "Classificação"].eq("Emergencial").mean() if base_emerg.any() else np.nan
        c2.metric("% Emergenciais", f"{pct_emerg:.1f}%" if pd.notna(pct_emerg) else "—")
    else:
        c2.metric("% Emergenciais", "—")
    if COL["orcamento"]:
        total_rs = dff[COL["orcamento"]].sum()
        c3.metric("Orçamento Total", f"R$ {total_rs:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    else:
        c3.metric("Orçamento Total", "—")
    if "Tempo Execução (dias)" in dff.columns and dff["Tempo Execução (dias)"].notna().any():
        c4.metric("Prazo Médio de Execução", f"{dff['Tempo Execução (dias)'].mean():.1f} dias")
    else:
        c4.metric("Prazo Médio de Execução", "—")

    # ---- Visualizações Analíticas (gráficos interativos) ----
    st.subheader("Visualizações Analíticas")
    if COL["dt_solic"]:
        df_mes = (
            dff.groupby(dff[COL["dt_solic"]].dt.to_period("M").dt.to_timestamp())
            .size()
            .reset_index(name="qtd")
        )
        fig1 = px.line(df_mes, x=COL["dt_solic"], y="qtd", title="Evolução Mensal das Vistorias", markers=True)
        fig1.update_traces(line_shape="spline")
        st.plotly_chart(fig1, use_container_width=True)

    if COL["status"]:
        fig2 = px.pie(dff, names=COL["status"], title="Distribuição por Status", hole=0.45)
        st.plotly_chart(fig2, use_container_width=True)

    if COL["om"]:
        top_om = dff.groupby(COL["om"]).size().nlargest(10).reset_index(name="Qtd")
        fig3 = px.bar(top_om, x="Qtd", y=COL["om"], orientation="h", title="Top 10 OMs — Quantidade de Vistorias")
        st.plotly_chart(fig3, use_container_width=True)

    if COL["orcamento"] and "Classificação" in dff.columns:
        by_class = dff.groupby("Classificação")[COL["orcamento"]].sum().reset_index()
        fig4 = px.bar(by_class, x="Classificação", y=COL["orcamento"], text_auto=".2s", title="Orçamento por Classificação")
        st.plotly_chart(fig4, use_container_width=True)

    # ---- Alerta sobre colunas faltantes (manutenção e troubleshooting) ----
    missing = [k for k, v in COL.items() if k in ["om", "status", "especialidade", "dt_solic"] and not v]
    if missing:
        st.warning(
            "Colunas não encontradas na base: " + ", ".join(missing)
            + ". O dashboard continua funcionando, mas alguns filtros/gráficos serão ocultados."
        )

    # ---- Tabela detalhada final (mostrando os dados filtrados) ----
    st.subheader("Tabela Detalhada de Vistorias")
    st.dataframe(dff, use_container_width=True, hide_index=True)

    # ---- Exportação DOCX resumida (até 25 linhas) ----
    st.subheader("Exportar Relatório Resumido")
    if st.button("Gerar DOCX"):
        try:
            from docx import Document  # <-- precisa do pacote python-docx instalado
        except ModuleNotFoundError:
            st.error(
                "Pacote **python-docx** não está instalado. "
                "Adicione `python-docx==0.8.11` ao seu `requirements.txt` e faça o deploy novamente."
            )
        else:
            doc = Document()
            doc.add_heading("Relatório Resumido — Vistorias CRO/1", 0)
            doc.add_paragraph(f"Gerado em {datetime.now():%d/%m/%Y %H:%M}")

            doc.add_heading("Indicadores Principais", level=2)
            doc.add_paragraph(f"Total de Vistorias: {len(dff)}")
            doc.add_paragraph(f"% Emergenciais: {100 * dff['Classificação'].eq('Emergencial').mean():.1f}%")
            if COL["orcamento"]:
                doc.add_paragraph(
                    f"Orçamento Total: R$ {dff[COL['orcamento']].sum():,.2f}"
                    .replace(",", "X").replace(".", ",").replace("X", ".")
                )
            if "Tempo Execução (dias)" in dff.columns:
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

            import io
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            st.download_button(
                "📄 Baixar Relatório DOCX",
                data=bio,
                file_name="Relatorio_Vistorias_CRO1.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
