# -*- coding: utf-8 -*-
"""
features/dashboard.py
Dashboard Operacional — Vistorias CRO/1
"""

from __future__ import annotations
from datetime import datetime
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.data_loader import read_df
from core.config import TAB_SOLICITACOES
from core.utils import pick_col


# ----------------------- Helpers -----------------------
def _mes_label(s):
    s = pd.to_datetime(s, errors="coerce")
    return s.dt.to_period("M").dt.to_timestamp()

def _fmt_rs(v):
    if pd.isna(v): return "—"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
def _to_date(x):
    """Converte para Timestamp (brasil: dayfirst=True), devolvendo NaT em erros."""
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, (datetime, pd.Timestamp)):
        return pd.to_datetime(x, errors="coerce")
    return pd.to_datetime(x, errors="coerce", dayfirst=True)

def _safe_num(x):
    """Converte valores monetários brasileiros para float."""
    try:
        if pd.isna(x):
            return np.nan
        if isinstance(x, str):
            return float(x.replace("R$", "").replace(".", "").replace(",", ".").strip())
        return float(x)
    except Exception:
        return np.nan


# ----------------------- Main page Feature -----------------------

def page():
    st.caption("Use os filtros ao lado para refinar os indicadores.")
    st.header("📊 Dashboard Operacional — Seção de Vistorias")

    # ---- Carregamento principal ----
    try:
        df = read_df(TAB_SOLICITACOES)
    except Exception as e:
        st.error(f"Erro ao carregar a base de dados: {e}")
        return

    if df is None or df.empty:
        st.warning("Não há registros na base de solicitações.")
        return

    # ---- Mapeamento tolerante das colunas principais ----
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

    # ---- Normalização de datas e valores ----
    for k in ["dt_solic", "dt_real_visita", "dt_conc"]:
        if COL[k]:
            df[COL[k]] = df[COL[k]].map(_to_date)

    if COL["orcamento"]:
        df[COL["orcamento"]] = df[COL["orcamento"]].map(_safe_num)

    # ---- Emergencialidade (KPI) ----
    if COL["prioridade"]:
        df["Classificação"] = np.where(
            df[COL["prioridade"]].astype(str).str.contains("urg|emerg", case=False, na=False),
            "Emergencial", "Não Emergencial"
        )
    else:
        df["Classificação"] = "Não Informado"

    # ----------------------- Filtros (robustos) -----------------------
    with st.sidebar:
        st.header("Filtros")

        # Reset rápido
        if st.button("🔄 Limpar filtros"):
            for k in ("om", "especialidade", "status", "periodo"):
                st.session_state.pop(k, None)

        incluir_sem_data = False
        if COL["dt_solic"]:
            serie_datas = pd.to_datetime(df[COL["dt_solic"]], errors="coerce", dayfirst=True)
            min_d = serie_datas.min()
            max_d = serie_datas.max()
            if pd.isna(min_d) or pd.isna(max_d):
                min_d = pd.Timestamp(2025, 1, 1)
                max_d = pd.Timestamp.today()

            periodo = st.date_input(
                "Período (Data da Solicitação)",
                key="periodo",
                value=(min_d.date(), max_d.date()),
            )
            incluir_sem_data = st.checkbox("Incluir vistorias sem data de solicitação", value=True)
        else:
            periodo = (datetime(2025, 1, 1).date(), datetime.today().date())

        def _opts(key):
            col = COL.get(key)
            if col and col in df.columns:
                return sorted(df[col].dropna().astype(str).unique())
            return []

        om_sel  = st.multiselect("OM", _opts("om"), key="om")
        esp_sel = st.multiselect("Especialidade", _opts("especialidade"), key="especialidade")
        stat_sel= st.multiselect("Status", _opts("status"), key="status")

    # Aplicação dos filtros
    mask = pd.Series(True, index=df.index)

    if COL["dt_solic"] and COL["dt_solic"] in df.columns:
        datas = pd.to_datetime(df[COL["dt_solic"]], errors="coerce", dayfirst=True)
        no_intervalo = datas.between(
            pd.to_datetime(periodo[0]), pd.to_datetime(periodo[1]), inclusive="both"
        )
        mask &= (no_intervalo | (datas.isna() & incluir_sem_data))

    if om_sel and COL["om"]:
        mask &= df[COL["om"]].astype(str).isin(om_sel)
    if esp_sel and COL["especialidade"]:
        mask &= df[COL["especialidade"]].astype(str).isin(esp_sel)
    if stat_sel and COL["status"]:
        mask &= df[COL["status"]].astype(str).isin(stat_sel)

    dff = df[mask].copy()

    # Prazo de execução
    if COL["dt_solic"] and COL["dt_conc"]:
        dff["Tempo Execução (dias)"] = (dff[COL["dt_conc"]] - dff[COL["dt_solic"]]).dt.days

    # ----------------------- Diagnóstico dos filtros -----------------------
    with st.expander("🔎 Diagnóstico dos filtros (clique para abrir)"):
        total_base = len(df)
        total_filtrado = len(dff)
        st.write(f"Total na base: **{total_base}** | Após filtros: **{total_filtrado}**")
        if COL["dt_solic"]:
            datas_all = pd.to_datetime(df[COL["dt_solic"]], errors="coerce", dayfirst=True)
            fora_periodo = datas_all[
                ~datas_all.between(pd.to_datetime(periodo[0]), pd.to_datetime(periodo[1]), inclusive="both")
                & datas_all.notna()
            ]
            qtd_nat = datas_all.isna().sum()
            st.write(f"- Linhas **sem data de solicitação (NaT)**: **{qtd_nat}**")
            st.write(f"- Linhas **com data fora do período**: **{fora_periodo.shape[0]}**")
            if qtd_nat > 0 and not incluir_sem_data:
                st.info("Há vistorias **sem data**. Marque 'Incluir vistorias sem data de solicitação' para exibi-las.")

    # ----------------------- KPIs -----------------------
    st.subheader("Indicadores de Desempenho")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de Vistorias", len(dff))

    if "Classificação" in dff.columns and not dff.empty:
        base_emerg = dff["Classificação"].isin(["Emergencial", "Não Emergencial"])
        pct_emerg = (100 * dff.loc[base_emerg, "Classificação"].eq("Emergencial").mean()
                     if base_emerg.any() else np.nan)
        c2.metric("% Emergenciais", f"{pct_emerg:.1f}%" if pd.notna(pct_emerg) else "—")
    else:
        c2.metric("% Emergenciais", "—")

    if COL["orcamento"] and COL["orcamento"] in dff.columns:
        total_rs = dff[COL["orcamento"]].sum(skipna=True)
        c3.metric(
            "Orçamento Total",
            f"R$ {total_rs:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            if not np.isnan(total_rs) else "—"
        )
    else:
        c3.metric("Orçamento Total", "—")

    if "Tempo Execução (dias)" in dff.columns and dff["Tempo Execução (dias)"].notna().any():
        c4.metric("Prazo Médio de Execução", f"{dff['Tempo Execução (dias)'].mean():.1f} dias")
    else:
        c4.metric("Prazo Médio de Execução", "—")

    # ----------------------- Visualizações no estilo Excel -----------------------
    st.subheader("📈 Visualizações no estilo Excel")
    tabs = st.tabs(["Linha", "Colunas / Barras", "Pizza / Rosca", "Radar"])

    # Linha (multi-série por Status)
    with tabs[0]:
        if COL["dt_solic"] and COL["status"]:
            dff["_mes"] = _mes_label(dff[COL["dt_solic"]])
            base = (
                dff.groupby(["_mes", COL["status"]])
                   .size()
                   .reset_index(name="qtd")
                   .sort_values("_mes")
            )
            fig = px.line(
                base, x="_mes", y="qtd", color=COL["status"],
                markers=True, title="Evolução mensal por Status"
            )
            fig.update_traces(mode="lines+markers")
            fig.update_layout(xaxis_title="Mês", yaxis_title="Qtd", legend_title="Status")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Preciso de Data da Solicitação e Status para montar esta linha.")

    # Colunas / Barras
    with tabs[1]:
        col_a, col_b = st.columns(2)
        # Colunas agrupadas por mês e classificação
        with col_a:
            if COL["dt_solic"] and "Classificação" in dff.columns:
                dff["_mes"] = _mes_label(dff[COL["dt_solic"]])
                base = (
                    dff.groupby(["_mes", "Classificação"])
                       .size()
                       .reset_index(name="qtd")
                )
                fig = px.bar(
                    base, x="_mes", y="qtd", color="Classificação", barmode="group",
                    title="Colunas — por mês x classificação"
                )
                fig.update_layout(xaxis_title="Mês", yaxis_title="Qtd")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Preciso de Data da Solicitação e Classificação.")

        # Barras horizontais Top 10 OMs
        with col_b:
            if COL["om"]:
                top_om = (
                    dff.groupby(COL["om"]).size().nlargest(10)
                       .reset_index(name="Qtd")
                       .sort_values("Qtd")
                )
                fig = px.bar(
                    top_om, x="Qtd", y=COL["om"], orientation="h",
                    title="Barras — Top 10 OMs por quantidade"
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Preciso da coluna OM.")

    # Pizza / Rosca
    with tabs[2]:
        col_a, col_b = st.columns(2)
        # Pizza por Status
        with col_a:
            if COL["status"]:
                fig = px.pie(
                    dff, names=COL["status"], title="Pizza — participação por Status"
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Preciso da coluna Status.")
        # Rosca por Classificação (em R$ se houver orçamento)
        with col_b:
            if "Classificação" in dff.columns and COL["orcamento"]:
                base = dff.groupby("Classificação")[COL["orcamento"]].sum().reset_index()
                fig = px.pie(
                    base, names="Classificação", values=COL["orcamento"],
                    hole=0.55, title="Rosca — orçamento por classificação"
                )
                fig.update_traces(textposition="inside")
                st.plotly_chart(fig, use_container_width=True)
            elif "Classificação" in dff.columns:
                fig = px.pie(dff, names="Classificação", hole=0.55,
                             title="Rosca — participação por classificação")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Preciso da Classificação (Emergencial/Não).")

    # Radar (categorias em eixo polar)
    with tabs[3]:
        eixo = None
        if COL["especialidade"]:
            eixo = COL["especialidade"]
            titulo = "Radar — distribuição por Especialidade"
        elif COL["status"]:
            eixo = COL["status"]
            titulo = "Radar — distribuição por Status"
        if eixo:
            base = dff.groupby(eixo).size().reset_index(name="qtd")
            base = base.sort_values("qtd", ascending=False).head(10)  # mantém legível

            fig = go.Figure()
            fig.add_trace(go.Scatterpolar(
                r=base["qtd"], theta=base[eixo], fill="toself", name="Qtd"
            ))
            fig.update_layout(
                title=titulo,
                polar=dict(radialaxis=dict(visible=True)),
                showlegend=False,
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Preciso de Especialidade ou Status para o radar.")

    # ----------------------- Visualizações (as suas originais) -----------------------
    st.subheader("Visualizações Analíticas")

    if COL["dt_solic"]:
        dff[COL["dt_solic"]] = pd.to_datetime(dff[COL["dt_solic"]], errors="coerce", dayfirst=True)
        df_mes = (
            dff.groupby(dff[COL["dt_solic"]].dt.to_period("M").dt.to_timestamp())
               .size()
               .reset_index(name="qtd")
            .sort_values(COL["dt_solic"])
        )
        fig1 = px.line(df_mes, x=COL["dt_solic"], y="qtd",
                       title="Evolução Mensal das Vistorias", markers=True)
        fig1.update_traces(line_shape="spline")
        st.plotly_chart(fig1, use_container_width=True)

    if COL["status"]:
        fig2 = px.pie(dff, names=COL["status"], title="Distribuição por Status", hole=0.45)
        st.plotly_chart(fig2, use_container_width=True)

    if COL["om"]:
        top_om = dff.groupby(COL["om"]).size().nlargest(10).reset_index(name="Qtd")
        fig3 = px.bar(top_om, x="Qtd", y=COL["om"], orientation="h",
                      title="Top 10 OMs — Quantidade de Vistorias")
        st.plotly_chart(fig3, use_container_width=True)

    if COL["orcamento"] and "Classificação" in dff.columns:
        by_class = dff.groupby("Classificação")[COL["orcamento"]].sum().reset_index()
        fig4 = px.bar(by_class, x="Classificação", y=COL["orcamento"], text_auto=".2s",
                      title="Orçamento por Classificação")
        st.plotly_chart(fig4, use_container_width=True)

    # Aviso de colunas faltantes
    missing = [k for k, v in COL.items() if k in ["om", "status", "especialidade", "dt_solic"] and not v]
    if missing:
        st.warning(
            "Colunas não encontradas na base: " + ", ".join(missing)
            + ". O dashboard continua funcionando, mas alguns filtros/gráficos serão ocultados."
        )

    # ----------------------- Tabela e Exportação -----------------------
    st.subheader("Tabela Detalhada de Vistorias")
    st.dataframe(dff, use_container_width=True, hide_index=True)

    st.subheader("Exportar Relatório Resumido")
    if st.button("Gerar DOCX"):
        try:
            from docx import Document
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
            if "Classificação" in dff.columns and not dff.empty:
                doc.add_paragraph(f"% Emergenciais: {100 * dff['Classificação'].eq('Emergencial').mean():.1f}%")
            if COL["orcamento"] and COL["orcamento"] in dff.columns:
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
                    row_cells[j].text = "" if pd.isna(row[c]) else str(row[c])

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
