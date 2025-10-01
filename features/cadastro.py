# -*- coding: utf-8 -*-
import streamlit as st
from datetime import date
import pandas as pd

from core.data_loader import append_row, read_df
from core.config import TAB_SOLICITACOES, TAB_VALIDACAO


def _clean(x): return "" if pd.isna(x) else str(x).strip()


def _load_oms_map() -> tuple[list[str], dict[str, str], dict[str, str]]:
    """options, display->sigla, sigla->diretoria"""
    options, disp2sig, sig2dir = [], {}, {}

    for tab in (TAB_VALIDACAO, TAB_SOLICITACOES):
        try:
            df = read_df(tab)
        except Exception:
            continue
        if df.empty:
            continue

        cols = {c.lower(): c for c in df.columns}
        c_sig = cols.get("om") or cols.get("om apoiada") or cols.get("sigla")
        c_nom = cols.get("organização militar") or cols.get("organização") or cols.get("om")
        c_dir = cols.get("diretoria responsável") or cols.get("diretoria")
        if not c_sig or not c_dir:
            continue

        tmp = pd.DataFrame({"sig": df[c_sig].map(_clean)})
        tmp["nome"] = df[c_nom].map(_clean) if c_nom else ""
        tmp["dir"] = df[c_dir].map(_clean)
        tmp = tmp[tmp["sig"] != ""].drop_duplicates("sig")

        for _, r in tmp.iterrows():
            label = f"{r['sig']} — {r['nome']}" if r["nome"] else r["sig"]
            if label not in options:
                options.append(label)
                disp2sig[label] = r["sig"]
                sig2dir[r["sig"]] = r["dir"]
        break

    options.append("Outra / não listada…")
    disp2sig["Outra / não listada…"] = ""
    return options, disp2sig, sig2dir


def _input_row(oms_df: pd.DataFrame | None):
    st.subheader("📥 Nova solicitação de vistoria")

    options, disp2sig, sig2dir = _load_oms_map()

    col1, col2 = st.columns(2)
    with col1:
        om_display = st.selectbox("OM APOIADA *", options, index=None, placeholder="Digite para buscar…")
        om_sigla = disp2sig.get(om_display or "", "")
        diretoria = st.text_input("Diretoria Responsável *", value=sig2dir.get(om_sigla, ""))
        objeto = st.text_input("OBJETO DE VISTORIA *", value="")
        urgencia = st.selectbox("Classificação de Urgência", ["Não Prioridade", "Prioridade", "Urgente"])
    with col2:
        situacao = st.selectbox("Situação", ["Não Atendida", "Em andamento", "Finalizada"])
        data_solic = st.date_input("DATA DA SOLICITAÇÃO", value=date.today())
        data_vist = st.date_input("DATA DA VISTORIA", value=None)

    st.markdown("### Complementares")
    colA, colB = st.columns(2)
    with colA:
        ref_opus = st.text_input("REFERÊNCIA OPUS")
        objetivo = st.text_input("OBJETIVO (ADICIONAR POSSÍVEL CONTATO)")
        vt_exec = st.text_input("VT EXECUTADA POR")
        status_sem = st.text_input("STATUS - ATUALIZAÇÃO SEMANAL")
    with colB:
        meio_resp = st.text_input("MEIO DE RESPOSTA DA SOLICITAÇÃO")
        opus_vist = st.text_input("Nº OPUS DA VISTORIA (SE FOR O CASO)")
        qt_total = st.number_input("QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO", min_value=0, step=1)
        qt_exec = st.number_input("QUANTIDADE DE DIAS PARA EXECUÇÃO", min_value=0, step=1)

    obs = st.text_area("OBSERVAÇÕES", height=100)

    erros = []
    if not objeto.strip(): erros.append("Informe o **OBJETO DE VISTORIA**.")
    if not om_sigla: erros.append("Selecione/informe a **OM APOIADA**.")
    if not diretoria.strip(): erros.append("Informe a **Diretoria Responsável**.")

    if erros:
        st.warning("• " + "<br>• ".join(erros), unsafe_allow_html=True)

    row = {
        "OBJETO DE VISTORIA": objeto.strip(),
        "OM APOIADA": om_sigla,
        "Diretoria Responsável": diretoria.strip(),
        "Classificação de Urgência": urgencia,
        "Situação": situacao,
        "DATA DA SOLICITAÇÃO": pd.to_datetime(data_solic).strftime("%Y-%m-%d"),
        "DATA DA SOLICITAÇÃO_2": pd.to_datetime(data_solic).strftime("%Y-%m-%d"),
        "REFERÊNCIA OPUS": ref_opus.strip(),
        "OBJETIVO (ADICIONAR POSSÍVEL CONTATO)": objetivo.strip(),
        "DATA DA VISTORIA": pd.to_datetime(data_vist).strftime("%Y-%m-%d") if data_vist else "",
        "VT EXECUTADA POR": vt_exec.strip(),
        "STATUS - ATUALIZAÇÃO SEMANAL": status_sem.strip(),
        "DATA/PREVISÃO DE CONCLUSÃO": "",
        "MEIO DE RESPOSTA DA SOLICITAÇÃO": meio_resp.strip(),
        "DATA DA RESPOSTA A SOLICITAÇÃO": "",
        "Nº OPUS DA VISTORIA (SE FOR O CASO)": opus_vist.strip(),
        "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO": int(qt_total) if pd.notna(qt_total) else "",
        "QUANTIDADE DE DIAS PARA EXECUÇÃO": int(qt_exec) if pd.notna(qt_exec) else "",
        "OBSERVAÇÕES": obs.strip(),
    }
    return row, (len(erros) == 0)


def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    try:
        base_preview = read_df(TAB_SOLICITACOES)
    except Exception:
        base_preview = pd.DataFrame()

    row, ok = _input_row(base_preview)

    if st.button("💾 Salvar na aba ACOMPANHAMENTO VISTORIAS", disabled=not ok, type="primary"):
        try:
            append_row(TAB_SOLICITACOES, row)
            st.success("Registro salvo na aba ACOMPANHAMENTO VISTORIAS.")
            st.rerun()
        except Exception as e:
            st.error(f"Falha ao salvar: {e}")

    st.divider()
    st.subheader("Visualização rápida da base (últimos 30)")
    if not base_preview.empty:
        cols_show = [c for c in [
            "OBJETO DE VISTORIA","OM APOIADA","Diretoria Responsável",
            "Classificação de Urgência","Situação","DATA DA SOLICITAÇÃO"
        ] if c in base_preview.columns]
        df_show = base_preview.copy()
        if "DATA DA SOLICITAÇÃO" in df_show.columns:
            df_show["_dt"] = pd.to_datetime(df_show["DATA DA SOLICITAÇÃO"], errors="coerce")
            df_show = df_show.sort_values("_dt", ascending=False).drop(columns=["_dt"])
        st.dataframe(df_show[cols_show].head(30), use_container_width=True, height=320)
    else:
        st.caption("Ainda não há registros.")
