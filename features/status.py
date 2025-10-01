# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import date
from core.config import TAB_SOLICITACOES, TAB_VALIDACAO
from core.data_loader import read_df, overwrite_tab_from_df


URGENCIAS = ["Não Prioridade", "Prioridade", "Urgente"]
SITUACOES = ["Não Atendida", "Em andamento", "Finalizada"]

REQ_COLS = [
    "OBJETO DE VISTORIA",
    "OM APOIADA",
    "Diretoria Responsável",
    "Classificação de Urgência",
    "Situação",
    "DATA DA SOLICITAÇÃO",
    "DATA DA VISTORIA",
    "STATUS - ATUALIZAÇÃO SEMANAL",
    "OBSERVAÇÕES",
]
def _clean(x): return "" if pd.isna(x) else str(x).strip()

def _load_oms_map():
    """Retorna (labels, display->sigla, sigla->diretoria)."""
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
        tmp["dir"]  = df[c_dir].map(_clean)
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

def page():
    st.header("🔁 VIS-003 — Controle de Status e Auditoria")

    # Base completa
    try:
        df = read_df(TAB_SOLICITACOES)
    except Exception as e:
        st.error(f"Falha ao ler a base: {e}")
        return

    if df.empty:
        st.info("Sem registros na aba ACOMPANHAMENTO VISTORIAS.")
        return

    # Normalização básica de colunas esperadas
    cols = {c.lower(): c for c in df.columns}
    c_obj = cols.get("objeto de vistoria")
    c_om  = cols.get("om apoiada") or cols.get("om")
    c_dir = cols.get("diretoria responsável") or cols.get("diretoria")
    c_sit = cols.get("situação") or cols.get("situacao")
    c_urg = cols.get("classificação de urgência") or cols.get("classificacao de urgencia")
    c_dtS = cols.get("data da solicitação") or cols.get("data da solicitacao")
    c_dtV = cols.get("data da vistoria")
    c_stw = cols.get("status - atualização semanal") or cols.get("status - atualizacao semanal")
    c_obs = cols.get("observações") or cols.get("observacoes")

    # Linha alvo (seleciona por índice para evitar ambiguidade)
    st.subheader("Selecione o registro")
    show_cols = [x for x in [c_obj, c_om, c_dir, c_sit, c_dtS] if x in df.columns]
    show = df[show_cols].copy() if show_cols else df.copy()
    show = show.reset_index().rename(columns={"index": "linha"})
    idx = st.selectbox(
        "Registro",
        options=show["linha"].tolist(),
        format_func=lambda i: " | ".join([_clean(x) for x in show.loc[show["linha"]==i, show_cols].iloc[0].tolist()]),
    )

    if idx is None:
        return

    # Mapeamento OM->Diretoria para atualização automática
    options, disp2sig, sig2dir = _load_oms_map()
    reg = df.loc[idx].copy()

    with st.form("frm_status"):
        # Campos essenciais
        objeto = st.text_input("OBJETO DE VISTORIA *", value=_clean(reg.get(c_obj, "")))
        om_label_default = next((k for k,v in disp2sig.items() if v == _clean(reg.get(c_om,""))), None)
        om_display = st.selectbox("OM APOIADA *", options, index=None, placeholder="Selecione…", key="om_disp", 
                                  index=options.index(om_label_default) if om_label_default in options else None)
        om_sigla = disp2sig.get(om_display or "", _clean(reg.get(c_om, "")))
        diretoria = st.text_input("Diretoria Responsável *", value=sig2dir.get(om_sigla, _clean(reg.get(c_dir, ""))))

        urg = st.selectbox("Classificação de Urgência", URGENCIAS,
                           index=URGENCIAS.index(_clean(reg.get(c_urg, URGENCIAS[0]))) if _clean(reg.get(c_urg,"")) in URGENCIAS else 0)
        sit = st.selectbox("Situação", SITUACOES,
                           index=SITUACOES.index(_clean(reg.get(c_sit, SITUACOES[0]))) if _clean(reg.get(c_sit,"")) in SITUACOES else 0)

        dt_sol = st.date_input("DATA DA SOLICITAÇÃO", value=pd.to_datetime(reg.get(c_dtS,""), errors="coerce").date() if c_dtS else date.today())
        dt_vis = st.date_input("DATA DA VISTORIA", value=pd.to_datetime(reg.get(c_dtV,""), errors="coerce").date() if c_dtV else None)

        stw   = st.text_input("STATUS - ATUALIZAÇÃO SEMANAL", value=_clean(reg.get(c_stw, "")))
        obs   = st.text_area("OBSERVAÇÕES", value=_clean(reg.get(c_obs, "")), height=100)

        salvar = st.form_submit_button("💾 Atualizar registro", type="primary")

    if not salvar:
        return

    # Validação
    faltando = []
    if not objeto:    faltando.append("OBJETO DE VISTORIA")
    if not om_sigla:  faltando.append("OM APOIADA")
    if not diretoria: faltando.append("Diretoria Responsável")
    if faltando:
        st.error("Preencha os campos obrigatórios: " + ", ".join(faltando))
        return

    # Aplica alterações no DataFrame e grava a aba toda
    if c_obj: df.at[idx, c_obj] = objeto
    if c_om:  df.at[idx, c_om]  = om_sigla
    if c_dir: df.at[idx, c_dir] = diretoria
    if c_urg: df.at[idx, c_urg] = urg
    if c_sit: df.at[idx, c_sit] = sit
    if c_dtS: df.at[idx, c_dtS] = pd.to_datetime(dt_sol).strftime("%Y-%m-%d")
    if c_dtV: df.at[idx, c_dtV] = pd.to_datetime(dt_vis).strftime("%Y-%m-%d") if dt_vis else ""
    if c_stw: df.at[idx, c_stw] = stw
    if c_obs: df.at[idx, c_obs] = obs

    try:
        overwrite_tab_from_df(TAB_SOLICITACOES, df, keep_header=True)
        st.success("Registro atualizado com sucesso.")
        st.rerun()
    except Exception as e:
        st.error(f"Falha ao salvar: {e}")
