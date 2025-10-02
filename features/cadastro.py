# features/cadastro.py
# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime
import pandas as pd

from core.data_loader import append_row, read_df

# =========================================================
# Helpers
# =========================================================
def _norm(s: str) -> str:
    return str(s).strip().lower()

def _find_col(df: pd.DataFrame, *cands: str) -> str:
    """Procura coluna por nome (exato ou contém), ignorando maiúsculas/acentos simples."""
    cols = {_norm(c): c for c in df.columns}
    for c in cands:
        k = _norm(c)
        if k in cols:
            return cols[k]
    for want in cands:
        w = _norm(want)
        for k, orig in cols.items():
            if w in k:
                return orig
    raise KeyError(f"Não achei colunas {cands} em {list(df.columns)}")

# =========================================================
# Leitura da aba de validação (robusta ao cabeçalho)
# =========================================================
def _load_oms_validadas() -> pd.DataFrame:
    """
    Lê a aba 'Validacao_de_Dados' e retorna colunas padronizadas:
    om_sigla, om_nome, diretoria.
    """
    # 1) lê SEM header=... (alguns conectores não aceitam)
    df = read_df("Validacao_de_Dados")

    # 2) Se a primeira linha parece ser o cabeçalho, promove-a a header
    def _maybe_promote_header(df0: pd.DataFrame) -> pd.DataFrame:
        if df0.empty:
            return df0
        cols_lower = [str(c).strip().lower() for c in df0.columns]
        row0_lower = [str(x).strip().lower() for x in df0.iloc[0].tolist()]
        has_expected_in_cols = any(
            ("diretoria" in c) or ("organização" in c) or ("organizacao" in c) or (c == "om")
            for c in cols_lower
        )
        has_expected_in_row0 = any(
            ("diretoria" in x) or ("organização" in x) or ("organizacao" in x) or (x == "om")
            for x in row0_lower
        )
        if not has_expected_in_cols and has_expected_in_row0:
            df1 = df0.copy()
            df1.columns = df1.iloc[0]
            df1 = df1.iloc[1:].reset_index(drop=True)
            return df1
        return df0

    df = _maybe_promote_header(df)

    # 3) Detecta colunas
    col_sigla = _find_col(df, "om", "sigla")
    col_nome  = _find_col(df, "organização militar", "organizacao militar", "om nome", "nome")
    col_dir   = _find_col(df, "diretoria responsável", "diretoria responsavel", "diretoria")

    out = df.rename(columns={
        col_sigla: "om_sigla",
        col_nome : "om_nome",
        col_dir  : "diretoria",
    })[["om_sigla", "om_nome", "diretoria"]].copy()

    # 4) Limpeza
    for c in ["om_sigla", "om_nome", "diretoria"]:
        out[c] = out[c].astype(str).str.strip()

    out = out[(out["om_sigla"] != "") & (out["diretoria"] != "")]
    out = out.drop_duplicates(subset=["om_sigla", "diretoria"])
    return out

def _build_om_options(oms_df: pd.DataFrame):
    """
    Retorna (options_display, disp_to_sigla, sigla_to_dir, disp_to_dir)
    - disp_to_dir permite mapear DIRETAMENTE pela string exibida no select.
    """
    options_display, disp_to_sigla, sigla_to_dir, disp_to_dir = [], {}, {}, {}
    for _, r in oms_df.iterrows():
        sig  = str(r["om_sigla"]).strip()
        nom  = str(r["om_nome"]).strip()
        dire = str(r["diretoria"]).strip()
        display = f"{sig} — {nom}" if nom else sig
        options_display.append(display)
        disp_to_sigla[display] = sig
        sigla_to_dir[sig] = dire
        disp_to_dir[display] = dire
    options_display.append("Outra / não listada…")
    disp_to_sigla["Outra / não listada…"] = ""
    disp_to_dir["Outra / não listada…"] = ""
    return options_display, disp_to_sigla, sigla_to_dir, disp_to_dir

# =========================================================
# UI
# =========================================================
def _input_row():
    st.subheader("📥 Nova solicitação de vistoria")
    oms_df = _load_oms_validadas()
    options, disp2sig, sig2dir, disp2dir = _build_om_options(oms_df)
    st.caption(f"{len(oms_df)} Organizações Militares carregadas da base de validação")

    st.session_state.setdefault("diretoria_auto", "")

    col1, col2 = st.columns(2)

    with col1:
        # OM (display)
        om_display = st.selectbox(
            "OM solicitante (sigla) *",
            options=options,
            index=None,
            placeholder="Selecione ou digite…",
            key="om_choice",
        )
        om_sigla = disp2sig.get(om_display or "", "")

        # Diretoria
        if om_display == "Outra / não listada…":
            om_sigla = st.text_input("Sigla da OM (manual)", key="om_sigla_out")
            diretoria = st.text_input("Diretoria responsável (manual)", key="diretoria_manual")
            st.session_state["diretoria_auto"] = ""
        else:
            # 1º tenta via display (confiável); se vazio, cai para sigla
            default_dir = disp2dir.get(om_display or "", "") or sig2dir.get(om_sigla or "", "")
            st.session_state["diretoria_auto"] = default_dir
            st.text_input("Diretoria responsável *", key="diretoria_auto", disabled=True)
            diretoria = st.session_state["diretoria_auto"]

        tipo_vistoria = st.selectbox(
            "Tipo de vistoria *",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
        )

    with col2:
        local = st.text_input("Local / instalação *")
        urgencia = st.selectbox("Urgência *", ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"], index=0)
        data_limite = st.date_input("Data limite (opcional)", value=None)

    motivo = st.text_area("Motivo / justificativa (NAOM) *", height=120)

    # Validação
    erros = []
    if not (om_sigla or "").strip():
        erros.append("Informe a **OM**.")
    if om_display == "Outra / não listada…":
        if not (diretoria or "").strip():
            erros.append("Informe a **diretoria** (manual).")
    else:
        if not (diretoria or "").strip():
            erros.append("Diretoria não encontrada para a OM selecionada.")
    if not (local or "").strip():
        erros.append("Informe o **local/instalação**.")
    if not (motivo or "").strip():
        erros.append("Descreva o **motivo/justificativa**.")

    if erros:
        st.warning("Campos obrigatórios não preenchidos:\n\n• " + "\n• ".join(erros))

    row = {
        "numero": "",
        "data_solicitacao": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "om_solicitante": (om_sigla or "").strip(),
        "diretoria": (diretoria or "").strip(),
        "tipo_vistoria": tipo_vistoria,
        "local": (local or "").strip(),
        "urgencia": urgencia,
        "data_limite": data_limite.strftime("%Y-%m-%d") if data_limite else "",
        "motivo": (motivo or "").strip(),
        "status_atual": "SOLICITADA",
    }
    return row, (len(erros) == 0)

def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")
    tabs_map = st.session_state.get("tabs_map", {})
    tab_solic = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")

    try:
        df_existente = read_df(tab_solic)
    except Exception:
        df_existente = pd.DataFrame()

    row, ok = _input_row()

    if st.button("💾 Salvar solicitação", type="primary", disabled=not ok):
        try:
            proximo = 1
            if not df_existente.empty and "numero" in df_existente.columns:
                nums = pd.to_numeric(df_existente["numero"], errors="coerce").dropna()
                if not nums.empty:
                    proximo = int(nums.max()) + 1
            row["numero"] = str(proximo)
            append_row(tab_solic, row)
            st.success(f"Solicitação **#{row['numero']}** cadastrada com sucesso!")
            st.rerun()
        except Exception as e:
            st.error(f"Falha ao salvar: {e}")

    st.divider()
    st.subheader("📄 Últimas solicitações")
    if not df_existente.empty:
        c_data = next((c for c in df_existente.columns if "data" in c.lower()), None)
        if c_data:
            df_existente[c_data] = pd.to_datetime(df_existente[c_data], errors="coerce")
            df_existente = df_existente.sort_values(c_data, ascending=False)
        st.dataframe(df_existente.head(50), use_container_width=True, height=360)
    else:
        st.caption("Ainda não há registros.")
