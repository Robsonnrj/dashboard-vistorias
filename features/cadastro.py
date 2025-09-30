# features/cadastro.py
# -*- coding: utf-8 -*-
import streamlit as st
from datetime import datetime, date
import pandas as pd

from core.data_loader import read_df, append_row


# ------------------------ helpers ------------------------ #
def _norm(s: str) -> str:
    s = str(s or "")
    return (
        s.strip()
         .lower()
         .replace("á", "a").replace("à", "a").replace("ã", "a").replace("â", "a")
         .replace("é", "e").replace("ê", "e")
         .replace("í", "i")
         .replace("ó", "o").replace("õ", "o").replace("ô", "o")
         .replace("ú", "u")
         .replace("ç", "c")
    )

def _pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Retorna o nome da coluna existente no DF que melhor casa com a lista de candidatos."""
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    # match exato “normalizado”
    for want in candidates:
        for c in cols:
            if _norm(c) == _norm(want):
                return c
    # match por “contém”
    for want in candidates:
        w = _norm(want)
        for c in cols:
            if w in _norm(c):
                return c
    return None

@st.cache_data(ttl=300, show_spinner=False)
def _load_oms_map(tab_valid: str | None, tab_base: str) -> pd.DataFrame:
    """
    Carrega um dataframe com colunas: om_sigla, om_nome, diretoria
    1) Tenta aba de validação (Validacao_de_Dados);
    2) Se não houver, deduz a partir da aba base (ACOMPANHAMENTO VISTORIAS).
    """
    def _from_validation(dfv: pd.DataFrame) -> pd.DataFrame:
        c_sigla = _pick_col(dfv, ["OM", "Sigla"])
        c_nome  = _pick_col(dfv, ["Organização Militar", "OM Nome", "Nome"])
        c_dir   = _pick_col(dfv, ["Diretoria Responsável", "Diretoria"])
        out = dfv[[c_sigla, c_nome, c_dir]].dropna(how="all").copy()
        out.columns = ["om_sigla", "om_nome", "diretoria"]
        out = out.dropna(subset=["om_sigla"]).drop_duplicates(subset=["om_sigla"])
        return out

    def _from_base(dfb: pd.DataFrame) -> pd.DataFrame:
        c_sigla = _pick_col(dfb, ["OM", "OM APOIADA", "OM APOIADORA"])
        c_dir   = _pick_col(dfb, ["Diretoria Responsável", "Diretoria"])
        if not c_sigla:
            return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])
        tmp = dfb[[c_sigla] + ([c_dir] if c_dir else [])].copy()
        tmp.columns = ["om_sigla"] + (["diretoria"] if c_dir else [])
        tmp["om_nome"] = ""  # não temos o nome completo aqui
        tmp = tmp.dropna(subset=["om_sigla"]).drop_duplicates(subset=["om_sigla"])
        # garante coluna diretoria
        if "diretoria" not in tmp.columns:
            tmp["diretoria"] = ""
        return tmp[["om_sigla", "om_nome", "diretoria"]]

    # 1) tenta validação
    if tab_valid:
        try:
            dfv = read_df(tab_valid)
            if not dfv.empty:
                return _from_validation(dfv)
        except Exception:
            pass

    # 2) fallback a partir da aba base
    dfb = read_df(tab_base)
    if not dfb.empty:
        return _from_base(dfb)

    return pd.DataFrame(columns=["om_sigla", "om_nome", "diretoria"])


def _input_row(oms_df: pd.DataFrame) -> tuple[dict, bool]:
    """Coleta os campos do formulário e devolve (linha, valido)."""
    st.subheader("📥 Nova solicitação de vistoria")
    col1, col2 = st.columns(2)

    # --- OM (autocomplete) ---
    # Opções formatadas: "SIGLA — Nome (Diretoria)"
    options = []
    display_to_sigla = {}
    for _, r in oms_df.iterrows():
        sig = str(r.get("om_sigla", "") or "").strip()
        nom = str(r.get("om_nome", "") or "").strip()
        dir_ = str(r.get("diretoria", "") or "").strip()
        if not sig:
            continue
        label = sig
        if nom:
            label += f" — {nom}"
        if dir_:
            label += f"  ({dir_})"
        options.append(label)
        display_to_sigla[label] = sig

    with col1:
        om_label = st.selectbox(
            "Organização Militar (OM)",
            options,
            index=0 if options else None,
            placeholder="Digite para buscar…",
        )
        om_solicitante = display_to_sigla.get(om_label, "")

    # Diretoria automática (com possibilidade de ajuste manual, se quiser)
    dir_auto = ""
    if om_solicitante:
        hit = oms_df.loc[oms_df["om_sigla"] == om_solicitante]
        if not hit.empty:
            dir_auto = str(hit.iloc[0].get("diretoria", "") or "")

    with col1:
        diretoria = st.text_input("Diretoria responsável", value=dir_auto)

    with col2:
        local = st.text_input("Local / instalação")
        urgencia = st.selectbox(
            "Urgência",
            ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"],
            index=0,
        )
        data_limite = st.date_input("Data limite (se houver)", value=None)

    with col2:
        tipo_vistoria = st.selectbox(
            "Tipo de vistoria",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
        )

    motivo = st.text_area("Motivo / justificativa (NAOM)", height=120)

    # Validações simples
    erros = []
    if not om_solicitante.strip():
        erros.append("Informe a **OM**.")
    if not local.strip():
        erros.append("Informe o **local/instalação**.")
    if not motivo.strip():
        erros.append("Descreva o **motivo/justificativa**.")

    if erros:
        # usa container para não envolver com @st.cache_data wrappers
        st.markdown("⚠️ " + "<br>• ".join([""] + erros), unsafe_allow_html=True)

    # Normaliza a data_limite para string ISO (ou vazio)
    if isinstance(data_limite, (date,)):
        data_limite_str = data_limite.strftime("%Y-%m-%d")
    else:
        data_limite_str = ""

    row = {
        "numero": "",  # será atribuído no salvar
        "data_solicitacao": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "om_solicitante": om_solicitante.strip(),
        "diretoria": diretoria.strip(),
        "tipo_vistoria": tipo_vistoria,
        "local": local.strip(),
        "urgencia": urgencia,
        "data_limite": data_limite_str,
        "motivo": motivo.strip(),
        "status_atual": "SOLICITADA",
    }
    return row, (len(erros) == 0)


def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    # Abas configuradas na sidebar
    tabs_map = st.session_state.get("tabs_map", {})
    tab_base  = tabs_map.get("solicitacoes", "ACOMPANHAMENTO VISTORIAS")
    tab_valid = tabs_map.get("validacao", "Validacao_de_Dados")

    # Carrega dados existentes (para gerar número sequencial)
    try:
        df_existente = read_df(tab_base)
    except Exception:
        df_existente = pd.DataFrame()

    # Carrega mapa de OMs (sigla -> diretoria/nome) com cache 5min
    oms_df = _load_oms_map(tab_valid, tab_base)

    # Formulário
    row, ok = _input_row(oms_df)

    # Salvar
    if st.button("💾 Salvar solicitação", type="primary", disabled=not ok):
        try:
            # Gera número sequencial simples
            proximo = 1
            if not df_existente.empty and "numero" in df_existente.columns:
                nums = pd.to_numeric(df_existente["numero"], errors="coerce").dropna()
                if not nums.empty:
                    proximo = int(nums.max()) + 1
            row["numero"] = str(proximo)

            # -------- MAPEAMENTO para o cabeçalho real da aba -------- #
            # Lê de novo só para pegar as colunas reais atuais
            df_cols = list(read_df(tab_base).columns)

            def col_real(cands: list[str], default: str) -> str:
                for c in cands:
                    if c in df_cols:
                        return c
                    # match tolerante
                    for dc in df_cols:
                        if _norm(dc) == _norm(c):
                            return dc
                # se não existir, volta o nome default (será ignorado se a aba ainda não possuir)
                return default

            # candidatos para cada campo
            map_out = {
                col_real(["NÚMERO", "numero", "Nº"], "numero"): row["numero"],
                col_real(["DATA DA SOLICITAÇÃO", "data_solicitacao", "DATA"], "data_solicitacao"): row["data_solicitacao"],
                col_real(["OM", "OM APOIADA", "OM APOIADORA"], "om_solicitante"): row["om_solicitante"],
                col_real(["Diretoria Responsável", "Diretoria"], "diretoria"): row["diretoria"],
                col_real(["Tipo de Vistoria", "tipo_vistoria"], "tipo_vistoria"): row["tipo_vistoria"],
                col_real(["Local", "Local / instalação", "INSTALAÇÃO"], "local"): row["local"],
                col_real(["Classificação de Urgência", "Urgência", "urgencia"], "urgencia"): row["urgencia"],
                col_real(["DATA LIMITE", "DATA_LIMITE", "data_limite"], "data_limite"): row["data_limite"],
                col_real(["Motivo", "OBJETIVO", "Justificativa", "motivo"], "motivo"): row["motivo"],
                col_real(["Situação", "STATUS - ATUALIZAÇÃO SEMANAL", "status_atual"], "status_atual"): row["status_atual"],
            }

            # filtra só chaves que de fato existem na aba alvo
            row_to_save = {k: v for k, v in map_out.items() if k in df_cols}

            # se por algum motivo nada casou, cai para salvar o dicionário “cru”
            if not row_to_save:
                row_to_save = row

            append_row(tab_base, row_to_save)
            st.success(f"Solicitação **#{row['numero']}** cadastrada com sucesso!")
            st.rerun()
        except Exception as e:
            st.error(f"Falha ao salvar: {e}")

    # Lista das últimas solicitações
    st.divider()
    st.subheader("📄 Últimas solicitações")

    if not df_existente.empty:
        # tenta ordenar por alguma coluna de data
        c_data = None
        for c in df_existente.columns:
            if "data" in c.lower():
                c_data = c
                break
        if c_data:
            df_existente[c_data] = pd.to_datetime(df_existente[c_data], errors="coerce")
            df_existente = df_existente.sort_values(c_data, ascending=False)

        st.dataframe(df_existente.head(50), use_container_width=True, height=360)
    else:
        st.caption("Ainda não há registros.")
