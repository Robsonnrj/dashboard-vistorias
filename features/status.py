# features/status.py
from __future__ import annotations

import streamlit as st
import pandas as pd
from datetime import datetime
from core.config import TAB_SOLICITACOES, TAB_AUDIT, TAB_VALIDACAO
from core.data_loader import read_df, overwrite_tab_from_df

URGENCIAS = ["Não Prioridade", "Prioridade", "Urgente"]
SITUACOES = ["Não Atendida", "Em andamento", "Finalizada"]

# -----------------------------
# Helpers
# -----------------------------
def _clean(x) -> str:
    return "" if pd.isna(x) else str(x).strip()

def _date_or(x, default: date) -> date:
    """Converte para date; se inválido/NaT, retorna 'default'."""
    d = pd.to_datetime(x, errors="coerce")
    return d.date() if pd.notna(d) else default

AUDIT_FIELDS = [
    "objeto de vistoria", "om apoiada", "diretoria responsável",
    "classificação de urgência", "situação",
    "data da solicitação", "data da vistoria",
    "status - atualização semanal", "observações",
]

def _now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _append_audit(numero: str, changes: dict):
    """
    Adiciona 1 linha por campo alterado na aba de auditoria (TAB_AUDIT).
    Espera: changes = {nome_campo: (valor_antigo, valor_novo), ...}
    """
    if not changes:
        return
    try:
        hist = read_df(TAB_AUDIT)
    except Exception:
        hist = pd.DataFrame()

    rows = []
    for campo, (antes, depois) in changes.items():
        rows.append({
            "numero": str(numero),
            "ts": _now_ts(),
            "campo": campo,
            "antes": "" if pd.isna(antes) else str(antes),
            "depois": "" if pd.isna(depois) else str(depois),
        })

    new_hist = (pd.concat([hist, pd.DataFrame(rows)], ignore_index=True)
                if not hist.empty else pd.DataFrame(rows))

    # salva de volta (mantendo cabeçalho)
    overwrite_tab_from_df(TAB_AUDIT, new_hist, keep_header=True)

def _collect_changes(orig_row: pd.Series, new_row: pd.Series, cols_map: dict) -> dict:
    """
    Compara valores antigos x novos e retorna um dict
    {nome_campo_humano: (antes, depois)} apenas para os que mudaram.
    """
    changes = {}
    # mapeia nomes humanos -> nomes de coluna reais no DF
    for human in AUDIT_FIELDS:
        col_real = cols_map.get(human)
        if not col_real:
            continue
        old = orig_row.get(col_real, "")
        new = new_row.get(col_real, "")
        # normaliza para comparar
        old_s = "" if pd.isna(old) else str(old).strip()
        new_s = "" if pd.isna(new) else str(new).strip()
        if old_s != new_s:
            changes[human] = (old_s, new_s)
    return changes

def _load_oms_map():
    """Monta opções de OM e mapas display->sigla e sigla->diretoria."""
    options, disp2sig, sig2dir = [], {}, {}
    # Prioriza a aba de validação; se falhar, tenta a de solicitações
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

def _load_audit_trail(numero: str) -> pd.DataFrame:
    """Filtra trilha de auditoria pelo número do registro."""
    try:
        hist = read_df(TAB_AUDIT)
    except Exception:
        return pd.DataFrame()
    if hist.empty or "numero" not in hist.columns:
        return pd.DataFrame()
    return hist[hist["numero"].astype(str) == str(numero)].sort_values("ts", ascending=False)

# -----------------------------
# Página
# -----------------------------
def page():
    st.header("🔁 VIS-003 — Controle de Status e Auditoria")

    try:
        df = read_df(TAB_SOLICITACOES)
    except Exception as e:
        st.error(f"Falha ao ler a base: {e}")
        return

    if df.empty:
        st.info("Sem registros na aba ACOMPANHAMENTO VISTORIAS.")
        return

    cols = {c.lower(): c for c in df.columns}
    c_obj = cols.get("objeto de vistoria")
    c_om  = cols.get("om apoiada") or cols.get("om")
    c_dir = cols.get("diretoria responsável") or cols.get("diretoria")
    c_sit = cols.get("situação") or cols.get("situacao")
    c_urg = cols.get("classificação de urgência") or cols.get("classificacao urgencia")
    c_dtS = cols.get("data da solicitação") or cols.get("data da solicitacao")
    c_dtV = cols.get("data da vistoria")
    c_stw = cols.get("status - atualização semanal") or cols.get("status - atualizacao semanal")
    c_obs = cols.get("observações") or cols.get("observacoes")

    st.subheader("Selecione o registro para editar")
    show_cols = [x for x in [c_obj, c_om, c_dir, c_sit, c_dtS] if x in df.columns]
    show = df[show_cols].copy() if show_cols else df.copy()
    show = show.reset_index().rename(columns={"index": "linha"})

    idx = st.selectbox(
        "Registro",
        options=show["linha"].tolist(),
        format_func=lambda i: " | ".join(
            [_clean(x) for x in show.loc[show["linha"] == i, show_cols].iloc[0].tolist()]
        ),
    )
    if idx is None:
        return

    options, disp2sig, sig2dir = _load_oms_map()
    reg = df.loc[idx].copy()

    form_uid = f"r{idx}"  # muda a cada registro selecionado

    with st.form(f"frm_status_{form_uid}"):
        objeto = st.text_input(
            "OBJETO DE VISTORIA *",
            value=_clean(reg.get(c_obj, "")),
            key=f"obj_{form_uid}",
        )
    
        # default do select de OM (pelo valor salvo)
        om_default = next((k for k, v in disp2sig.items() if v == _clean(reg.get(c_om, ""))), None)
        default_index = options.index(om_default) if (om_default in options) else None
    
        om_display = st.selectbox(
            "OM APOIADA *",
            options=options,
            index=default_index,
            placeholder="Selecione…",
            key=f"om_{form_uid}",
        )
        om_sigla = disp2sig.get(om_display or "", _clean(reg.get(c_om, "")))
    
        diretoria = st.text_input(
            "Diretoria Responsável *",
            value=sig2dir.get(om_sigla, _clean(reg.get(c_dir, ""))),
            key=f"dir_{form_uid}",
        )
    
        urg = st.selectbox(
            "Classificação de Urgência",
            URGENCIAS,
            index=URGENCIAS.index(_clean(reg.get(c_urg, URGENCIAS[0])))
                  if _clean(reg.get(c_urg, "")) in URGENCIAS else 0,
            key=f"urg_{form_uid}",
        )
    
        sit = st.selectbox(
            "Situação",
            SITUACOES,
            index=SITUACOES.index(_clean(reg.get(c_sit, SITUACOES[0])))
                  if _clean(reg.get(c_sit, "")) in SITUACOES else 0,
            key=f"sit_{form_uid}",
        )
    
        dt_sol = st.date_input(
            "DATA DA SOLICITAÇÃO",
            value=_date_or(reg.get(c_dtS, ""), date.today()),
            key=f"dts_{form_uid}",
        )
        dt_vis = st.date_input(
            "DATA DA VISTORIA",
            value=_date_or(reg.get(c_dtV, ""), date.today()),
            key=f"dtv_{form_uid}",
        )
    
        stw = st.text_input(
            "STATUS - ATUALIZAÇÃO SEMANAL",
            value=_clean(reg.get(c_stw, "")),
            key=f"stw_{form_uid}",
        )
        obs = st.text_area(
            "OBSERVAÇÕES",
            value=_clean(reg.get(c_obs, "")),
            height=100,
            key=f"obs_{form_uid}",
        )
    
        salvar = st.form_submit_button("💾 Atualizar registro", type="primary", key=f"save_{form_uid}")
    
    
          if not salvar:
        return

    # valida obrigatórios
    faltando = []
    if not objeto:    faltando.append("OBJETO DE VISTORIA")
    if not om_sigla:  faltando.append("OM APOIADA")
    if not diretoria: faltando.append("Diretoria Responsável")
    if faltando:
        st.error("Preencha os campos obrigatórios: " + ", ".join(faltando))
        return

    # ----- prepara atualização do DF principal -----
    # monta um "row novo" só com os campos relevantes
    new_vals = {}
    if c_obj: new_vals[c_obj] = objeto
    if c_om:  new_vals[c_om]  = om_sigla
    if c_dir: new_vals[c_dir] = diretoria
    if c_urg: new_vals[c_urg] = urg
    if c_sit: new_vals[c_sit] = sit
    if c_dtS: new_vals[c_dtS] = pd.to_datetime(dt_sol).strftime("%Y-%m-%d")
    if c_dtV: new_vals[c_dtV] = pd.to_datetime(dt_vis).strftime("%Y-%m-%d") if dt_vis else ""
    if c_stw: new_vals[c_stw] = stw
    if c_obs: new_vals[c_obs] = obs

    # localiza a linha pelo "numero" se existir; senão usa o índice selecionado
    cols_all = {c.lower(): c for c in df.columns}
    c_num = cols_all.get("numero")
    if c_num and c_num in df.columns and str(df.at[idx, c_num]).strip():
        numero = str(df.at[idx, c_num]).strip()
        ixs = df.index[df[c_num].astype(str).str.strip() == numero]
        if len(ixs) > 0:
            idx_target = ixs[0]
        else:
            idx_target = idx  # fallback
    else:
        numero = str(df.index.get_loc(idx))  # sem coluna numero; usa posição como id
        idx_target = idx

    # coleta changes (diff) para auditoria ANTES de escrever
    cols_human_to_real = {
        "objeto de vistoria": c_obj,
        "om apoiada": c_om,
        "diretoria responsável": c_dir,
        "classificação de urgência": c_urg,
        "situação": c_sit,
        "data da solicitação": c_dtS,
        "data da vistoria": c_dtV,
        "status - atualização semanal": c_stw,
        "observações": c_obs,
    }
    orig_row = df.loc[idx_target].copy()
    # cria uma série com os novos valores para comparar
    new_row = orig_row.copy()
    for k, v in new_vals.items():
        new_row[k] = v
    changes = _collect_changes(orig_row, new_row, cols_human_to_real)

    # aplica atualização na linha alvo
    for k, v in new_vals.items():
        df.at[idx_target, k] = v

    # grava DF de volta e registra auditoria
    try:
        overwrite_tab_from_df(TAB_SOLICITACOES, df, keep_header=True)
        _append_audit(numero, changes)
        st.success("Registro atualizado com sucesso.")
    except Exception as e:
        st.error(f"Falha ao salvar: {e}")

        st.success("Registro atualizado com sucesso.")
    except Exception as e:
        st.error(f"Falha ao salvar: {e}")

    # Trilha de Auditoria
    st.subheader("📝 Trilha de Auditoria")
    
    # evita TypeError com pd.NA usando _clean() antes do "or"
    num = _clean(reg.get("numero", "")) or _clean(reg.get("Numero", ""))
    
    hist_df = _load_audit_trail(num)
    if hist_df.empty:
        st.info("Sem registros de auditoria para este número.")
    else:
        st.dataframe(hist_df, use_container_width=True, height=300)
