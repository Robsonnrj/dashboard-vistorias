# features/cadastro.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import streamlit as st
from datetime import datetime, date
import pandas as pd

from core.data_loader import append_row, read_df
from core.config import TAB_VALIDACAO


# =========================================================
# Helpers
# =========================================================
def _clean(x) -> str:
    """Normaliza valores vindos do Sheets."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.upper() in {"#N/A", "N/A", "NA", "NAN", "NONE", "-", ""}:
        return ""
    return s


def _build_om_catalog() -> tuple[list[str], dict[str, str], dict[str, str]]:
    """
    Lê a aba de validação e monta:
      - options: rótulos para o select (ex.: "1º BPE — 1º Batalhão ...")
      - disp2sig: rótulo -> sigla
      - sig2dir : sigla  -> diretoria
    """
    options: list[str] = []
    disp2sig: dict[str, str] = {}
    sig2dir: dict[str, str] = {}
    
    try:
        # Lê os dados sem processar header ainda
        df_raw = read_df(TAB_VALIDACAO)
        
        # Procura pela linha que contém os cabeçalhos reais
        header_row = None
        for idx in range(min(5, len(df_raw))):  # Procura nas primeiras 5 linhas
            row_values = df_raw.iloc[idx].astype(str).str.upper().tolist()
            if "OM" in row_values and any("ORGANIZA" in str(v) for v in row_values):
                header_row = idx
                break
        
        if header_row is None:
            st.sidebar.warning("⚠️ Não foi possível localizar o cabeçalho na aba de validação")
            return options, disp2sig, sig2dir
        
        # Pega os dados a partir da linha seguinte ao header
        df = df_raw.iloc[header_row + 1:].reset_index(drop=True)
        
        # Define nomes únicos para as colunas baseado no header encontrado
        header_values = df_raw.iloc[header_row].tolist()
        
        # Cria nomes únicos para colunas duplicadas
        col_names = []
        col_counts = {}
        for val in header_values:
            val_str = _clean(str(val))
            if not val_str or val_str == "<NA>":
                val_str = f"col_{len(col_names)}"
            
            if val_str in col_counts:
                col_counts[val_str] += 1
                val_str = f"{val_str}_{col_counts[val_str]}"
            else:
                col_counts[val_str] = 0
            
            col_names.append(val_str)
        
        df.columns = col_names
        
        # Debug: mostra informações
        st.sidebar.write("🔍 DEBUG: Estrutura após ajuste")
        st.sidebar.write(f"Total de linhas: {len(df)}")
        st.sidebar.write(f"Colunas: {col_names[:6]}")  # Mostra apenas as primeiras 6
        
        # Mostra preview dos dados (convertendo para dict para evitar erro de colunas duplicadas)
        preview_data = {
            "col": col_names[:12],
            "val_linha_0": df.iloc[0].tolist()[:12] if len(df) > 0 else [],
            "val_linha_1": df.iloc[1].tolist()[:12] if len(df) > 1 else [],
            "val_linha_2": df.iloc[2].tolist()[:12] if len(df) > 2 else [],
        }
        st.sidebar.write("Preview dos dados:")
        st.sidebar.json(preview_data)
        
        # Identifica as colunas importantes (pega a primeira ocorrência)
        col_sig = None
        col_nome = None
        col_dir = None
        
        for i, col in enumerate(col_names):
            col_upper = col.upper()
            if col_sig is None and ("OM" == col_upper or "SIGLA" in col_upper):
                col_sig = col
            if col_nome is None and "ORGANIZA" in col_upper:
                col_nome = col
            if col_dir is None and "DIRETORIA" in col_upper:
                col_dir = col
        
        st.sidebar.write("✅ Colunas identificadas:")
        st.sidebar.write(f"  - Sigla: {col_sig}")
        st.sidebar.write(f"  - Nome: {col_nome}")
        st.sidebar.write(f"  - Diretoria: {col_dir}")
        
        if col_sig and len(df) > 0:
            # Cria dataframe limpo
            tmp = pd.DataFrame({
                "sig": df[col_sig].apply(_clean),
                "nome": df[col_nome].apply(_clean) if col_nome else "",
                "dir": df[col_dir].apply(_clean) if col_dir else "",
            })
            
            # Remove linhas vazias e inválidas
            tmp = tmp[tmp["sig"] != ""]
            tmp = tmp[tmp["sig"].str.upper() != "OM"]
            tmp = tmp[tmp["sig"].str.upper() != "NR"]
            tmp = tmp[tmp["sig"].str.len() > 0]
            tmp = tmp[~tmp["sig"].str.contains("unnamed", case=False, na=False)]
            
            # Remove duplicatas (mantém a primeira ocorrência)
            tmp = tmp.drop_duplicates(subset=["sig"], keep="first")
            
            st.sidebar.write(f"✅ Total de OMs válidas: {len(tmp)}")
            
            # Mostra as primeiras 10 OMs encontradas
            if len(tmp) > 0:
                st.sidebar.write("Primeiras OMs encontradas:")
                for idx, row in tmp.head(10).iterrows():
                    st.sidebar.text(f"  {row['sig']} → {row['nome'][:30]}... | {row['dir']}")
            
            # Cria as opções para o selectbox
            for _, r in tmp.iterrows():
                if r['sig']:
                    label = f"{r['sig']} — {r['nome']}" if r['nome'] else r['sig']
                    options.append(label)
                    disp2sig[label] = r["sig"]
                    sig2dir[r["sig"]] = r["dir"] if r["dir"] else ""
            
            options = sorted(options)
        else:
            st.sidebar.warning("⚠️ Nenhuma coluna de OM identificada ou DataFrame vazio!")
            
    except Exception as e:
        st.sidebar.error(f"❌ Erro ao carregar dados: {e}")
        import traceback
        st.sidebar.code(traceback.format_exc(), language="python")
    
    # Adiciona opção para OM não listada
    options.append("Outra / não listada…")
    disp2sig["Outra / não listada…"] = ""
    
    st.sidebar.write(f"📊 Total final no dropdown: {len(options)}")
    
    return options, disp2sig, sig2dir


# ---------------------------------------------------------
# Callback: ao trocar a OM, atualiza Diretoria
# ---------------------------------------------------------
def _on_om_change(disp2sig: dict[str, str], sig2dir: dict[str, str]):
    choice = st.session_state.get("om_choice")
    sig = disp2sig.get(choice or "", "")
    st.session_state["dir_resp"] = sig2dir.get(sig, "") if sig else ""
    st.session_state["om_sigla_out"] = ""


# ---------------------------------------------------------
# Chaves do formulário
# ---------------------------------------------------------
FORM_KEYS = [
    "om_choice", "om_sigla_out", "dir_resp",
    "tipo_vist", "local_inst", "urg", "data_limite", "motivo",
    "ref_opus", "objetivo", "vt_exec", "status_sem", "obs",
    "data_vist", "data_prev_conc", "meio_resp", "data_resp",
    "num_opus", "qd_total", "qd_exec",
]


def _reset_form():
    for k in FORM_KEYS:
        if k in st.session_state:
            del st.session_state[k]


# =========================================================
# Formulário (UI)
# =========================================================
def _input_row(om_options: list[str], disp2sig: dict[str, str], sig2dir: dict[str, str]):
    st.subheader("📥 Nova solicitação de vistoria")

    c1, c2 = st.columns(2)

    # -------- coluna 1 --------
    with c1:
        choice = st.selectbox(
            "OM solicitante (sigla)",
            options=om_options,
            index=None,
            placeholder="Selecione a OM…",
            key="om_choice",
            on_change=_on_om_change,
            kwargs={"disp2sig": disp2sig, "sig2dir": sig2dir},
        )

        om_sigla = ""
        if choice:
            om_sigla = disp2sig.get(choice, "")
            if choice == "Outra / não listada…":
                om_sigla = st.text_input("Informe a OM (sigla)", key="om_sigla_out").strip()

        diretoria = st.text_input(
            "Diretoria responsável", 
            key="dir_resp",
            help="Preenchimento automático baseado na OM selecionada"
        )

        tipo_vistoria = st.selectbox(
            "Tipo de vistoria",
            ["Periódica", "Emergencial", "Preventiva", "Extraordinária"],
            index=0,
            key="tipo_vist",
        )

    # -------- coluna 2 --------
    with c2:
        local = st.text_input("Local / instalação", key="local_inst")
        urgencia = st.selectbox(
            "Urgência", ["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"], index=0, key="urg"
        )
        data_limite: date | None = st.date_input(
            "Data limite (se houver)", value=None, format="YYYY/MM/DD", key="data_limite"
        )

    motivo = st.text_area("Motivo / justificativa (NAOM)", height=120, key="motivo")

    # ---------- Complementares ----------
    st.markdown("### Complementares")
    cc1, cc2 = st.columns(2)
    with cc1:
        referencia_opus = st.text_input("REFERÊNCIA OPUS", key="ref_opus")
        objetivo_contato = st.text_input("OBJETIVO (ADICIONAR POSSÍVEL CONTATO)", key="objetivo")
        vt_exec_por = st.text_input("VT EXECUTADA POR", key="vt_exec")
        status_atual = st.text_input("STATUS - ATUALIZAÇÃO SEMANAL", key="status_sem")
        obs = st.text_area("OBSERVAÇÕES", height=90, key="obs")

    with cc2:
        data_vistoria: date | None = st.date_input(
            "DATA DA VISTORIA", value=None, format="YYYY/MM/DD", key="data_vist"
        )
        data_prev_conc: date | None = st.date_input(
            "DATA/PREVISÃO DE CONCLUSÃO", value=None, format="YYYY/MM/DD", key="data_prev_conc"
        )
        meio_resposta = st.text_input("MEIO DE RESPOSTA DA SOLICITAÇÃO", key="meio_resp")
        data_resposta: date | None = st.date_input(
            "DATA DA RESPOSTA A SOLICITAÇÃO", value=None, format="YYYY/MM/DD", key="data_resp"
        )
        num_opus = st.text_input("Nº OPUS DA VISTORIA (SE FOR O CASO)", key="num_opus")
        qd_total = st.number_input(
            "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO", min_value=0, step=1, value=0, key="qd_total"
        )
        qd_exec = st.number_input(
            "QUANTIDADE DE DIAS PARA EXECUÇÃO", min_value=0, step=1, value=0, key="qd_exec"
        )

    # ---------- Validações ----------
    erros = []
    if not om_sigla:
        erros.append("Informe a OM solicitante.")
    if not local.strip():
        erros.append("Informe o local/instalação.")
    if not motivo.strip():
        erros.append("Descreva o motivo/justificativa (NAOM).")

    if erros:
        st.warning("⚠️ Preencha os campos obrigatórios:\n• " + "\n• ".join(erros))

    row = {
        "OBJETO DE VISTORIA": motivo.strip(),
        "OM APOIADA": om_sigla,
        "Diretoria Responsável": diretoria.strip(),
        "Classificação de Urgência": urgencia,
        "Situação": "SOLICITADA",
        "DATA DA SOLICITAÇÃO": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "DATA DA SOLICITAÇÃO_2": data_limite.strftime("%Y-%m-%d") if data_limite else "",
        "REFERÊNCIA OPUS": referencia_opus.strip(),
        "OBJETIVO (ADICIONAR POSSÍVEL CONTATO)": objetivo_contato.strip(),
        "DATA DA VISTORIA": data_vistoria.strftime("%Y-%m-%d") if data_vistoria else "",
        "VT EXECUTADA POR": vt_exec_por.strip(),
        "STATUS - ATUALIZAÇÃO SEMANAL": status_atual.strip(),
        "DATA/PREVISÃO DE CONCLUSÃO": data_prev_conc.strftime("%Y-%m-%d") if data_prev_conc else "",
        "MEIO DE RESPOSTA DA SOLICITAÇÃO": meio_resposta.strip(),
        "DATA DA RESPOSTA A SOLICITAÇÃO": data_resposta.strftime("%Y-%m-%d") if data_resposta else "",
        "Nº OPUS DA VISTORIA (SE FOR O CASO)": num_opus.strip(),
        "QUANTIDADE DE DIAS PARA TOTAL ATENDIMENTO": (str(int(qd_total)) if qd_total else ""),
        "QUANTIDADE DE DIAS PARA EXECUÇÃO": (str(int(qd_exec)) if qd_exec else ""),
        "OBSERVAÇÕES": obs.strip(),
    }

    ok = (len(erros) == 0)
    return row, ok


# =========================================================
# Página
# =========================================================
def page():
    st.header("📝 VIS-001 — Cadastro de Solicitação de Vistoria")

    # Catálogo de OMs
    om_options, disp2sig, sig2dir = _build_om_catalog()

    # Formulário
    row, ok = _input_row(om_options, disp2sig, sig2dir)

    st.divider()
    if st.button("💾 Salvar na aba ACOMPANHAMENTO VISTORIAS", type="primary", disabled=not ok):
        try:
            clean_row = {k: ("" if pd.isna(v) else str(v)) for k, v in row.items()}
            append_row("ACOMPANHAMENTO VISTORIAS", clean_row)
            st.success("✅ Registro salvo com sucesso!")
            _reset_form()
            st.rerun()
        except Exception as e:
            st.error(f"❌ Falha ao salvar: {e}")
            import traceback
            st.error(traceback.format_exc())
