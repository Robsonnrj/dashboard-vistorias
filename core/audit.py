# -*- coding: utf-8 -*-
import pandas as pd
from datetime import datetime
from .data_loader import read_df, overwrite_tab_from_df, append_row
from .config import TAB_SOLICITACOES, TAB_AUDIT

def registrar_historico(numero: str, de: str, para: str, justificativa: str, responsavel: str):
    """Acrescenta um evento de auditoria em TAB_AUDIT (se a aba existir/for usada)."""
    row = {
        "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "numero": str(numero),
        "status_de": str(de or ""),
        "status_para": str(para or ""),
        "justificativa": str(justificativa or ""),
        "responsavel": str(responsavel or ""),
    }
    try:
        append_row(TAB_AUDIT, row)
    except Exception:
        # Se a aba não existir, simplesmente ignora (sistema continua funcionando)
        pass

def atualizar_status(numero: str, novo_status: str, justificativa: str, responsavel: str):
    """Atualiza o status do registro na aba base e registra a trilha."""
    df = read_df(TAB_SOLICITACOES)
    if df.empty or "numero" not in df.columns:
        raise ValueError("Planilha base sem dados ou sem a coluna 'numero'.")

    # localiza linha pelo número
    m = pd.to_numeric(df["numero"], errors="coerce") == pd.to_numeric(numero, errors="coerce")
    if not m.any():
        raise ValueError(f"Número {numero} não encontrado na aba '{TAB_SOLICITACOES}'.")

    # status atual (se existir)
    col_status = next((c for c in df.columns if c.lower().strip() in ("status_atual","status")), None)
    if not col_status:
        # se não existir, cria a coluna
        col_status = "status_atual"
        if col_status not in df.columns:
            df[col_status] = ""

    antigo = df.loc[m, col_status].iloc[0] if m.any() else ""
    # atualiza
    df.loc[m, col_status] = novo_status

    # escreve de volta
    overwrite_tab_from_df(TAB_SOLICITACOES, df, keep_header=True)

    # audita (se possível)
    registrar_historico(numero=str(numero), de=str(antigo or ""), para=str(novo_status or ""),
                        justificativa=justificativa, responsavel=responsavel)

def trilha(numero: str) -> pd.DataFrame:
    """Retorna a trilha (se a aba existir)."""
    try:
        hist = read_df(TAB_AUDIT)
    except Exception:
        return pd.DataFrame()
    if hist.empty or "numero" not in hist.columns:
        return pd.DataFrame()
    return hist[ (hist["numero"].astype(str) == str(numero)) ].sort_values("ts", ascending=False)
