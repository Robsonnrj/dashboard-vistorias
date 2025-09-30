# -*- coding: utf-8 -*-
import pandas as pd
from .data_loader import read_df, append_row, overwrite_tab

def registrar_historico(row: dict):
    append_row("historicos", row)

def trilha(numero: str) -> pd.DataFrame:
    df = read_df("historicos")
    if df.empty:
        return df
    return df[df["numero"].astype(str) == str(numero)].sort_values("timestamp")

def atualizar_status(numero: str, novo_status: str, justificativa: str, responsavel: str):
    df = read_df("solicitacoes")
    if df.empty:
        raise RuntimeError("Base de solicitações vazia.")
    m = df["numero"].astype(str) == str(numero)
    if not m.any():
        raise RuntimeError(f"Solicitação {numero} não encontrada.")
    status_de = df.loc[m, "status_atual"].iloc[0]
    # atualiza
    df.loc[m, "status_atual"] = novo_status
    overwrite_tab("solicitacoes", df)
    # registra auditoria
    registrar_historico({
        "numero": numero,
        "status_de": status_de,
        "status_para": novo_status,
        "justificativa": justificativa,
        "responsavel": responsavel
    })
