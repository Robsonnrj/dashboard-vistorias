# core/audit.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import pandas as pd
from datetime import datetime
from typing import Optional
import streamlit as st

from .data_loader import read_df, overwrite_tab_from_df, append_row
from .config import TAB_SOLICITACOES, TAB_AUDIT


def registrar_historico(
    numero: str,
    de: str,
    para: str,
    justificativa: str,
    responsavel: str,
    campo_alterado: Optional[str] = "status"
) -> bool:
    """
    Registra um evento de auditoria na aba de histórico.
    
    Args:
        numero: Identificador do registro
        de: Valor anterior
        para: Novo valor
        justificativa: Motivo da alteração
        responsavel: Quem fez a alteração
        campo_alterado: Nome do campo alterado
    
    Returns:
        True se registrado com sucesso, False caso contrário
    """
    row = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "numero": str(numero),
        "campo": str(campo_alterado),
        "valor_anterior": str(de or ""),
        "valor_novo": str(para or ""),
        "justificativa": str(justificativa or ""),
        "responsavel": str(responsavel or "Sistema"),
    }
    
    try:
        append_row(TAB_AUDIT, row)
        return True
    except Exception as e:
        # Log do erro mas não interrompe o fluxo
        st.warning(f"⚠️ Não foi possível registrar auditoria: {e}")
        return False


def atualizar_status(
    numero: str,
    novo_status: str,
    justificativa: str = "",
    responsavel: str = "Sistema"
) -> bool:
    """
    Atualiza o status do registro e registra a trilha de auditoria.
    
    Args:
        numero: Identificador do registro
        novo_status: Novo status a ser aplicado
        justificativa: Motivo da mudança
        responsavel: Quem fez a alteração
    
    Returns:
        True se atualizado com sucesso
    """
    try:
        df = read_df(TAB_SOLICITACOES, use_cache=False)
        
        if df.empty:
            st.error("❌ Planilha base está vazia")
            return False
        
        # Procura coluna de número/ID
        col_numero = None
        for col in df.columns:
            if "numero" in col.lower() or "id" in col.lower():
                col_numero = col
                break
        
        if not col_numero:
            st.error("❌ Coluna de número/ID não encontrada")
            return False
        
        # Localiza o registro
        mask = pd.to_numeric(df[col_numero], errors="coerce") == pd.to_numeric(numero, errors="coerce")
        
        if not mask.any():
            st.error(f"❌ Registro número {numero} não encontrado")
            return False
        
        # Procura coluna de status
        col_status = None
        for col in df.columns:
            if "status" in col.lower() or "situação" in col.lower() or "situacao" in col.lower():
                col_status = col
                break
        
        if not col_status:
            # Cria a coluna se não existir
            col_status = "Situação"
            df[col_status] = ""
        
        # Pega o status anterior
        status_anterior = df.loc[mask, col_status].iloc[0] if mask.any() else ""
        
        # Atualiza o status
        df.loc[mask, col_status] = novo_status
        
        # Salva de volta
        overwrite_tab_from_df(TAB_SOLICITACOES, df, keep_header=True)
        
        # Registra auditoria
        registrar_historico(
            numero=numero,
            de=status_anterior,
            para=novo_status,
            justificativa=justificativa,
            responsavel=responsavel,
            campo_alterado="status"
        )
        
        return True
        
    except Exception as e:
        st.error(f"❌ Erro ao atualizar status: {e}")
        return False


def trilha(numero: str) -> pd.DataFrame:
    """
    Retorna o histórico de alterações de um registro.
    
    Args:
        numero: Identificador do registro
    
    Returns:
        DataFrame com o histórico ordenado por data
    """
    try:
        hist = read_df(TAB_AUDIT, use_cache=False)
        
        if hist.empty:
            return pd.DataFrame()
        
        # Procura coluna de número
        col_numero = None
        for col in hist.columns:
            if "numero" in col.lower():
                col_numero = col
                break
        
        if not col_numero:
            return pd.DataFrame()
        
        # Filtra pelo número
        filtered = hist[hist[col_numero].astype(str) == str(numero)].copy()
        
        # Ordena por timestamp
        col_ts = None
        for col in filtered.columns:
            if "timestamp" in col.lower() or "data" in col.lower():
                col_ts = col
                break
        
        if col_ts:
            filtered[col_ts] = pd.to_datetime(filtered[col_ts], errors="coerce")
            filtered = filtered.sort_values(col_ts, ascending=False)
        
        return filtered
        
    except Exception:
        return pd.DataFrame()


__all__ = ["registrar_historico", "atualizar_status", "trilha"]
