# core/models.py
# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass, asdict, field
from datetime import datetime
from typing import Optional, Literal
import re

StatusVistoria = Literal[
    "SOLICITADA",
    "AGENDADA",
    "EM_EXECUCAO",
    "FINALIZADA",
    "RELATORIO_GERADO",
    "INTEGRADA_OBRAS",
    "CANCELADA",
]

TipoVistoria = Literal["Periódica", "Emergencial", "Preventiva", "Extraordinária"]
UrgenciaVistoria = Literal["NÃO PRIORITÁRIO", "PRIORIDADE", "URGENTE"]


@dataclass
class SolicitacaoVistoria:
    """Modelo de dados para uma solicitação de vistoria (VIS-001)."""
    
    # Campos obrigatórios
    om_solicitante: str  # sigla
    local: str
    motivo: str
    
    # Campos com valores padrão
    numero: str = ""  # Gerado automaticamente se vazio
    om_nome: str = ""
    diretoria: str = ""
    coordenadas: str = ""
    tipo_vistoria: TipoVistoria = "Periódica"
    urgencia: UrgenciaVistoria = "NÃO PRIORITÁRIO"
    data_limite: Optional[str] = None
    anexos: str = ""
    status_atual: StatusVistoria = "SOLICITADA"
    criado_por: str = "Sistema"
    criado_em: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    
    def __post_init__(self):
        """Validações e ajustes após criação."""
        if not self.numero:
            self.numero = self._gerar_numero()
        
        # Validações básicas
        if not self.om_solicitante:
            raise ValueError("OM solicitante é obrigatória")
        if not self.local:
            raise ValueError("Local é obrigatório")
        if not self.motivo:
            raise ValueError("Motivo é obrigatório")
    
    def _gerar_numero(self) -> str:
        """Gera um número único para a solicitação."""
        ano = datetime.now().year
        # Idealmente, buscar o último número usado e incrementar
        # Por simplicidade, usar timestamp
        timestamp = datetime.now().strftime("%m%d%H%M%S")
        return f"VIS-{ano}-{timestamp}"
    
    def to_row(self) -> dict:
        """Converte para dicionário para salvar no Sheets."""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: dict) -> SolicitacaoVistoria:
        """Cria uma instância a partir de um dicionário."""
        return cls(**{k: v for k, v in data.items() if k in cls.__annotations__})


@dataclass
class HistoricoStatus:
    """Modelo de dados para histórico de status (VIS-003)."""
    
    numero: str
    status_de: str
    status_para: str
    responsavel: str
    justificativa: str = ""
    timestamp: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    
    def to_row(self) -> dict:
        """Converte para dicionário para salvar no Sheets."""
        return asdict(self)


@dataclass
class RegistroRelatorio:
    """Modelo de dados para relatórios gerados (VIS-005)."""
    
    numero: str
    titulo: str
    gerado_por: str
    arquivo_pdf: str = ""
    gerado_em: str = field(default_factory=lambda: datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    
    def to_row(self) -> dict:
        """Converte para dicionário para salvar no Sheets."""
        return asdict(self)


__all__ = [
    "SolicitacaoVistoria",
    "HistoricoStatus",
    "RegistroRelatorio",
    "StatusVistoria",
    "TipoVistoria",
    "UrgenciaVistoria",
]
