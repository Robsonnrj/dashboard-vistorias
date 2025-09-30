# -*- coding: utf-8 -*-
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Optional, Literal

StatusVistoria = Literal[
    "SOLICITADA",
    "AGENDADA",
    "EM_EXECUCAO",
    "FINALIZADA",
    "RELATORIO_GERADO",
    "INTEGRADA_OBRAS",
]

TipoVistoria = Literal["Periódica", "Emergencial", "Preventiva", "Extraordinária"]

@dataclass
class SolicitacaoVistoria:
    # VIS-001
    numero: str                 # ex.: NAOM/2025-0001 (pode ser gerado)
    om_solicitante: str         # sigla
    om_nome: str                # nome completo
    diretoria: str
    local: str                  # endereço ou instalação
    coordenadas: str            # opcional "lat,lon"
    tipo_vistoria: TipoVistoria
    motivo: str
    urgencia: str               # ex.: NÃO PRIORITÁRIO / PRIORIDADE / URGENTE
    data_limite: Optional[str]  # ISO (YYYY-MM-DD)
    anexos: str                 # referências (ex.: URL do DIEx, Drive)
    status_atual: StatusVistoria = "SOLICITADA"
    criado_por: str = "usuario"
    criado_em: str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def to_row(self) -> dict:
        return asdict(self)

@dataclass
class HistoricoStatus:
    # VIS-003
    numero: str                 # vincula à solicitação
    status_de: StatusVistoria
    status_para: StatusVistoria
    justificativa: str
    responsavel: str
    timestamp: str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def to_row(self) -> dict:
        return asdict(self)

@dataclass
class RegistroRelatorio:
    # VIS-005 (metadados)
    numero: str
    titulo: str
    arquivo_pdf: str            # nome do arquivo gerado (salvo no Drive ou local)
    gerado_por: str
    gerado_em: str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def to_row(self) -> dict:
        return asdict(self)
