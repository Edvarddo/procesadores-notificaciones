from dataclasses import dataclass
from typing import Optional


@dataclass
class RegistroNotificacion:
    rit: int
    anio: int
    id_notificacion: str
    hora: str
    codigo: str
    observacion: str


@dataclass
class ResultadoProceso:
    rit: int
    anio: int
    id_notificacion: str
    hora: Optional[str] = None
    codigo: Optional[str] = None
    observacion: Optional[str] = None
    estado: str = ""
    mensaje: str = ""
    gestion_usada: Optional[int] = None