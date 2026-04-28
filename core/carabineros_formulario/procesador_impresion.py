"""
Procesador para archivos de impresión de notificaciones de Carabineros.
Extrae IDs de notificación y genera CSV para automatización.
"""

import os
import pandas as pd
from typing import List
from dataclasses import dataclass


@dataclass
class RegistroImpresion:
    """Registro extraído de archivo de impresión"""
    id_notificacion: str


def buscar_columna_id_notificacion(df: pd.DataFrame) -> str:
    """
    Busca la columna que contiene el ID de notificación.
    Intenta varios nombres comunes.
    """
    nombres_comunes = [
        'ID', 'id', 'ID_NOTIFICACION', 'id_notificacion', 'ID Notificación',
        'Número', 'numero', 'Notificación', 'notificacion',
        'NOTIFICACION', 'ID_NOTI', 'id_noti'
    ]
    
    for col in nombres_comunes:
        if col in df.columns:
            return col
    
    # Si no encuentra, retorna la primera columna como fallback
    return df.columns[0]


def leer_archivo_impresion(ruta_archivo: str) -> List[RegistroImpresion]:
    """
    Lee archivo de impresión (XLS/XLSX) y extrae IDs de notificación.
    
    Args:
        ruta_archivo: Ruta al archivo Excel
        
    Returns:
        Lista de RegistroImpresion con IDs de notificación
    """
    
    if not os.path.exists(ruta_archivo):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta_archivo}")
    
    # Detectar tipo de archivo
    if ruta_archivo.endswith('.xlsx'):
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
    elif ruta_archivo.endswith('.xls'):
        df = pd.read_excel(ruta_archivo, engine='xlrd')
    else:
        # Intentar leer como CSV
        df = pd.read_csv(ruta_archivo)
    
    # Remover filas completamente vacías
    df = df.dropna(how='all')
    
    # Buscar columna de ID
    col_id = buscar_columna_id_notificacion(df)
    
    registros = []
    for idx, row in df.iterrows():
        id_notif = str(row[col_id]).strip()
        
        # Filtrar filas inválidas:
        # - Vacías o NaN
        # - Headers (contienen palabras como "notificación", "id", "detalle", "fecha", etc.)
        # - Que no sean principalmente números
        if not id_notif or id_notif.lower() == 'nan':
            continue
        
        # Ignorar filas que parecen ser headers o texto descriptivo
        palabras_header = [
            'id', 'notificación', 'numero', 'detalle', 'fecha', 'emisión', 
            'impresión', 'nombre', 'estado', 'cabecera', 'encabezado',
            'header', 'column', 'field'
        ]
        
        es_header = any(palabra in id_notif.lower() for palabra in palabras_header)
        
        if es_header:
            continue
        
        # Validar que sea principalmente numérico (al menos 5 dígitos)
        # Los IDs de notificación son típicamente números como: 18060023, 12345678, etc.
        solo_numeros = ''.join(c for c in id_notif if c.isdigit())
        
        if len(solo_numeros) < 5:  # Debe tener al menos 5 dígitos
            continue
        
        # Si pasó todas las validaciones, agregar como registro válido
        registros.append(RegistroImpresion(id_notificacion=id_notif))
    
    if not registros:
        raise ValueError("No se encontraron IDs de notificación válidos en el archivo")
    
    return registros


def generar_csv_desde_impresion(
    ruta_impresion: str,
    codigo: str,
    hora: str,
    ruta_salida: str = None,
    observacion: str = "."
) -> str:
    """
    Lee archivo de impresión, extrae IDs y genera CSV con código + hora + observación.
    
    Args:
        ruta_impresion: Ruta al archivo de impresión
        codigo: Código a asignar (ej: D2)
        hora: Hora a asignar (ej: 1205)
        ruta_salida: Ruta del CSV de salida (opcional)
        observacion: Observación a asignar (por defecto: ".")
        
    Returns:
        Ruta del archivo CSV generado
    """
    
    # Validar entrada
    if not hora.isdigit() or len(hora) != 4:
        raise ValueError("Hora debe tener formato HHMM (ej: 1205)")
    
    if not codigo or len(codigo.strip()) == 0:
        raise ValueError("Código no puede estar vacío")
    
    # Leer archivo de impresión
    registros = leer_archivo_impresion(ruta_impresion)
    
    # Si no especifica ruta de salida, usar el mismo directorio
    if not ruta_salida:
        dir_base = os.path.dirname(ruta_impresion)
        nombre_base = os.path.splitext(os.path.basename(ruta_impresion))[0]
        ruta_salida = os.path.join(dir_base, f"{nombre_base}_procesado.csv")
    
    # Crear DataFrame sin columnas rit y anio (no se usan en impresiones)
    # El procesador normal buscará por id_notificacion
    datos = []
    for reg in registros:
        datos.append({
            'id_notificacion': reg.id_notificacion,
            'codigo': codigo,
            'hora': hora,
            'observacion': observacion
        })
    
    df = pd.DataFrame(datos)
    
    # Asegurar que el directorio existe
    os.makedirs(os.path.dirname(ruta_salida) or '.', exist_ok=True)
    
    # Guardar CSV
    df.to_csv(ruta_salida, index=False, encoding='utf-8-sig')
    
    return ruta_salida


def previsualizar_impresion(ruta_archivo: str, max_registros: int = 5) -> List[RegistroImpresion]:
    """
    Lee archivo de impresión y retorna primeros N registros para preview.
    """
    registros = leer_archivo_impresion(ruta_archivo)
    return registros[:max_registros]
