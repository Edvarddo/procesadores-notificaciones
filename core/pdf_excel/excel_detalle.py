#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path

import pandas as pd

from .clave_utils import limpiar, norm_nombre, norm_fecha, crear_clave, crear_hash_simple


def limpiar_clave(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = limpiar(str(valor))
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto


def obtener_tipo_final(df: pd.DataFrame) -> pd.Series:
    """
    Regla:
    - base = TIPO TRAMITE
    - si TIPO DE CAUSA == 'EXHORTO', entonces TIPO = 'Exhorto'
    """
    if "TIPO TRAMITE" in df.columns:
        tipo_tramite = df["TIPO TRAMITE"].apply(limpiar)
    else:
        tipo_tramite = pd.Series([""] * len(df), index=df.index)

    if "TIPO DE CAUSA" in df.columns:
        tipo_causa = df["TIPO DE CAUSA"].apply(limpiar)
    else:
        tipo_causa = pd.Series([""] * len(df), index=df.index)

    tipo_final = tipo_tramite.copy()

    mask_exhorto = tipo_causa.str.upper() == "EXHORTO"
    tipo_final.loc[mask_exhorto] = "Exhorto"

    return tipo_final


def procesar_excel(archivo_xls: str | Path) -> pd.DataFrame:
    archivo = Path(archivo_xls)
    df = pd.read_excel(archivo, header=6)

    df.columns = [str(c).strip().upper() for c in df.columns]

    columnas_necesarias = [
        "RUC",
        "RIT",
        "AÑO",
        "NOMBRE PARTICIPANTE",
        "DIRECCIÓN",
        "FECHA AUDIENCIA",
    ]

    for col in columnas_necesarias:
        if col not in df.columns:
            raise ValueError(f"Falta la columna: {col}. Detectadas: {df.columns.tolist()}")

    resultado = pd.DataFrame({
        "RUC": df["RUC"].apply(limpiar_clave),
        "RIT": df["RIT"].apply(limpiar_clave),
        "ANO": df["AÑO"].apply(limpiar_clave),
        "NOMBRE": df["NOMBRE PARTICIPANTE"].apply(norm_nombre),
        "FECHA_AUDIENCIA": df["FECHA AUDIENCIA"].apply(norm_fecha),
        "DIRECCION": df["DIRECCIÓN"].apply(limpiar),
    })

    if "ID NOTIFICACIÓN" in df.columns:
        resultado["ID_NOTIFICACION"] = df["ID NOTIFICACIÓN"].apply(limpiar_clave)
    elif "ID NOTIFICACION" in df.columns:
        resultado["ID_NOTIFICACION"] = df["ID NOTIFICACION"].apply(limpiar_clave)
    else:
        resultado["ID_NOTIFICACION"] = ""

    if "RUT" in df.columns:
        resultado["RUT"] = df["RUT"].apply(limpiar)

    if "ID CAUSA" in df.columns:
        resultado["ID_CAUSA"] = df["ID CAUSA"].apply(limpiar)

    if "TIPO NOTIFICACIÓN" in df.columns:
        resultado["TIPO_NOTIFICACION"] = df["TIPO NOTIFICACIÓN"].apply(limpiar)
    elif "TIPO NOTIFICACION" in df.columns:
        resultado["TIPO_NOTIFICACION"] = df["TIPO NOTIFICACION"].apply(limpiar)

    resultado["TIPO"] = obtener_tipo_final(df)

    if "TIPO DE CAUSA" in df.columns:
        resultado["TIPO_CAUSA"] = df["TIPO DE CAUSA"].apply(limpiar)

    resultado = resultado[
        (resultado["RUC"] != "") &
        (resultado["RIT"] != "") &
        (resultado["NOMBRE"] != "") &
        (resultado["FECHA_AUDIENCIA"] != "")
    ].copy()

    resultado["CLAVE"] = resultado.apply(
        lambda row: crear_clave(
            row["RUC"],
            row["RIT"],
            row["NOMBRE"],
            row["FECHA_AUDIENCIA"],
        ),
        axis=1,
    )

    resultado["HASH"] = resultado.apply(
        lambda row: crear_hash_simple(
            row["RUC"],
            row["RIT"],
            row["NOMBRE"],
            row["FECHA_AUDIENCIA"],
        ),
        axis=1,
    )

    resultado["CLAVE_DUPLICADA"] = resultado.duplicated(subset=["CLAVE"], keep=False)

    columnas_salida = [
        "ID_NOTIFICACION",
        "RUC",
        "RIT",
        "ANO",
        "NOMBRE",
        "FECHA_AUDIENCIA",
        "TIPO",
        "TIPO_CAUSA",
        "TIPO_NOTIFICACION",
        "DIRECCION",
        "RUT",
        "ID_CAUSA",
        "CLAVE",
        "HASH",
        "CLAVE_DUPLICADA",
    ]
    columnas_salida = [c for c in columnas_salida if c in resultado.columns]

    return resultado[columnas_salida]
