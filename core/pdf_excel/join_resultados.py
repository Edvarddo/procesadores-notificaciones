#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re

import pandas as pd


def limpiar_direccion(dir_texto):
    if not isinstance(dir_texto, str):
        return dir_texto

    texto = dir_texto
    patrones = [
        r"^\s*AVENIDA\s*,?\s*",
        r"^\s*CALLE\s*,?\s*",
        r"^\s*PASAJE\s*,?\s*",
        r"^\s*AV\.?\s*,?\s*",
        r"^\s*PJE\.?\s*,?\s*",
    ]

    for patron in patrones:
        texto = re.sub(patron, "", texto, flags=re.IGNORECASE)

    return re.sub(r"\s+", " ", texto).strip()


def reordenar_nombre(nombre):
    if not isinstance(nombre, str):
        return nombre

    partes = nombre.strip().split()
    if len(partes) < 3:
        return nombre

    apellidos = partes[-2:]
    nombres = partes[:-2]
    return " ".join(apellidos + nombres)


def unir_dataframes(df_pdf: pd.DataFrame, df_excel: pd.DataFrame) -> pd.DataFrame:
    if "HASH" not in df_pdf.columns:
        raise ValueError("Falta la columna HASH en df_pdf")
    if "HASH" not in df_excel.columns:
        raise ValueError("Falta la columna HASH en df_excel")

    df_final = df_pdf.merge(
        df_excel,
        on="HASH",
        how="left",
        indicator=True,
        suffixes=("_PDF", "_EXCEL"),
    )

    if "ID_NOTIFICACION" in df_final.columns:
        tiene_id = df_final["ID_NOTIFICACION"].notna() & (
            df_final["ID_NOTIFICACION"].astype(str).str.strip() != ""
        )

        con_id = df_final[tiene_id].copy()
        sin_id = df_final[~tiene_id].copy()

        con_id = con_id.drop_duplicates(subset=["HASH", "ID_NOTIFICACION"], keep="first")
        sin_id = sin_id.drop_duplicates(keep="first")

        df_final = pd.concat([con_id, sin_id], ignore_index=True)
    else:
        df_final = df_final.drop_duplicates(keep="first").copy()

    if "DIRECCION" in df_final.columns:
        df_final["DIRECCION_LIMPIA"] = df_final["DIRECCION"].apply(limpiar_direccion)

    if "NOMBRE_PDF" in df_final.columns:
        df_final["NOMBRE_PDF_ORDENADO"] = df_final["NOMBRE_PDF"].apply(reordenar_nombre)

    if "NOMBRE_EXCEL" in df_final.columns:
        df_final["NOMBRE_EXCEL_ORDENADO"] = df_final["NOMBRE_EXCEL"].apply(reordenar_nombre)

    columnas_finales = [
        "HASH",
        "RUC_PDF",
        "RIT_PDF",
        "ANO_PDF",
        #"FECHA_AUDIENCIA_PDF",
        #"NOMBRE_PDF",
        "NOMBRE_PDF_ORDENADO",
        "TIPO_NOTIFICACION_PDF",
        "HORA",
        "CODIGO",
        #"ID_NOTIFICACION",
        "TIPO",
        #"DIRECCION",
        "DIRECCION_LIMPIA",
        #"RUT",
        #"_merge",
    ]

    columnas_finales = [c for c in columnas_finales if c in df_final.columns]
    return df_final[columnas_finales]
