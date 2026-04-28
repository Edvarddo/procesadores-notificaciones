#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import unicodedata
from pathlib import Path

import fitz
import pandas as pd

from .clave_utils import limpiar, norm_fecha, crear_clave, crear_hash_simple

PATRON_RUC = re.compile(r"^\d{7,}-[\dKk]$")
PATRON_HORA = re.compile(r"^\d{2}:\d{2}$")
PATRON_FECHA = re.compile(r"^\d{2}-\d{2}-\d{4}$")
PATRON_NUMERO = re.compile(r"^\d+$")
PATRON_CODIGO = re.compile(r"^[A-Z]\d+$")

RUIDO_EXACTO = {
    "CERTIFICADA",
    "CERTIFICADA-",
    "PENDIENTE",
    "PENDIENTE DE",
    "TRASPASO",
    "DE",
    "C/CERTIFICACION EN",
    "BUSQUEDA",
}


def quitar_tildes(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def norm(texto: str) -> str:
    return limpiar(quitar_tildes(str(texto))).upper()


def norm_nombre_compare(texto: str) -> str:
    t = norm(texto)
    t = re.sub(r"[^A-Z0-9 ]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def es_ruido(linea: str) -> bool:
    l = norm(linea)
    if l in RUIDO_EXACTO:
        return True
    if l.startswith("CERTIFICADA"):
        return True
    if l.startswith("PENDIENTE DE"):
        return True
    return False


def mapear_tipo(tokens_tipo: list[str]) -> str:
    t = norm(" ".join(tokens_tipo))

    if "CEDULA" in t:
        return "Por cedula"

    if "PERSONAL" in t:
        return "Personal/Art44"

    if "SUST" in t and "44" in t:
        return "Personal/Art44"

    return t


def extraer_lineas(pdf_path: str | Path) -> list[str]:
    doc = fitz.open(str(pdf_path))
    lineas = []

    for page in doc:
        texto = page.get_text("text")
        for linea in texto.splitlines():
            linea = limpiar(linea)
            if linea:
                lineas.append(linea)

    return lineas


def construir_bloques(lineas: list[str]) -> list[list[str]]:
    indices_ruc = [i for i, l in enumerate(lineas) if PATRON_RUC.fullmatch(l)]
    bloques = []

    for pos, inicio in enumerate(indices_ruc):
        fin = indices_ruc[pos + 1] if pos + 1 < len(indices_ruc) else len(lineas)
        bloques.append(lineas[inicio:fin])

    return bloques


def detectar_tipo_y_posicion(bloque: list[str], idx_hora: int):
    for i in range(6, idx_hora):
        t = norm(bloque[i])

        if t == "CEDULA":
            return i, [bloque[i]]

        if "PERSONAL/ART" in t and len(t.split()) <= 2:
            tokens = [bloque[i]]
            j = i + 1
            while j < idx_hora and norm(bloque[j]) == "44":
                tokens.append(bloque[j])
                j += 1
            return i, tokens

        if t.startswith("SUST"):
            tokens = [bloque[i]]
            j = i + 1
            while j < idx_hora and (norm(bloque[j]) == "44" or "ARTICULO" in norm(bloque[j])):
                tokens.append(bloque[j])
                j += 1
            return i, tokens

    return -1, []


def extraer_fecha_audiencia(bloque: list[str]) -> str:
    fechas = [limpiar(x) for x in bloque if PATRON_FECHA.fullmatch(limpiar(x))]
    if not fechas:
        return ""
    if len(fechas) >= 2:
        return fechas[1]
    return fechas[0]


def extraer_gestiones(bloque: list[str]) -> list[tuple[str, str, int]]:
    """
    Detecta gestiones en ambos formatos:
    - HORA -> CODIGO
    - CODIGO -> HORA

    Devuelve:
        [(codigo, hora, idx_hora), ...]
    en el orden en que aparecen.
    """
    gestiones = []
    vistos = set()

    for i in range(len(bloque) - 1):
        a = limpiar(bloque[i])
        b = limpiar(bloque[i + 1])

        a_norm = norm(bloque[i])
        b_norm = norm(bloque[i + 1])

        # Caso 1: HORA -> CODIGO
        if PATRON_HORA.fullmatch(a) and PATRON_CODIGO.fullmatch(b_norm):
            clave = (b_norm, a, i)
            if clave not in vistos:
                gestiones.append((b_norm, a, i))
                vistos.add(clave)

        # Caso 2: CODIGO -> HORA
        elif PATRON_CODIGO.fullmatch(a_norm) and PATRON_HORA.fullmatch(b):
            clave = (a_norm, b, i + 1)
            if clave not in vistos:
                gestiones.append((a_norm, b, i + 1))
                vistos.add(clave)

    return gestiones


def parsear_bloque(bloque: list[str]) -> dict | None:
    if len(bloque) < 9:
        return None

    ruc = bloque[0]
    if not PATRON_RUC.fullmatch(ruc):
        return None

    rit = bloque[2] if len(bloque) > 2 else ""
    ano = bloque[3] if len(bloque) > 3 else ""

    if not rit.isdigit() or not ano.isdigit():
        return None

    gestiones = extraer_gestiones(bloque)
    if not gestiones:
        return None

    # Última gestión encontrada = más actualizada
    codigo_certificacion, hora_gestion, idx_hora = gestiones[-1]

    idx_tipo, tipo_tokens = detectar_tipo_y_posicion(bloque, idx_hora)
    nombre_original = ""

    if idx_tipo != -1:
        nombre_tokens = []
        for x in bloque[6:idx_tipo]:
            if es_ruido(x):
                continue
            if PATRON_FECHA.fullmatch(limpiar(x)):
                continue
            if PATRON_HORA.fullmatch(limpiar(x)):
                continue
            if PATRON_NUMERO.fullmatch(limpiar(x)):
                continue
            if PATRON_CODIGO.fullmatch(norm(x)):
                continue
            nombre_tokens.append(x)

        nombre_original = limpiar(" ".join(nombre_tokens))

    else:
        linea_nombre_tipo = ""
        idx_linea_nombre_tipo = -1

        for i in range(6, idx_hora):
            nx = norm(bloque[i])
            if "CEDULA" in nx or "PERSONAL/ART." in nx or "PERSONAL/ART" in nx or "SUST" in nx:
                linea_nombre_tipo = bloque[i]
                idx_linea_nombre_tipo = i
                break

        if not linea_nombre_tipo:
            return None

        linea_norm = norm(linea_nombre_tipo)

        if "CEDULA" in linea_norm:
            pos = linea_norm.rfind("CEDULA")
            nombre_original = limpiar(linea_nombre_tipo[:pos])
            tipo_tokens = ["Cedula"]

        elif "PERSONAL/ART." in linea_norm:
            pos = linea_norm.rfind("PERSONAL/ART.")
            nombre_original = limpiar(linea_nombre_tipo[:pos])
            tipo_tokens = ["Personal/Art."]
            if idx_linea_nombre_tipo + 1 < idx_hora and norm(bloque[idx_linea_nombre_tipo + 1]) == "44":
                tipo_tokens.append("44")

        elif "PERSONAL/ART" in linea_norm:
            pos = linea_norm.rfind("PERSONAL/ART")
            nombre_original = limpiar(linea_nombre_tipo[:pos])
            tipo_tokens = ["Personal/Art"]
            if idx_linea_nombre_tipo + 1 < idx_hora and norm(bloque[idx_linea_nombre_tipo + 1]) == "44":
                tipo_tokens.append("44")

        elif "SUST" in linea_norm:
            m = re.search(r"\bSUST\b|\bSUST\.", linea_norm)
            if not m:
                return None
            pos = m.start()
            nombre_original = limpiar(linea_nombre_tipo[:pos])
            tipo_tokens = ["Sust."]
            j = idx_linea_nombre_tipo + 1
            while j < idx_hora and ("ARTICULO" in norm(bloque[j]) or norm(bloque[j]) == "44"):
                tipo_tokens.append(bloque[j])
                j += 1

    if not nombre_original:
        return None

    nombre = limpiar(nombre_original)
    fecha_audiencia = extraer_fecha_audiencia(bloque)

    ruc_norm = norm(ruc)
    rit_norm = limpiar(rit)
    ano_norm = limpiar(ano)
    nombre_norm = norm_nombre_compare(nombre)
    fecha_audiencia_norm = norm_fecha(fecha_audiencia)

    clave = crear_clave(ruc_norm, rit_norm, nombre_norm, fecha_audiencia_norm)
    hash_id = crear_hash_simple(ruc_norm, rit_norm, nombre_norm, fecha_audiencia_norm)

    return {
        "RUC": ruc_norm,
        "RIT": rit_norm,
        "ANO": ano_norm,
        "TIPO_NOTIFICACION": mapear_tipo(tipo_tokens),
        "NOMBRE": nombre_norm,
        "FECHA_AUDIENCIA": fecha_audiencia_norm,
        "HORA": hora_gestion.replace(":", ""),
        "CODIGO": codigo_certificacion,
        "CLAVE": clave,
        "HASH": hash_id,
    }


def procesar_pdf(pdf_path: str | Path) -> tuple[pd.DataFrame, list[list[str]]]:
    lineas = extraer_lineas(pdf_path)
    bloques = construir_bloques(lineas)

    filas = []
    no_parseados = []

    for bloque in bloques:
        fila = parsear_bloque(bloque)
        if fila:
            filas.append(fila)
        else:
            no_parseados.append(bloque)

    df = pd.DataFrame(filas)

    if not df.empty:
        df["HORA_NUM"] = pd.to_numeric(df["HORA"], errors="coerce")
        df = df.sort_values(["HORA_NUM", "RUC", "RIT"], ascending=[True, True, True]).drop(columns=["HORA_NUM"])
        df = df[
            [
                "RUC",
                "RIT",
                "ANO",
                "TIPO_NOTIFICACION",
                "NOMBRE",
                "FECHA_AUDIENCIA",
                "HORA",
                "CODIGO",
                "CLAVE",
                "HASH",
            ]
        ]

    return df, no_parseados
