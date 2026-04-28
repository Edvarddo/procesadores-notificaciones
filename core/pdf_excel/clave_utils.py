import re
import unicodedata
import pandas as pd
import hashlib


def limpiar(texto: str) -> str:
    return re.sub(r"\s+", " ", str(texto)).strip()


def quitar_tildes(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    )


def norm_nombre(texto) -> str:
    if pd.isna(texto):
        return ""
    return limpiar(quitar_tildes(str(texto))).upper()


def norm_fecha(valor) -> str:
    if pd.isna(valor):
        return ""

    texto = limpiar(str(valor))
    if not texto:
        return ""

    # Ya viene normalizada
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", texto):
        return texto

    # Formato dd-mm-yyyy
    if re.fullmatch(r"\d{2}-\d{2}-\d{4}", texto):
        try:
            dt = pd.to_datetime(texto, format="%d-%m-%Y", errors="raise")
            return dt.strftime("%Y-%m-%d")
        except Exception:
            pass

    # Formato dd/mm/yyyy
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", texto):
        try:
            dt = pd.to_datetime(texto, format="%d/%m/%Y", errors="raise")
            return dt.strftime("%Y-%m-%d")
        except Exception:
            pass

    # Último intento genérico
    try:
        dt = pd.to_datetime(texto, errors="raise")
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return texto


def crear_clave(ruc, rit, nombre, fecha) -> str:
    return f"{limpiar(ruc)}|{limpiar(rit)}|{norm_nombre(nombre)}|{norm_fecha(fecha)}"


def crear_clave_completa(ruc, rit, nombre, fecha) -> str:
    return crear_clave(ruc, rit, nombre, fecha)


def crear_hash_simple(ruc, rit, nombre, fecha) -> str:
    clave = crear_clave(ruc, rit, nombre, fecha)
    return hashlib.md5(clave.encode("utf-8")).hexdigest()
