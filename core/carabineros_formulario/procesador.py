import pandas as pd
from pathlib import Path

from core.carabineros_formulario.utils.browser import crear_driver
from core.carabineros_formulario.data.loader import cargar_registros_csv
from core.carabineros_formulario.services.procesamiento_service import (
    procesar_registros,
    exportar_resultados,
    limpiar_registros,
)


def previsualizar_carabineros(csv_path: str):
    csv_path = str(Path(csv_path).expanduser().resolve())
    registros = cargar_registros_csv(csv_path)
    return registros[:20]


def exportar_registros_limpios_csv(registros, ruta_salida: str) -> str:
    filas = []
    for r in registros:
        filas.append({
            "rit": getattr(r, "rit", ""),
            "anio": getattr(r, "anio", ""),
            "id_notificacion": getattr(r, "id_notificacion", ""),
            "hora": getattr(r, "hora", ""),
            "codigo": getattr(r, "codigo", ""),
            "observacion": getattr(r, "observacion", ""),
        })

    df = pd.DataFrame(filas)
    df.to_csv(ruta_salida, index=False, encoding="utf-8-sig")
    return ruta_salida


def limpiar_csv_carabineros(csv_path: str) -> str:
    csv_path = Path(csv_path).expanduser().resolve()

    registros = cargar_registros_csv(str(csv_path))

    filas = []
    vistos = set()

    for r in registros:
        rit = str(getattr(r, "rit", "")).strip()
        anio = str(getattr(r, "anio", "")).strip()
        id_notificacion = str(getattr(r, "id_notificacion", "")).strip()
        hora = str(getattr(r, "hora", "")).strip()
        codigo = str(getattr(r, "codigo", "")).strip().upper()
        observacion = str(getattr(r, "observacion", "")).strip()

        if not rit or not anio or not id_notificacion:
            continue

        clave = (rit, anio, id_notificacion, hora, codigo, observacion)
        if clave in vistos:
            continue
        vistos.add(clave)

        filas.append({
            "rit": rit,
            "anio": anio,
            "id_notificacion": id_notificacion,
            "hora": hora,
            "codigo": codigo,
            "observacion": observacion,
        })

    salida = csv_path.with_name(f"{csv_path.stem}_limpio.csv")
    df = pd.DataFrame(filas)
    df.to_csv(salida, index=False, encoding="utf-8-sig")

    return str(salida)


def generar_csv_cinj_desde_excel(ruta_entrada: str, hora: str, codigo: str) -> str:
    ruta_entrada = Path(ruta_entrada).expanduser().resolve()
    ext = ruta_entrada.suffix.lower()

    hora = str(hora).strip()
    codigo = str(codigo).strip().upper()

    if not hora.isdigit() or len(hora) != 4:
        raise ValueError("La hora debe tener formato HHMM, por ejemplo 1205.")

    if not codigo:
        raise ValueError("Debes indicar un código, por ejemplo D2.")

    if ext == ".csv":
        df = pd.read_csv(ruta_entrada, dtype=str, encoding="utf-8-sig").fillna("")
        df.columns = [str(c).strip().lower() for c in df.columns]

        for col in ["rit", "anio", "id_notificacion"]:
            if col not in df.columns:
                raise ValueError(f"El archivo CSV no contiene la columna requerida: {col}")

        salida_df = pd.DataFrame({
            "rit": df["rit"].astype(str).str.strip(),
            "anio": df["anio"].astype(str).str.strip(),
            "id_notificacion": df["id_notificacion"].astype(str).str.strip(),
            "observacion": "",
            "hora": hora,
            "codigo": codigo,
        })

    else:
        # Para tu archivo DetalleImpresion, la fila de encabezados suele venir en la fila 6
        df = pd.read_excel(ruta_entrada, header=6, dtype=str).fillna("")
        df.columns = [str(c).strip().upper() for c in df.columns]
        print(df.columns)

        for col in ["RIT", "AÑO", "ID NOTIFICACIÓN"]:
            if col not in df.columns:
                raise ValueError(f"No se encontró la columna requerida en el Excel: {col}")

        salida_df = pd.DataFrame({
            "rit": df["RIT"].astype(str).str.strip(),
            "anio": df["AÑO"].astype(str).str.strip(),
            "id_notificacion": df["ID NOTIFICACIÓN"].astype(str).str.strip(),
            "observacion": "",
            "hora": hora,
            "codigo": codigo,
        })

    salida_df = salida_df[
        (salida_df["rit"] != "") &
        (salida_df["anio"] != "") &
        (salida_df["id_notificacion"] != "")
    ].copy()

    salida_df = salida_df.drop_duplicates(
        subset=["rit", "anio", "id_notificacion"],
        keep="first"
    )

    salida = ruta_entrada.with_name(f"{ruta_entrada.stem}_cinj.csv")
    salida_df.to_csv(salida, index=False, encoding="utf-8-sig")

    return str(salida)


def limpiar_en_cinj_carabineros(csv_path: str) -> str:
    csv_path = Path(csv_path).expanduser().resolve()

    driver = crear_driver()
    try:
        registros = cargar_registros_csv(str(csv_path))
        resultados = limpiar_registros(driver, registros)

        salida = csv_path.with_name(f"{csv_path.stem}_limpieza_cinj.csv")
        exportar_resultados(resultados, str(salida))

        return str(salida)
    finally:
        driver.quit()


def ejecutar_carabineros(csv_path: str, fecha_certificacion: str | None = None) -> str:
    csv_path = str(Path(csv_path).expanduser().resolve())

    driver = crear_driver()
    try:
        registros = cargar_registros_csv(csv_path)
        resultados = procesar_registros(driver, registros, fecha_certificacion)

        salida = Path(csv_path).with_name(f"{Path(csv_path).stem}_resultado.csv")
        exportar_resultados(resultados, str(salida))

        return str(salida)
    finally:
        driver.quit()
