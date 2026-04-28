from pathlib import Path

from .pdf_notificaciones import procesar_pdf
from .excel_detalle import procesar_excel
from .join_resultados import unir_dataframes


def procesar_pdf_excel(ruta_pdf: str | Path, ruta_excel: str | Path) -> dict:
    archivo_pdf = Path(ruta_pdf)
    archivo_excel = Path(ruta_excel)

    if not archivo_pdf.exists():
        raise FileNotFoundError(f"No existe el PDF: {archivo_pdf}")

    if not archivo_excel.exists():
        raise FileNotFoundError(f"No existe el Excel: {archivo_excel}")

    # 1) Procesar PDF
    df_pdf, no_parseados = procesar_pdf(str(archivo_pdf))
    if df_pdf.empty:
        raise ValueError("No se pudieron extraer registros desde el PDF.")

    # 2) Procesar Excel
    df_excel = procesar_excel(str(archivo_excel))

    # 3) Validaciones
    if "HASH" not in df_pdf.columns:
        raise ValueError("El DataFrame del PDF no contiene la columna HASH")

    if "HASH" not in df_excel.columns:
        raise ValueError("El DataFrame del Excel no contiene la columna HASH")

    # 4) Join
    df_final = unir_dataframes(df_pdf, df_excel)

    # 5) Estadísticas
    con_match = int((df_final["_merge"] == "both").sum()) if "_merge" in df_final.columns else 0
    sin_match = int((df_final["_merge"] == "left_only").sum()) if "_merge" in df_final.columns else 0

    hash_pdf_dup = int(df_pdf["HASH"].duplicated(keep=False).sum()) if "HASH" in df_pdf.columns else 0
    hash_excel_dup = int(df_excel["HASH"].duplicated(keep=False).sum()) if "HASH" in df_excel.columns else 0

    return {
        "df_final": df_final,
        "df_pdf": df_pdf,
        "df_excel": df_excel,
        "no_parseados": no_parseados,
        "nombre_base_pdf": archivo_pdf.stem,
        "estadisticas": {
            "filas_pdf": int(len(df_pdf)),
            "filas_excel": int(len(df_excel)),
            "filas_final": int(len(df_final)),
            "con_match": con_match,
            "sin_match": sin_match,
            "hashes_duplicados_pdf": hash_pdf_dup,
            "hashes_duplicados_excel": hash_excel_dup,
            "bloques_no_parseados": int(len(no_parseados)),
        }
    }
