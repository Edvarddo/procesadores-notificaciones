import pandas as pd
from core.carabineros_formulario.data.models import RegistroNotificacion


def cargar_registros_csv(path: str) -> list[RegistroNotificacion]:
    df = pd.read_csv(path, dtype=str).fillna("")

    # normalizar columnas
    df.columns = [str(c).strip().lower() for c in df.columns]

    # mapear variantes
    rename_map = {
        "id_notificacion": "id_notificacion",
        "id notificacion": "id_notificacion",
        "id_notificación": "id_notificacion",
        "hora": "hora",
        "codigo": "codigo",
        "código": "codigo",
        "observacion": "observacion",
        "observación": "observacion",
    }
    df = df.rename(columns=rename_map)

    # SOLO requerimos ID
    if "id_notificacion" not in df.columns:
        raise Exception("El archivo debe tener la columna 'id_notificacion'")

    registros: list[RegistroNotificacion] = []

    for _, row in df.iterrows():
        registros.append(
            RegistroNotificacion(
                rit=0,  # ya no se usa
                anio=0,  # ya no se usa
                id_notificacion=str(row.get("id_notificacion", "")).strip(),
                hora=str(row.get("hora", "")).strip(),
                codigo=str(row.get("codigo", "")).strip(),
                observacion=str(row.get("observacion", "")).strip(),
            )
        )

    return registros