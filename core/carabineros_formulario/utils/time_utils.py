def separar_hora_minuto(valor) -> tuple[str, str]:
    texto = str(valor).strip()

    # quitar todo excepto dígitos
    solo = "".join(ch for ch in texto if ch.isdigit())

    if not solo:
        raise Exception(f"No se pudo interpretar la hora: {valor}")

    # formato 08:59 o 0859
    if len(solo) == 4:
        hh = solo[:2]
        mm = solo[2:]

    # formato 8:59 o 859
    elif len(solo) == 3:
        hh = "0" + solo[0]
        mm = solo[1:]

    # formato 9:05 o 905
    elif len(solo) == 2:
        raise Exception(f"No se pudo interpretar la hora: {valor}")

    else:
        raise Exception(f"No se pudo interpretar la hora: {valor}")

    if not (0 <= int(hh) <= 23 and 0 <= int(mm) <= 59):
        raise Exception(f"Hora inválida: {valor}")

    return hh, mm