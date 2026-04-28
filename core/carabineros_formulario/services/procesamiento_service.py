from dataclasses import asdict
import os
import pandas as pd
import time

from core.carabineros_formulario.config.settings import PAUSA_CORTA, PAUSA_MEDIA, PAUSA_LARGA
from core.carabineros_formulario.data.models import ResultadoProceso
from core.carabineros_formulario.pages.login_page import LoginPage
from core.carabineros_formulario.pages.certificaciones_page import CertificacionesPage
from core.carabineros_formulario.pages.caratula_page import CaratulaPage
from core.carabineros_formulario.utils.logger import log_error, log_info, log_ok, log_warn
from core.carabineros_formulario.config.settings import (
    CINJ_URL,
    CINJ_USER,
    CINJ_PASS,
    TIMEOUT,
    OUTPUT_RESULTADOS_DIR,
)


def exportar_resultados(resultados: list[ResultadoProceso], nombre_archivo: str = "resultados.csv") -> str:
    os.makedirs(OUTPUT_RESULTADOS_DIR, exist_ok=True)
    path = os.path.join(OUTPUT_RESULTADOS_DIR, nombre_archivo)

    df = pd.DataFrame([asdict(r) for r in resultados])
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return path


def procesar_registros(driver, registros, fecha_certificacion: str | None = None) -> list[ResultadoProceso]:
    resultados: list[ResultadoProceso] = []

    login_page = LoginPage(driver, TIMEOUT)
    cert_page = CertificacionesPage(driver, TIMEOUT)
    caratula_page = CaratulaPage(driver, TIMEOUT)

    log_info("Abriendo login")
    login_page.abrir(CINJ_URL)

    log_info("Iniciando sesión")
    login_page.login(CINJ_USER, CINJ_PASS)
    log_ok("Login correcto")

    log_info("Navegando a certificaciones")
    cert_page.ir_a_certificaciones()
    log_ok("Navegación a certificaciones completa")

    log_info("Fecha de certificación a usar: " + (fecha_certificacion or "hoy"))
    cert_page.establecer_fecha_certificacion(fecha_certificacion)
    log_ok("Fecha de certificación ingresada")

    for i, reg in enumerate(registros, start=1):
        log_info(f"Procesando registro {i}: RIT={reg.rit}, AÑO={reg.anio}, ID={reg.id_notificacion}")

        try:
            # CAMBIO: ahora busca directo por ID de notificación
            cert_page.buscar_notificaciones_por_id(reg.id_notificacion)

            fila = cert_page.buscar_fila_por_id(reg.id_notificacion)

            if fila is None:
                log_warn("No se encontró fila con ID coincidente")
                resultados.append(
                    ResultadoProceso(
                        rit=reg.rit,
                        anio=reg.anio,
                        id_notificacion=reg.id_notificacion,
                        hora=reg.hora,
                        codigo=reg.codigo,
                        observacion=reg.observacion,
                        estado="no_encontrado",
                        mensaje="No se encontró ID_NOTIFICACION en el listado"
                    )
                )
                continue

            # (reservas manejadas tras guardar la carátula)

            cert_page.abrir_caratula_de_fila(fila)
            gestion = caratula_page.buscar_primera_gestion_disponible()

            if gestion is None:
                log_warn("No hay gestión disponible")
                caratula_page.cerrar()
                resultados.append(
                    ResultadoProceso(
                        rit=reg.rit,
                        anio=reg.anio,
                        id_notificacion=reg.id_notificacion,
                        hora=reg.hora,
                        codigo=reg.codigo,
                        observacion=reg.observacion,
                        estado="sin_gestion_disponible",
                        mensaje="Las 3 gestiones están ocupadas o deshabilitadas"
                    )
                )
                continue

            caratula_page.ingresar_datos_en_gestion(
                gestion,
                reg.hora,
                reg.codigo,
                reg.observacion
            )
            time.sleep(PAUSA_MEDIA)

            resultado_guardado = caratula_page.guardar()

            # Si aparece el modal de reserva, cerrarlo solo con el botón de guardar (2)
            try:
                if resultado_guardado == "modal_reserva":
                    log_info("Modal de reserva detectado tras guardar carátula. Intentando cerrar con botón guardar...")
                    cerrado_modal_global = cert_page.guardar_y_cerrar_modal_reserva()
                    if cerrado_modal_global:
                        log_ok("Modal de reserva cerrado automáticamente")
                        time.sleep(PAUSA_CORTA)
                        # si la carátula ya se cerró como consecuencia del modal, considerarlo ok
                        if not caratula_page.modal_sigue_abierto():
                            resultado_guardado = "ok"
                    else:
                        log_warn("No se pudo cerrar el modal de reserva automáticamente tras guardar")
            except Exception as e:
                log_warn(f"Error intentando cerrar modal de reserva tras guardar: {e}")

            if resultado_guardado == "ok":
                log_ok(f"Registro guardado en gestión {gestion}")
                resultados.append(
                    ResultadoProceso(
                        rit=reg.rit,
                        anio=reg.anio,
                        id_notificacion=reg.id_notificacion,
                        hora=reg.hora,
                        codigo=reg.codigo,
                        observacion=reg.observacion,
                        estado="ok",
                        mensaje="Guardado correctamente",
                        gestion_usada=gestion
                    )
                )
                continue

            if resultado_guardado == "no_certificable":
                log_warn(f"Registro no certificable: {reg.id_notificacion}")
                resultados.append(
                    ResultadoProceso(
                        rit=reg.rit,
                        anio=reg.anio,
                        id_notificacion=reg.id_notificacion,
                        hora=reg.hora,
                        codigo=reg.codigo,
                        observacion=reg.observacion,
                        estado="no_certificable",
                        mensaje="No se pudo certificar por lógica de negocio o flujo del sistema",
                        gestion_usada=gestion
                    )
                )
                continue

            if resultado_guardado == "modal_reserva":
                log_warn(f"Modal de reserva pendiente para: {reg.id_notificacion}")
                resultados.append(
                    ResultadoProceso(
                        rit=reg.rit,
                        anio=reg.anio,
                        id_notificacion=reg.id_notificacion,
                        hora=reg.hora,
                        codigo=reg.codigo,
                        observacion=reg.observacion,
                        estado="reservado",
                        mensaje="La carátula abrió modal de reserva"
                    )
                )
                continue

            if resultado_guardado == "modal_abierto":
                log_warn("La carátula sigue abierta. Esperando cierre manual...")
                while caratula_page.modal_sigue_abierto():
                    time.sleep(PAUSA_LARGA)

                resultados.append(
                    ResultadoProceso(
                        rit=reg.rit,
                        anio=reg.anio,
                        id_notificacion=reg.id_notificacion,
                        hora=reg.hora,
                        codigo=reg.codigo,
                        observacion=reg.observacion,
                        estado="guardado_manual",
                        mensaje="La carátula requirió cierre/gestión manual",
                        gestion_usada=gestion
                    )
                )
                continue

        except Exception as e:
            log_error(f"Error procesando registro {reg.id_notificacion}: {e}")
            resultados.append(
                ResultadoProceso(
                    rit=reg.rit,
                    anio=reg.anio,
                    id_notificacion=reg.id_notificacion,
                    hora=reg.hora,
                    codigo=reg.codigo,
                    observacion=reg.observacion,
                    estado="error",
                    mensaje=str(e)
                )
            )

    return resultados


def limpiar_registros(driver, registros) -> list[ResultadoProceso]:
    resultados: list[ResultadoProceso] = []

    login_page = LoginPage(driver, TIMEOUT)
    cert_page = CertificacionesPage(driver, TIMEOUT)

    log_info("Abriendo login")
    login_page.abrir(CINJ_URL)

    log_info("Iniciando sesión")
    login_page.login(CINJ_USER, CINJ_PASS)
    log_ok("Login correcto")

    log_info("Navegando a certificaciones")
    cert_page.ir_a_certificaciones()
    log_ok("Navegación a certificaciones completa")

    for i, reg in enumerate(registros, start=1):
        log_info(f"Limpiando registro {i}: RIT={reg.rit}, AÑO={reg.anio}, ID={reg.id_notificacion}")

        try:
            # Mantengo la llamada igual para no romper nada,
            # pero limpiar_notificacion_por_id ya debería usar búsqueda por ID internamente
            estado_limpieza = cert_page.limpiar_notificacion_por_id(
                reg.rit,
                reg.anio,
                reg.id_notificacion
            )

            if estado_limpieza == "no_encontrado":
                log_warn("No se encontró fila con ID coincidente para limpiar")
                resultados.append(
                    ResultadoProceso(
                        rit=reg.rit,
                        anio=reg.anio,
                        id_notificacion=reg.id_notificacion,
                        hora=reg.hora,
                        codigo=reg.codigo,
                        observacion=reg.observacion,
                        estado="no_encontrado",
                        mensaje="No se encontró ID_NOTIFICACION para limpiar"
                    )
                )
                continue

            if estado_limpieza == "ya_limpio":
                log_warn(f"Registro ya estaba limpio: {reg.id_notificacion}")
                resultados.append(
                    ResultadoProceso(
                        rit=reg.rit,
                        anio=reg.anio,
                        id_notificacion=reg.id_notificacion,
                        hora=reg.hora,
                        codigo=reg.codigo,
                        observacion=reg.observacion,
                        estado="ya_limpio",
                        mensaje="El registro ya estaba limpio"
                    )
                )
                continue

            if estado_limpieza == "limpiado":
                log_ok(f"Registro limpiado: {reg.id_notificacion}")
                resultados.append(
                    ResultadoProceso(
                        rit=reg.rit,
                        anio=reg.anio,
                        id_notificacion=reg.id_notificacion,
                        hora=reg.hora,
                        codigo=reg.codigo,
                        observacion=reg.observacion,
                        estado="ok",
                        mensaje="Limpieza realizada correctamente"
                    )
                )
                continue

        except Exception as e:
            log_error(f"Error limpiando registro {reg.id_notificacion}: {e}")
            resultados.append(
                ResultadoProceso(
                    rit=reg.rit,
                    anio=reg.anio,
                    id_notificacion=reg.id_notificacion,
                    hora=reg.hora,
                    codigo=reg.codigo,
                    observacion=reg.observacion,
                    estado="error",
                    mensaje=str(e)
                )
            )

    return resultados