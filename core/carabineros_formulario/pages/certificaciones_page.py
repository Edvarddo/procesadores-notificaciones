from datetime import date, datetime

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException
import time
from core.carabineros_formulario.config.settings import PAUSA_CORTA, PAUSA_MEDIA
from core.carabineros_formulario.locators.certificaciones_locators import CertificacionesLocators


class CertificacionesPage:
    def __init__(self, driver, timeout: int):
        self.driver = driver
        self.wait = WebDriverWait(driver, timeout)

    def _click(self, locator):
        elem = self.wait.until(EC.element_to_be_clickable(locator))
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", elem
        )
        elem.click()
        return elem

    def _write(self, locator, value):
        elem = self.wait.until(EC.element_to_be_clickable(locator))
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", elem
        )
        elem.clear()
        elem.send_keys(str(value))
        return elem

    def ir_a_certificaciones(self) -> None:
        self._click(CertificacionesLocators.MENU)
        self._click(CertificacionesLocators.PROCESOS)
        self._click(CertificacionesLocators.CERTIFICACION)
        self._click(CertificacionesLocators.CERTIFICACIONES)

    def construir_rango_fechas_actual(self) -> str:
        hoy = date.today()
        inicio_anio = date(hoy.year, 1, 1)
        return f"{inicio_anio:%d/%m/%Y} - {hoy:%d/%m/%Y}"

    def buscar_notificaciones(self, rit: int, anio: int, rango_fechas: str | None = None) -> None:
        if rango_fechas is None:
            rango_fechas = self.construir_rango_fechas_actual()

        campo_fecha = self.wait.until(
            EC.element_to_be_clickable(CertificacionesLocators.FECHA)
        )

        print("🔍 Campo fecha encontrado")
        print("🟡 Valor ANTES:", campo_fecha.get_attribute("value"))
        print("🟡 Se intentará escribir:", rango_fechas)

        self.driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            campo_fecha
        )

        # dar foco
        #campo_fecha.click()
        #time.sleep(0.2)

        # seleccionar todo en Mac
        campo_fecha.send_keys(Keys.COMMAND, "a")
        time.sleep(0.2)
        campo_fecha.send_keys(Keys.BACKSPACE)
        time.sleep(0.2)

        # forzar limpieza real en Angular/Material
        self.driver.execute_script("""
            arguments[0].value = '';
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, campo_fecha)

        time.sleep(0.2)

        # escribir nuevo rango
        campo_fecha.send_keys(rango_fechas)

        # disparar eventos para que Angular lo procese
        self.driver.execute_script("""
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
        """, campo_fecha)

        time.sleep(0.3)

        valor_final = campo_fecha.get_attribute("value")
        print("🟢 Valor DESPUÉS:", valor_final)
        print("🟢 HTML:", campo_fecha.get_attribute("outerHTML"))

        if rango_fechas not in valor_final:
            raise Exception(f"No se pudo setear correctamente la fecha. Valor final: {valor_final}")

        self._write(CertificacionesLocators.RIT, rit)
        self._write(CertificacionesLocators.ANIO, anio)
        time.sleep(PAUSA_CORTA)
        self._click(CertificacionesLocators.BUSCAR)
        time.sleep(PAUSA_MEDIA)

    def obtener_filas_resultado(self):
        self.wait.until(
            EC.presence_of_all_elements_located(CertificacionesLocators.FILAS_RESULTADO)
        )
        return self.driver.find_elements(*CertificacionesLocators.FILAS_RESULTADO)

    def obtener_id_de_fila(self, fila) -> str:
        celda = fila.find_element(*CertificacionesLocators.CELDA_ID_REL)
        return celda.text.strip()

    def buscar_fila_por_id(self, id_notificacion: str):
        filas = self.obtener_filas_resultado()

        for fila in filas:
            id_web = self.obtener_id_de_fila(fila)
            if id_web == str(id_notificacion).strip():
                return fila

        return None

    def abrir_caratula_de_fila(self, fila) -> None:
        # Buscar el icono de carátula y su botón contenedor
        icon = None
        try:
            icon = fila.find_element(*CertificacionesLocators.BTN_CARATULA_REL)
        except Exception:
            try:
                icon = fila.find_element(By.XPATH, ".//td[2]//mat-icon | .//td[2]//button")
            except Exception:
                print("Error: no se encontró el icono/botón de carátula en la fila")
                raise

        # intentar obtener el botón padre si el icono está dentro de uno
        btn = None
        try:
            btn = icon.find_element(By.XPATH, "ancestor::button[1]")
        except Exception:
            btn = icon

        # Intentar varios métodos de click con reintentos
        last_exc = None
        for intento in range(3):
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                time.sleep(0.1)
                try:
                    btn.click()
                except Exception:
                    try:
                        self.driver.execute_script("arguments[0].click();", btn)
                    except Exception:
                        try:
                            btn.send_keys("\n")
                        except Exception as e:
                            last_exc = e

                # esperar a que aparezca el diálogo de carátula
                try:
                    WebDriverWait(self.driver, 6).until(
                        EC.visibility_of_element_located((By.XPATH, "//app-caratula-penal"))
                    )
                    time.sleep(PAUSA_MEDIA)
                    return
                except Exception:
                    # no apareció aún, esperar y reintentar
                    time.sleep(0.5 + intento * 0.5)
                    continue

            except Exception as e:
                last_exc = e
                time.sleep(0.5)

        # si llegamos aquí, no se abrió la carátula
        try:
            ts = int(time.time())
            path = f"/tmp/caratula_error_{ts}.png"
            self.driver.get_screenshot_as_file(path)
            print(f"[ERROR] No apareció la carátula tras intentar abrirla. Captura en: {path}")
        except Exception:
            pass

        if last_exc:
            raise Exception(f"No se abrió la carátula: {last_exc}")
        raise Exception("No se abrió la carátula: desconocido")

    def limpiar_fila(self, fila) -> str:
        btn = fila.find_element(*CertificacionesLocators.BTN_LIMPIAR_REL())
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            btn
        )
        time.sleep(0.2)

        try:
            btn.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", btn)

        time.sleep(0.5)

        # si aparece alerta, cerrarla y reportar que ya estaba limpio
        if self.cerrar_alerta_modal_si_aparece():
            return "ya_limpio"

        return "limpiado"


    def limpiar_notificacion_por_id(self, rit: int, anio: int, id_notificacion: str) -> str:
        self.buscar_notificaciones(rit, anio)

        fila = self.buscar_fila_por_id(id_notificacion)
        if fila is None:
            return "no_encontrado"

        return self.limpiar_fila(fila)

    def construir_fecha_certificacion_actual(self) -> str:
        hoy = date.today()
        return f"{hoy:%d/%m/%Y}"

    def establecer_fecha_certificacion(self, fecha_certificacion: str | None = None) -> None:
        if not fecha_certificacion:
            fecha_certificacion = self.construir_fecha_certificacion_actual()

        campo_fecha = self.wait.until(
            EC.element_to_be_clickable(CertificacionesLocators.FECHA_CERTIFICACION)
        )

        valor_actual = campo_fecha.get_attribute("value").strip()

        print("📌 Fecha certificación actual:", valor_actual)
        print("📌 Se intentará escribir fecha certificación:", fecha_certificacion)

        # si ya está correcta, no tocarla
        if valor_actual == fecha_certificacion:
            print("📌 La fecha de certificación ya está correcta, no se modifica.")
            return

        self.driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            campo_fecha
        )
        time.sleep(0.2)

        campo_fecha.click()
        time.sleep(0.2)

        # MAC
        campo_fecha.send_keys(Keys.COMMAND, "a")
        time.sleep(0.2)
        campo_fecha.send_keys(Keys.BACKSPACE)
        time.sleep(0.2)

        # limpieza forzada
        self.driver.execute_script("""
            arguments[0].value = '';
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, campo_fecha)

        time.sleep(0.2)

        # escribir la nueva fecha
        campo_fecha.send_keys(fecha_certificacion)
        time.sleep(0.2)

        # disparar eventos para Angular
        self.driver.execute_script("""
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
        """, campo_fecha)

        time.sleep(0.3)

        valor_final = campo_fecha.get_attribute("value").strip()
        print("📌 Fecha certificación final:", valor_final)

        if valor_final != fecha_certificacion:
            raise Exception(
                f"No se pudo setear la fecha de certificación. "
                f"Esperado={fecha_certificacion}, actual={valor_final}"
            )
            if not fecha_certificacion:
                fecha_certificacion = self.construir_fecha_certificacion_actual()

            campo_fecha = self.wait.until(
                EC.element_to_be_clickable(CertificacionesLocators.FECHA_CERTIFICACION)
            )

            print("📌 Fecha certificación actual:", campo_fecha.get_attribute("value"))
            print("📌 Se intentará escribir fecha certificación:", fecha_certificacion)

            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                campo_fecha
            )
            time.sleep(0.2)

            campo_fecha.click()
            time.sleep(0.2)

            campo_fecha.send_keys(Keys.COMMAND, "a")
            time.sleep(0.2)
            campo_fecha.send_keys(Keys.BACKSPACE)
            time.sleep(0.2)

            self.driver.execute_script("""
                arguments[0].value = '';
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, campo_fecha)

            time.sleep(0.2)
            campo_fecha.send_keys(fecha_certificacion)

            self.driver.execute_script("""
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
            """, campo_fecha)

            time.sleep(0.3)

            valor_final = campo_fecha.get_attribute("value").strip()
            print("📌 Fecha certificación final:", valor_final)

            if fecha_certificacion not in valor_final:
                raise Exception(
                    f"No se pudo setear la fecha de certificación. Valor final: {valor_final}"
                )
    def alerta_modal_abierta(self) -> bool:
        dialogs = self.driver.find_elements(*CertificacionesLocators.ALERTA_MODAL)
        return any(d.is_displayed() for d in dialogs)


    def cerrar_alerta_modal_si_aparece(self) -> bool:
        try:
            WebDriverWait(self.driver, 2).until(
                EC.visibility_of_element_located(CertificacionesLocators.ALERTA_MODAL)
            )

            return self.guardar_y_cerrar_modal_reserva()

        except Exception:
            return False

    def guardar_y_cerrar_modal_reserva(self) -> bool:
        """Cierra el modal de alerta guardando (botón derecho/guardar)"""
        try:
            # Espera a que el modal esté visible (usar un wait algo mayor)
            WebDriverWait(self.driver, 6).until(
                EC.visibility_of_element_located(CertificacionesLocators.ALERTA_MODAL)
            )

            # Usar únicamente el botón 2 del modal de reserva
            btn = WebDriverWait(self.driver, 4).until(
                EC.element_to_be_clickable(CertificacionesLocators.BOTON_GUARDAR_RESERVA)
            )

            if not btn:
                return False

            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                btn
            )
            time.sleep(0.2)

            try:
                btn.click()
            except WebDriverException:
                try:
                    self.driver.execute_script("arguments[0].click();", btn)
                except Exception:
                    pass

            # esperar que el modal desaparezca
            try:
                WebDriverWait(self.driver, 6).until(
                    EC.invisibility_of_element_located(CertificacionesLocators.ALERTA_MODAL)
                )
            except Exception:
                # si no desaparece, devolver False pero no lanzar
                return False

            time.sleep(0.3)
            return True

        except WebDriverException as e:
            try:
                timestamp = int(time.time())
                path = f"/tmp/reserva_error_{timestamp}.png"
                self.driver.get_screenshot_as_file(path)
                print(f"[ERROR] WebDriverException al cerrar modal. Captura en: {path}")
            except Exception:
                pass
            return False
        except Exception:
            return False

    def es_notificacion_reservada(self, fila) -> bool:
        """Verifica si la fila de notificación está marcada como reservada"""
        try:
            celda_id = fila.find_element(*CertificacionesLocators.CELDA_RESERVADO_REL)
            clases = celda_id.get_attribute("class") or ""
            es_reservada = "reservado" in clases.lower() or "reservd" in clases.lower()
            
            if es_reservada:
                print(f"⚠️  Notificación reservada detectada - Clases: {clases}")
            
            return es_reservada
        except Exception as e:
            print(f"Error verificando si es reservada: {e}")
            return False
    def buscar_notificaciones_por_id(self, id_notificacion: str, rango_fechas: str | None = None) -> None:
        if rango_fechas is None:
            rango_fechas = self.construir_rango_fechas_actual()

        campo_fecha = self.wait.until(
            EC.element_to_be_clickable(CertificacionesLocators.FECHA)
        )

        valor_actual = campo_fecha.get_attribute("value").strip()

        print("🔍 Campo fecha encontrado")
        print("🟡 Valor ANTES:", valor_actual)
        print("🟡 Se intentará escribir:", rango_fechas)

        if valor_actual != rango_fechas:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});",
                campo_fecha
            )

            campo_fecha.click()
            time.sleep(0.2)

            campo_fecha.send_keys(Keys.COMMAND, "a")
            time.sleep(0.2)
            campo_fecha.send_keys(Keys.BACKSPACE)
            time.sleep(0.2)

            self.driver.execute_script("""
                arguments[0].value = '';
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, campo_fecha)

            time.sleep(0.2)
            campo_fecha.send_keys(rango_fechas)

            self.driver.execute_script("""
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
            """, campo_fecha)

            time.sleep(PAUSA_MEDIA)

        print("🟢 Valor DESPUÉS:", campo_fecha.get_attribute("value"))

        campo_id = self.wait.until(
            EC.element_to_be_clickable(CertificacionesLocators.ID_NOTIFICACION_CAMPO)
        )

        self.driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            campo_id
        )
        time.sleep(PAUSA_CORTA)

        campo_id.click()
        campo_id.clear()
        time.sleep(0.1)
        campo_id.send_keys(str(id_notificacion))
        time.sleep(PAUSA_CORTA)

        self.driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(PAUSA_CORTA)

        self._click(CertificacionesLocators.BUSCAR)
        time.sleep(PAUSA_MEDIA)

