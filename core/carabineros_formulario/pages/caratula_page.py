import time
from core.carabineros_formulario.config.settings import PAUSA_CORTA, PAUSA_MEDIA
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from core.carabineros_formulario.locators.caratula_locators import CaratulaLocators
from core.carabineros_formulario.utils.time_utils import separar_hora_minuto


class CaratulaPage:
    def __init__(self, driver, timeout: int):
        self.driver = driver
        self.wait = WebDriverWait(driver, timeout)

    def obtener_dialog_abierto(self):
        return self.wait.until(
            EC.visibility_of_element_located(CaratulaLocators.DIALOG)
        )

    def obtener_bloque_gestion(self, n: int):
        dialog = self.obtener_dialog_abierto()
        return dialog.find_element(*CaratulaLocators.bloque_gestion(n))

    def gestion_disponible(self, n: int) -> bool:
        try:
            dialog = self.obtener_dialog_abierto()

            textarea = dialog.find_element(*CaratulaLocators.textarea_observacion(n))
            input_codigo = dialog.find_element(*CaratulaLocators.input_codigo(n))
            select_hora = dialog.find_element(*CaratulaLocators.select_hora(n))
            select_minuto = dialog.find_element(*CaratulaLocators.select_minuto(n))

            return (
                textarea.is_enabled()
                and input_codigo.is_enabled()
                and select_hora.is_enabled()
                and select_minuto.is_enabled()
            )
        except Exception:
            return False

    def buscar_primera_gestion_disponible(self) -> int | None:
        for n in (1, 2, 3):
            if self.gestion_disponible(n):
                return n
        return None

    def ingresar_datos_en_gestion(self, n: int, hora: str, codigo: str, observacion: str) -> None:
        dialog = self.obtener_dialog_abierto()
        hh, mm = separar_hora_minuto(hora)

        print("DEBUG hora original:", repr(hora))
        print("DEBUG hh:", hh)
        print("DEBUG mm:", mm)
        print("DEBUG codigo:", repr(codigo))
        print("DEBUG observacion:", repr(observacion))
        print("DEBUG gestion n:", n)

        select_hora_elem = dialog.find_element(*CaratulaLocators.select_hora(n))
        select_minuto_elem = dialog.find_element(*CaratulaLocators.select_minuto(n))
        input_codigo = dialog.find_element(*CaratulaLocators.input_codigo(n))
        textarea = dialog.find_element(*CaratulaLocators.textarea_observacion(n))

        print("DEBUG select hora enabled:", select_hora_elem.is_enabled())
        print("DEBUG select minuto enabled:", select_minuto_elem.is_enabled())
        print("DEBUG input codigo enabled:", input_codigo.is_enabled())
        print("DEBUG textarea enabled:", textarea.is_enabled())

        # 1) OBSERVACION
        textarea.click()
        textarea.clear()
        time.sleep(0.1)
        textarea.send_keys(observacion)
        time.sleep(PAUSA_CORTA)

        observacion_final = textarea.get_attribute("value").strip()
        print("DEBUG observacion final:", repr(observacion_final))

        if observacion_final != str(observacion).strip():
            raise Exception(
                f"No se pudo fijar la observación. Esperado={observacion!r}, actual={observacion_final!r}"
            )

        # 2) CODIGO
        input_codigo.click()
        input_codigo.clear()
        time.sleep(0.1)
        input_codigo.send_keys(codigo)
        time.sleep(PAUSA_CORTA)

        codigo_final = input_codigo.get_attribute("value").strip()
        print("DEBUG codigo final:", repr(codigo_final))

        if codigo_final != str(codigo).strip():
            raise Exception(
                f"No se pudo fijar el código. Esperado={codigo!r}, actual={codigo_final!r}"
            )

        # 3) HORA / MINUTO AL FINAL
        select_hora = Select(select_hora_elem)
        select_minuto = Select(select_minuto_elem)

        print("DEBUG opciones hora:", [o.text.strip() for o in select_hora.options])
        print("DEBUG opciones minuto:", [o.text.strip() for o in select_minuto.options])

        print("DEBUG hora antes:", select_hora.first_selected_option.text.strip())
        print("DEBUG minuto antes:", select_minuto.first_selected_option.text.strip())

        # tocar hora/minuto puede abrir diálogo historial
        try:
            select_hora_elem.click()
            time.sleep(0.2)
        except Exception:
            pass

        if self.historial_dialog_abierto():
            print("DEBUG apareció diálogo historial tras tocar hora")
            self.cerrar_historial_si_aparece()

        try:
            select_minuto_elem.click()
            time.sleep(0.2)
        except Exception:
            pass

        if self.historial_dialog_abierto():
            print("DEBUG apareció diálogo historial tras tocar minuto")
            self.cerrar_historial_si_aparece()

        # Fijar ambos por JS para que el frontend no los pise entre una selección y otra
        self.driver.execute_script("""
            const selectHora = arguments[0];
            const selectMin = arguments[1];
            const hh = arguments[2];
            const mm = arguments[3];

            const setByText = (sel, txt) => {
                const opt = [...sel.options].find(o => o.text.trim() === txt);
                if (!opt) return false;
                sel.value = opt.value;
                sel.dispatchEvent(new Event('input', { bubbles: true }));
                sel.dispatchEvent(new Event('change', { bubbles: true }));
                return true;
            };

            const okHora = setByText(selectHora, hh);
            const okMin = setByText(selectMin, mm);

            selectHora.dispatchEvent(new Event('blur', { bubbles: true }));
            selectMin.dispatchEvent(new Event('blur', { bubbles: true }));

            return [okHora, okMin];
        """, select_hora_elem, select_minuto_elem, hh, mm)

        time.sleep(PAUSA_MEDIA)

        hora_final = Select(select_hora_elem).first_selected_option.text.strip()
        minuto_final = Select(select_minuto_elem).first_selected_option.text.strip()

        print("DEBUG hora después:", hora_final)
        print("DEBUG minuto después:", minuto_final)

        if hora_final != hh:
            raise Exception(f"No se pudo fijar la hora. Esperado={hh}, actual={hora_final}")

        if minuto_final != mm:
            raise Exception(f"No se pudo fijar el minuto. Esperado={mm}, actual={minuto_final}")

        # IMPORTANTE:
        # después de esto NO hacer click en código, textarea, dialog ni nada.
        # el siguiente paso debe ser SOLO guardar.
        
    def guardar(self) -> str:
        dialog = self.obtener_dialog_abierto()
        btn = dialog.find_element(*CaratulaLocators.BOTON_GUARDAR)
        time.sleep(PAUSA_CORTA)

        print("DEBUG botón guardar texto:", btn.text)
        print("DEBUG botón guardar enabled:", btn.is_enabled())
        print("DEBUG botón guardar displayed:", btn.is_displayed())
        print("DEBUG botón guardar HTML:", btn.get_attribute("outerHTML"))

        self.driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            btn
        )
        time.sleep(0.2)

        try:
            btn.click()
            time.sleep(PAUSA_MEDIA)
            print("DEBUG click normal en guardar ejecutado")
        except Exception as e:
            print("DEBUG falló click normal:", e)
            try:
                self.driver.execute_script("arguments[0].click();", btn)
                time.sleep(PAUSA_MEDIA)
                print("DEBUG click JS en guardar ejecutado")
            except Exception as e2:
                print("DEBUG falló click JS:", e2)
                return "modal_abierto"

        # si apareció el modal de reserva, NO intentar cerrar alertas genéricas aquí
        if self.alerta_modal_abierta():
            print("DEBUG apareció alerta modal después de guardar")
            return "modal_reserva"

        print("DEBUG modal sigue abierto después de guardar:", self.modal_sigue_abierto())

        if not self.modal_sigue_abierto():
            return "ok"

        return "modal_abierto"

    def cerrar(self) -> None:
        dialog = self.obtener_dialog_abierto()
        print(dialog)
        btn = dialog.find_element(*CaratulaLocators.BOTON_CERRAR)
        self.driver.execute_script("arguments[0].click();", btn)

        WebDriverWait(self.driver, 10).until(
            EC.invisibility_of_element_located(CaratulaLocators.DIALOG)
        )

    def modal_sigue_abierto(self) -> bool:
        dialogs = self.driver.find_elements(*CaratulaLocators.DIALOG)
        return any(d.is_displayed() for d in dialogs)


    def historial_dialog_abierto(self) -> bool:
        dialogs = self.driver.find_elements(*CaratulaLocators.HISTORIAL_DIALOG)
        return any(d.is_displayed() for d in dialogs)


    def cerrar_historial_si_aparece(self) -> bool:
        try:
            dialogs = self.driver.find_elements(*CaratulaLocators.HISTORIAL_DIALOG)
            visibles = [d for d in dialogs if d.is_displayed()]

            if not visibles:
                return False

            btn = self.wait.until(
                EC.element_to_be_clickable(CaratulaLocators.BOTON_CERRAR_HISTORIAL)
            )

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
            return True

        except Exception:
            return False
    def alerta_modal_abierta(self) -> bool:
        dialogs = self.driver.find_elements(*CaratulaLocators.ALERTA_MODAL)
        return any(d.is_displayed() for d in dialogs)


    def cerrar_alerta_modal_si_aparece(self) -> bool:
        try:
            dialogs = self.driver.find_elements(*CaratulaLocators.ALERTA_MODAL)
            visibles = [d for d in dialogs if d.is_displayed()]
            if not visibles:
                return False

            botones = self.driver.find_elements(*CaratulaLocators.BOTON_CERRAR_ALERTA)
            btn = next((b for b in botones if b.is_displayed() and b.is_enabled()), None)

            if btn is None:
                return False

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
            return True

        except Exception:
            return False
    def cerrar_todas_las_alertas_si_aparecen(self) -> int:
        total_cerradas = 0
        max_intentos = 5

        for _ in range(max_intentos):
            if not self.alerta_modal_abierta():
                break

            cerrada = self.cerrar_alerta_modal_si_aparece()
            if not cerrada:
                break

            total_cerradas += 1
            time.sleep(PAUSA_MEDIA)

        return total_cerradas
