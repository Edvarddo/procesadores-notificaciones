from selenium.webdriver.common.by import By


class CaratulaLocators:
    DIALOG = (By.XPATH, "//app-caratula-penal")
    BOTON_CERRAR = (By.XPATH, ".//div[2]/button[2]")
    BOTON_GUARDAR = (By.XPATH, ".//div[2]/button[2]")
    HISTORIAL_DIALOG = (
        By.XPATH,
        ".//app-historial-certificaciones"
    )

    BOTON_CERRAR_HISTORIAL = (
        By.XPATH, 
        ".//app-historial-certificaciones//button[contains(., 'Cerrar') or contains(., 'Cancelar')]"
        "| .//app-historial-certificaciones//div[3]/button[1] "
        "| .//app-historial-certificaciones//div[4]/button[1]"
    )
    ALERTA_MODAL = (By.XPATH, "//app-alerta-modal")
    BOTON_CERRAR_ALERTA = (
        By.XPATH,
	".//app-alerta-modal/div[2]//button[contains(., 'Guardar Certificación') or contains(., 'Guardar')]"
        " | //app-alerta-modal/div[2]/button[2]"
	" | //app-alerta-modal/div[2]/button"
        " |//app-alerta-modal//button[contains(., 'Cerrar') or contains(., 'Aceptar') or contains(., 'OK')]"
        " | //*[@id[starts-with(., 'mat-dialog-')]]/app-alerta-modal/div[2]/button"
    )

    @staticmethod
    def bloque_gestion(n: int):
        return (By.XPATH, f".//div[1]/div[2]/table[2]/tbody[{n}]")

    @staticmethod
    def select_hora(n: int):
        return (
            By.XPATH,
            f".//div[1]/div[2]/table[2]/tbody[{n}]/tr[2]/td[3]/span/select[1]"
        )

    @staticmethod
    def select_minuto(n: int):
        return (
            By.XPATH,
            f".//div[1]/div[2]/table[2]/tbody[{n}]/tr[2]/td[3]/span/select[2]"
        )

    @staticmethod
    def input_codigo(n: int):
        return (
            By.XPATH,
            f".//div[1]/div[2]/table[2]/tbody[{n}]/tr[2]/td[4]/div/input"
        )

    @staticmethod
    def textarea_observacion(n: int):
        return (
            By.XPATH,
            f".//div[1]/div[2]/table[2]/tbody[{n}]/tr[4]/td[2]/textarea"
        )

