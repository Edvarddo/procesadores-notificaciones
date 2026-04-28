from selenium.webdriver.common.by import By


class CertificacionesLocators:
    MENU = (By.CSS_SELECTOR, "button.btn-menu")

    # Estos pueden necesitar ajuste real según inspección fina
    PROCESOS = (By.XPATH, "//mat-expansion-panel")
    CERTIFICACION = (By.XPATH, "//*[@id='cdk-accordion-child-0']")
    CERTIFICACIONES = (By.XPATH, "//*[@id='cdk-accordion-child-1']//button")

    FECHA = (By.XPATH, "//*[@id='mat-input-7']")
    RIT = (By.XPATH, "//*[@id='mat-input-2']")
    ANIO = (By.XPATH, "//*[@id='mat-input-3']")
    FECHA_CERTIFICACION = (By.XPATH, "//*[@id='mat-input-5']")
    ID_NOTIFICACION_CAMPO = (
    By.XPATH,
        "//*[@id='mat-input-4']"
        " | //input[@placeholder='ID de la Notificación']"
        " | //mat-label[contains(., 'ID de la Notificación')]/ancestor::mat-form-field//input"
    )
    BUSCAR = (
        By.XPATH,
        "//app-certificaciones//button[contains(., 'Buscar') or contains(., 'BUSCAR')]"
    )

    FILAS_RESULTADO = (
        By.XPATH,
        "//app-certificaciones//div[4]//table/tbody/tr"
    )

    CELDA_ID_REL = (By.XPATH, "./td[1]")
    BTN_CARATULA_REL = (By.XPATH, "./td[2]//mat-icon")

    ALERTA_MODAL = (By.XPATH, "//app-alerta-modal")
    BOTON_CERRAR_ALERTA = (By.XPATH, "//app-alerta-modal/div[2]/button")
    
    # Locators para notificaciones reservadas
    ALERTA_MODAL_DIALOG = (By.XPATH, "//mat-dialog-container[contains(@id, 'mat-dialog-')]")
    BOTON_GUARDAR_RESERVA = (By.XPATH, "//mat-dialog-container[contains(@id, 'mat-dialog-')]//app-alerta-modal/div[2]/button[2]")
    CELDA_RESERVADO_REL = (By.XPATH, "./td[1]")
    
    @staticmethod
    def BTN_LIMPIAR_REL():
        return (By.XPATH, "./td[40]/div/mat-icon[2]")
