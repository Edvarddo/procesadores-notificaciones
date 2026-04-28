from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from core.carabineros_formulario.locators.login_locators import LoginLocators


class LoginPage:
    def __init__(self, driver, timeout: int):
        self.driver = driver
        self.wait = WebDriverWait(driver, timeout)

    def abrir(self, url: str) -> None:
        self.driver.get(url)

    def login(self, usuario: str, password: str) -> None:
        campo_usuario = self.wait.until(
            EC.element_to_be_clickable(LoginLocators.USUARIO)
        )
        campo_usuario.clear()
        campo_usuario.send_keys(usuario)

        campo_password = self.wait.until(
            EC.element_to_be_clickable(LoginLocators.PASSWORD)
        )
        campo_password.clear()
        campo_password.send_keys(password)
        campo_password.send_keys(Keys.RETURN)

        self.wait.until(EC.url_contains("/home/"))
