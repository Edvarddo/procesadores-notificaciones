from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def esperar_clickable(driver, locator, timeout: int):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(locator)
    )


def esperar_visible(driver, locator, timeout: int):
    return WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located(locator)
    )


def esperar_presente(driver, locator, timeout: int):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located(locator)
    )


def esperar_todos_presentes(driver, locator, timeout: int):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located(locator)
    )
