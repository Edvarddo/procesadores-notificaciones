from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from core.carabineros_formulario.config.settings import HEADLESS


def crear_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")

    if HEADLESS:
        options.add_argument("--headless=new")

    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
