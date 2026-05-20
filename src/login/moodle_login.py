import os
import time

from dotenv import load_dotenv

from selenium.webdriver.common.by import By

from utils.navegador import iniciar_navegador


load_dotenv()


MOODLE_URL = os.getenv("MOODLE_URL")
MOODLE_USER = os.getenv("MOODLE_USER")
MOODLE_PASSWORD = os.getenv("MOODLE_PASSWORD")


def fazer_login():

    driver = iniciar_navegador()

    driver.get(MOODLE_URL)

    time.sleep(3)

    campo_usuario = driver.find_element(
        By.ID,
        "username"
    )

    campo_senha = driver.find_element(
        By.ID,
        "password"
    )

    campo_usuario.send_keys(MOODLE_USER)

    campo_senha.send_keys(MOODLE_PASSWORD)

    botao_login = driver.find_element(
        By.ID,
        "loginbtn"
    )

    botao_login.click()

    time.sleep(5)

    print("\nLogin realizado com sucesso.")

    return driver