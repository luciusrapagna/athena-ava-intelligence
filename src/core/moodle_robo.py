import time
from datetime import datetime
from pathlib import Path

from selenium.webdriver.common.by import By

from src.utils.navegador import iniciar_navegador


BASE_DIR = Path(__file__).resolve().parents[2]

OUTPUTS_DIR = BASE_DIR / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)


MOODLE_URL = "https://www.unilagos-presencial.eadmax.net/my/"
MOODLE_USER = "luciano"
MOODLE_PASSWORD = "@Luciano1771@"


def executar_robo_moodle(periodo: str = "5º e 6º períodos"):
    log = []
    inicio = datetime.now()

    log.append("ATHENA SCIENTIFIC | AVA INTELLIGENCE")
    log.append("=" * 60)
    log.append(f"Início: {inicio.strftime('%Y-%m-%d %H:%M:%S')}")
    log.append(f"Período selecionado: {periodo}")

    driver = None
    sucesso = False
    screenshot_path = None

    try:
        log.append("\n[1] Abrindo navegador headless...")
        driver = iniciar_navegador()
        log.append("Navegador iniciado.")

        log.append(f"\n[2] Acessando Moodle: {MOODLE_URL}")
        driver.get(MOODLE_URL)
        time.sleep(5)

        log.append("\n[3] Localizando campos de login...")

        username_input = driver.find_element(By.ID, "username")
        password_input = driver.find_element(By.ID, "password")

        log.append("Campos localizados.")

        log.append("\n[4] Preenchendo credenciais...")

        username_input.clear()
        username_input.send_keys(MOODLE_USER)

        password_input.clear()
        password_input.send_keys(MOODLE_PASSWORD)

        time.sleep(2)

        log.append("\n[5] Enviando login via clique JavaScript...")

        login_button = driver.find_element(By.ID, "loginbtn")

        driver.execute_script(
            "arguments[0].click();",
            login_button
        )

        time.sleep(10)

        titulo = driver.title
        url_atual = driver.current_url

        log.append("\n[6] Verificando sessão")
        log.append(f"Título da página: {titulo}")
        log.append(f"URL atual: {url_atual}")

        screenshot_path = OUTPUTS_DIR / "moodle_login_teste.png"
        driver.save_screenshot(str(screenshot_path))

        log.append(f"Screenshot salvo em: {screenshot_path}")

        if "login" in url_atual.lower():
            log.append("\nATENÇÃO: O Moodle permaneceu na tela de login.")
            sucesso = False
        else:
            log.append("\nLogin realizado com sucesso.")
            sucesso = True

    except Exception as erro:
        log.append("\n[ERRO]")
        log.append(str(erro))
        sucesso = False

    finally:
        if driver:
            driver.quit()
            log.append("\nNavegador fechado.")

    fim = datetime.now()

    log.append(f"\nFim: {fim.strftime('%Y-%m-%d %H:%M:%S')}")
    log.append(f"Duração: {fim - inicio}")
    log.append("=" * 60)

    log_path = OUTPUTS_DIR / "fase1_robo_moodle_log.txt"
    log_path.write_text("\n".join(log), encoding="utf-8")

    return {
        "sucesso": sucesso,
        "log": "\n".join(log),
        "log_path": str(log_path),
        "screenshot": str(screenshot_path) if screenshot_path else None
    }
