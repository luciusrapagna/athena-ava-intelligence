import time
import pandas as pd

from datetime import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from relatorios.gerar_relatorio_word import gerar_relatorio_word


URL_RELATORIO_ACESSOS = (
    "https://www.unilagos-presencial.eadmax.net/report/log/index.php?id=18"
)


def classificar_risco(dias):
    if dias >= 7:
        return "🔴 ALTO RISCO"
    elif dias >= 5:
        return "🟡 MÉDIO RISCO"
    else:
        return "🟢 BAIXO RISCO"


def converter_data_moodle(texto):
    meses = {
        "janeiro": "01",
        "fevereiro": "02",
        "março": "03",
        "abril": "04",
        "maio": "05",
        "junho": "06",
        "julho": "07",
        "agosto": "08",
        "setembro": "09",
        "outubro": "10",
        "novembro": "11",
        "dezembro": "12",
    }

    try:
        texto = str(texto).lower().strip()

        for mes_pt, mes_num in meses.items():
            texto = texto.replace(mes_pt, mes_num)

        return pd.to_datetime(
            texto,
            dayfirst=True,
            errors="coerce"
        )

    except Exception:
        return pd.NaT


def abrir_relatorio_acessos(driver, dias_analise):
    print("\nAbrindo relatório de acessos...")

    driver.get(URL_RELATORIO_ACESSOS)

    wait = WebDriverWait(driver, 30)

    botao_logs = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//input[@value='Obter estes logs']"
            )
        )
    )

    botao_logs.click()

    time.sleep(10)

    print("\nLogs carregados.")

    tbody = wait.until(
        EC.presence_of_element_located(
            (By.TAG_NAME, "tbody")
        )
    )

    linhas = tbody.find_elements(By.TAG_NAME, "tr")

    dados = []

    for linha in linhas:
        colunas = linha.find_elements(By.TAG_NAME, "td")

        valores = [
            coluna.text.strip()
            for coluna in colunas
        ]

        if valores:
            dados.append(valores)

    print("\nTOTAL DE LINHAS:")
    print(len(dados))

    if len(dados) == 0:
        print("\nNenhum dado encontrado.")
        return driver

    df = pd.DataFrame(dados)

    print("\nPRIMEIRA LINHA:")
    print(df.iloc[0])

    quantidade_colunas = df.shape[1]

    print("\nTOTAL DE COLUNAS:")
    print(quantidade_colunas)

    df.columns = [
        f"Coluna_{i}"
        for i in range(quantidade_colunas)
    ]

    coluna_data = "Coluna_0"
    coluna_nome = "Coluna_1"

    df["Hora"] = df[coluna_data].apply(
        converter_data_moodle
    )

    df = df.dropna(
        subset=["Hora"]
    )

    agora = datetime.now()

    df["Dias_sem_acesso"] = (
        agora - df["Hora"]
    ).dt.days

    df["Risco"] = df["Dias_sem_acesso"].apply(
        classificar_risco
    )

    df["Nome"] = df[coluna_nome]

    data_hoje = datetime.now().strftime("%d-%m-%Y")

    caminho_excel = (
        f"outputs/excel/relatorio_risco_{data_hoje}.xlsx"
    )

    caminho_csv = (
        f"outputs/csv/relatorio_risco_{data_hoje}.csv"
    )

    df.to_excel(
        caminho_excel,
        index=False
    )

    df.to_csv(
        caminho_csv,
        index=False,
        encoding="utf-8-sig"
    )

    gerar_relatorio_word(
        df,
        dias_analise
    )

    print("\nRelatórios criados com sucesso.")
    print(caminho_excel)
    print(caminho_csv)

    return driver