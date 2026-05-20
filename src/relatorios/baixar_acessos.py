import time
import pandas as pd

from datetime import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from relatorios.gerar_relatorio_word import gerar_relatorio_word

from relatorios.monitoramento_alunos import (
    gerar_monitoramento_alunos
)


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

            texto = texto.replace(
                mes_pt,
                mes_num
            )

        return pd.to_datetime(
            texto,
            dayfirst=True,
            errors="coerce"
        )

    except Exception:

        return pd.NaT


def abrir_relatorio_acessos(
    driver,
    dias_analise,
    caminho_pdf_alunos
):

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

    linhas = tbody.find_elements(
        By.TAG_NAME,
        "tr"
    )

    dados = []

    for linha in linhas:

        colunas = linha.find_elements(
            By.TAG_NAME,
            "td"
        )

        valores = [
            coluna.text.strip()
            for coluna in colunas
        ]

        if valores:

            dados.append(valores)

    print("\nTOTAL DE LINHAS CAPTURADAS:")

    print(len(dados))

    if len(dados) == 0:

        print("\nNenhum dado encontrado.")

        return driver

    df_logs = pd.DataFrame(dados)

    quantidade_colunas = df_logs.shape[1]

    df_logs.columns = [
        f"Coluna_{i}"
        for i in range(quantidade_colunas)
    ]

    df_logs["Hora"] = df_logs[
        "Coluna_0"
    ].apply(converter_data_moodle)

    df_logs["Nome"] = df_logs[
        "Coluna_1"
    ]

    df_logs = df_logs.dropna(
        subset=["Hora"]
    )

    agora = datetime.now()

    df_logs["Dias_sem_acesso"] = (
        agora - df_logs["Hora"]
    ).dt.days

    df_logs["Risco"] = df_logs[
        "Dias_sem_acesso"
    ].apply(classificar_risco)

    df_logs = df_logs[
        df_logs["Dias_sem_acesso"] <= dias_analise
    ]

    df_monitoramento = gerar_monitoramento_alunos(
        df_logs,
        dias_analise,
        caminho_pdf_alunos
    )

    df_alto_risco = df_monitoramento[
        df_monitoramento["Risco"].str.contains(
            "ALTO",
            na=False
        )
    ]

    df_medio_risco = df_monitoramento[
        df_monitoramento["Risco"].str.contains(
            "MÉDIO",
            na=False
        )
    ]

    df_sem_acesso = df_monitoramento[
        df_monitoramento[
            "Entrou no AVA"
        ] == "Não"
    ]

    data_hoje = datetime.now().strftime(
        "%d-%m-%Y"
    )

    caminho_excel = (
        f"outputs/excel/monitoramento_alunos_{data_hoje}.xlsx"
    )

    caminho_csv = (
        f"outputs/csv/logs_acessos_{data_hoje}.csv"
    )

    with pd.ExcelWriter(
        caminho_excel,
        engine="openpyxl"
    ) as writer:

        df_logs.to_excel(
            writer,
            sheet_name="Logs_Brutos",
            index=False
        )

        df_monitoramento.to_excel(
            writer,
            sheet_name="Monitoramento_Alunos",
            index=False
        )

        df_alto_risco.to_excel(
            writer,
            sheet_name="Alunos_Alto_Risco",
            index=False
        )

        df_medio_risco.to_excel(
            writer,
            sheet_name="Alunos_Medio_Risco",
            index=False
        )

        df_sem_acesso.to_excel(
            writer,
            sheet_name="Alunos_Sem_Acesso",
            index=False
        )

    df_logs.to_csv(
        caminho_csv,
        index=False,
        encoding="utf-8-sig"
    )

    gerar_relatorio_word(
        df_monitoramento,
        dias_analise
    )

    print("\nRelatórios criados com sucesso.")

    print(caminho_excel)

    print(caminho_csv)

    return driver