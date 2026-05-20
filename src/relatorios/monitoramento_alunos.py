import re
import fitz
import pandas as pd
from datetime import datetime


COLUNAS_MONITORAMENTO = [
    "Matrícula",
    "Aluno",
    "Período",
    "Turma",
    "Entrou no AVA",
    "Total de acessos",
    "Último acesso",
    "Dias sem acesso",
    "Risco",
    "Sugestão de ação",
]


def classificar_risco_monitoramento(dias_sem_acesso):
    if dias_sem_acesso >= 7:
        return "🔴 ALTO RISCO"
    elif dias_sem_acesso >= 5:
        return "🟡 MÉDIO RISCO"
    return "🟢 SEM RISCO"


def sugerir_acao(entrou, risco):
    if entrou == "Não":
        return "Contato imediato para verificar dificuldade de acesso, login, vínculo no AVA ou ausência de participação."
    if "ALTO" in risco:
        return "Contato ativo da coordenação/tutoria e acompanhamento pedagógico prioritário."
    if "MÉDIO" in risco:
        return "Acompanhamento preventivo e orientação para retomada da participação no AVA."
    return "Manter acompanhamento regular."


def limpar_texto(texto):
    texto = str(texto).strip()
    texto = re.sub(r"\.{2,}", "", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def extrair_periodo_turma(texto):
    texto = str(texto).upper()

    periodo = ""
    turma = ""

    match_periodo = re.search(r"(\d+)P", texto)
    if match_periodo:
        periodo = f"{match_periodo.group(1)}º período"

    match_turma = re.search(r"MED\s+([A-Z])", texto)
    if match_turma:
        turma = match_turma.group(1)

    return periodo, turma


def ler_lista_alunos_pdf(caminho_pdf):
    alunos = []
    turma_atual = ""
    periodo_atual = ""
    letra_turma_atual = ""

    documento = fitz.open(caminho_pdf)

    for pagina in documento:
        texto_pagina = pagina.get_text("text")

        for linha in texto_pagina.splitlines():
            if "Turma:" in linha:
                turma_atual = linha
                periodo_atual, letra_turma_atual = extrair_periodo_turma(linha)

        words = pagina.get_text("words")

        linhas_por_y = {}

        for w in words:
            x0, y0, x1, y1, palavra = w[:5]
            y_chave = round(y0, 1)

            if y_chave not in linhas_por_y:
                linhas_por_y[y_chave] = []

            linhas_por_y[y_chave].append((x0, palavra))

        for y in sorted(linhas_por_y.keys()):
            itens = sorted(linhas_por_y[y], key=lambda x: x[0])
            palavras = [p for _, p in itens]

            if not palavras:
                continue

            primeira = palavras[0]

            if re.fullmatch(r"\d{8,10}", primeira):
                matricula = primeira

                nome_partes = []

                for x, palavra in itens[1:]:
                    if "." in palavra:
                        break

                    if palavra.lower() in [
                        "subtotal:",
                        "total:",
                        "página:",
                        "faculdade",
                    ]:
                        break

                    nome_partes.append(palavra)

                nome = limpar_texto(" ".join(nome_partes))

                if len(nome) >= 5:
                    alunos.append(
                        {
                            "Matrícula": matricula,
                            "Aluno": nome,
                            "Período": periodo_atual,
                            "Turma": letra_turma_atual,
                        }
                    )

    documento.close()

    df_alunos = pd.DataFrame(
        alunos,
        columns=["Matrícula", "Aluno", "Período", "Turma"]
    )

    df_alunos = df_alunos.drop_duplicates(
        subset=["Matrícula", "Aluno"]
    )

    print("\nTotal de alunos encontrados no PDF:")
    print(len(df_alunos))

    print("\nPrimeiros alunos extraídos:")
    print(df_alunos.head(10))

    df_alunos.to_excel(
        "outputs/excel/debug_alunos_extraidos_pdf.xlsx",
        index=False
    )

    return df_alunos


def gerar_monitoramento_alunos(df_logs, dias_analise, caminho_pdf_alunos):
    df_alunos = ler_lista_alunos_pdf(caminho_pdf_alunos)

    resultados = []
    agora = datetime.now()

    if "Nome" not in df_logs.columns:
        df_logs["Nome"] = ""

    if "Hora" not in df_logs.columns:
        df_logs["Hora"] = pd.NaT

    df_logs["Nome"] = df_logs["Nome"].astype(str)

    for _, aluno_linha in df_alunos.iterrows():
        matricula = aluno_linha["Matrícula"]
        aluno = aluno_linha["Aluno"]
        periodo = aluno_linha["Período"]
        turma = aluno_linha["Turma"]

        logs_aluno = df_logs[
            df_logs["Nome"].str.upper().str.contains(
                str(aluno).upper(),
                na=False,
                regex=False
            )
        ]

        if len(logs_aluno) > 0:
            ultimo_acesso = logs_aluno["Hora"].max()
            total_acessos = len(logs_aluno)
            dias_sem_acesso = (agora - ultimo_acesso).days
            entrou = "Sim"
        else:
            ultimo_acesso = "Sem registro"
            total_acessos = 0
            dias_sem_acesso = dias_analise
            entrou = "Não"

        risco = classificar_risco_monitoramento(dias_sem_acesso)

        resultados.append(
            {
                "Matrícula": matricula,
                "Aluno": aluno,
                "Período": periodo,
                "Turma": turma,
                "Entrou no AVA": entrou,
                "Total de acessos": total_acessos,
                "Último acesso": ultimo_acesso,
                "Dias sem acesso": dias_sem_acesso,
                "Risco": risco,
                "Sugestão de ação": sugerir_acao(entrou, risco),
            }
        )

    df_monitoramento = pd.DataFrame(
        resultados,
        columns=COLUNAS_MONITORAMENTO
    )

    print("\nTotal de alunos no monitoramento:")
    print(len(df_monitoramento))

    print("\nPrimeiros alunos no monitoramento:")
    print(df_monitoramento.head(10))

    return df_monitoramento