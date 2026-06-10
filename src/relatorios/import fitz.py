import fitz
import pandas as pd
from datetime import datetime


CAMINHO_PDF_ALUNOS = "data/lista_alunos/lista_oficial_alunos.pdf"


def classificar_risco_monitoramento(dias_sem_acesso):
    if dias_sem_acesso >= 7:
        return "🔴 ALTO RISCO"
    elif dias_sem_acesso >= 5:
        return "🟡 MÉDIO RISCO"
    else:
        return "🟢 BAIXO RISCO"


def limpar_nome(nome):
    nome = str(nome).strip()
    nome = " ".join(nome.split())
    return nome


def linha_parece_nome(linha):
    linha = limpar_nome(linha)

    if len(linha) < 8:
        return False

    if linha.isdigit():
        return False

    termos_ignorar = [
        "nome",
        "matrícula",
        "curso",
        "turma",
        "período",
        "relatório",
        "data",
        "página",
        "cpf",
        "unilagos",
        "faculdade",
    ]

    linha_minuscula = linha.lower()

    for termo in termos_ignorar:
        if termo in linha_minuscula:
            return False

    quantidade_letras = sum(c.isalpha() for c in linha)

    if quantidade_letras < 6:
        return False

    return True


def ler_lista_alunos_pdf(caminho_pdf=CAMINHO_PDF_ALUNOS):
    nomes = []

    documento = fitz.open(caminho_pdf)

    for pagina in documento:
        texto = pagina.get_text()
        linhas = texto.split("\n")

        for linha in linhas:
            linha = limpar_nome(linha)

            if linha_parece_nome(linha):
                nomes.append(linha)

    documento.close()

    nomes_unicos = sorted(list(set(nomes)))

    print("\nTotal de alunos encontrados no PDF:")
    print(len(nomes_unicos))

    return nomes_unicos


def gerar_monitoramento_alunos(df_logs, dias_analise):
    alunos_pdf = ler_lista_alunos_pdf()

    resultados = []

    agora = datetime.now()

    df_logs["Nome"] = df_logs["Nome"].astype(str)

    for aluno in alunos_pdf:
        aluno_limpo = limpar_nome(aluno)

        logs_aluno = df_logs[
            df_logs["Nome"].str.upper().str.contains(
                aluno_limpo.upper(),
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
                "Aluno": aluno_limpo,
                "Entrou no AVA": entrou,
                "Total de acessos": total_acessos,
                "Último acesso": ultimo_acesso,
                "Dias sem acesso": dias_sem_acesso,
                "Risco": risco,
            }
        )

    df_monitoramento = pd.DataFrame(resultados)

    return df_monitoramento