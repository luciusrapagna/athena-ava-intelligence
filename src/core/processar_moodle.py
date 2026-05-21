import pandas as pd
from datetime import datetime


def ler_relatorio_moodle(caminho):
    caminho = str(caminho)

    if caminho.endswith(".csv"):
        df = pd.read_csv(caminho)
    else:
        df = pd.read_excel(caminho)

    df.columns = [str(c).strip() for c in df.columns]

    return df


def normalizar_nome(nome):
    return str(nome).strip().upper()


def cruzar_alunos_moodle(alunos, caminho_relatorio):
    df_moodle = ler_relatorio_moodle(caminho_relatorio)

    colunas = df_moodle.columns.tolist()

    coluna_nome = None
    coluna_data = None

    for c in colunas:
        c_lower = c.lower()

        if "nome" in c_lower or "usuário" in c_lower or "usuario" in c_lower:
            coluna_nome = c

        if "data" in c_lower or "hora" in c_lower or "acesso" in c_lower:
            coluna_data = c

    dados = []

    hoje = datetime.today()

    for aluno in alunos:
        matricula = aluno.get("matricula", "")
        nome = aluno.get("nome", "")

        acessos = 0
        ultimo_acesso = None
        dias_sem_acesso = 999

        if coluna_nome:
            filtro = df_moodle[
                df_moodle[coluna_nome]
                .astype(str)
                .str.upper()
                .str.contains(normalizar_nome(nome), na=False)
            ]

            acessos = len(filtro)

            if coluna_data and not filtro.empty:
                datas = pd.to_datetime(
                    filtro[coluna_data],
                    errors="coerce",
                    dayfirst=True
                ).dropna()

                if not datas.empty:
                    ultimo_acesso = datas.max()
                    dias_sem_acesso = (hoje - ultimo_acesso).days

        dados.append({
            "Matrícula": matricula,
            "Aluno": nome,
            "Acessos": acessos,
            "Data do último acesso": ultimo_acesso.strftime("%d/%m/%Y") if ultimo_acesso is not None else "",
            "Dias sem acesso": dias_sem_acesso if dias_sem_acesso != 999 else "",
        })

    return pd.DataFrame(dados)
