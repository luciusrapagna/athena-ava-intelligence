import pandas as pd


def classificar_risco(dias_sem_acesso):
    if dias_sem_acesso == "" or pd.isna(dias_sem_acesso):
        return "Risco elevado"

    dias_sem_acesso = int(dias_sem_acesso)

    if dias_sem_acesso <= 3:
        return "Sem risco"
    elif dias_sem_acesso <= 7:
        return "Risco médio"
    return "Risco elevado"


def gerar_parecer(nome, dias_sem_acesso, acessos):
    if dias_sem_acesso == "" or pd.isna(dias_sem_acesso):
        return f"{nome} não apresentou acesso identificado no período analisado. Recomenda-se contato ativo."

    dias_sem_acesso = int(dias_sem_acesso)

    if dias_sem_acesso <= 3:
        return f"{nome} apresenta acesso recente ao AVA, sem risco atual."
    elif dias_sem_acesso <= 7:
        return f"{nome} está há {dias_sem_acesso} dias sem acessar o AVA, indicando risco médio."
    return f"{nome} está há {dias_sem_acesso} dias sem acessar o AVA, indicando risco elevado."


def finalizar_ranking(df_base, periodo):
    df = df_base.copy()

    df["Período"] = periodo
    df["Turma"] = ""

    df["Ranking"] = df["Acessos"].rank(
        ascending=False,
        method="min"
    ).astype(int)

    df["Risco"] = df["Dias sem acesso"].apply(classificar_risco)

    df["Parecer pedagógico"] = df.apply(
        lambda x: gerar_parecer(
            x["Aluno"],
            x["Dias sem acesso"],
            x["Acessos"]
        ),
        axis=1
    )

    return df[
        [
            "Matrícula",
            "Aluno",
            "Período",
            "Turma",
            "Acessos",
            "Data do último acesso",
            "Dias sem acesso",
            "Ranking",
            "Risco",
            "Parecer pedagógico"
        ]
    ].sort_values("Ranking")
