from docx import Document

from datetime import datetime


def gerar_relatorio_word(df, dias_analise):

    data_hoje = datetime.now().strftime("%d-%m-%Y")

    nome_relatorio = (
        f"outputs/word/relatorio_monitoramento_{data_hoje}.docx"
    )

    doc = Document()

    doc.add_heading(
        f"ATHENA AVA INTELLIGENCE - {data_hoje}",
        level=0
    )

    doc.add_heading(
        "Relatório Institucional de Monitoramento Acadêmico",
        level=1
    )

    doc.add_paragraph(
        f"Período analisado: últimos {dias_analise} dias."
    )

    doc.add_heading(
        "Introdução Metodológica",
        level=1
    )

    doc.add_paragraph(
        f"""
O presente relatório foi elaborado pelo sistema Athena AVA Intelligence com o objetivo de monitorar indicadores de engajamento acadêmico no Ambiente Virtual de Aprendizagem institucional.

A análise considerou registros de acesso estudantil referentes aos últimos {dias_analise} dias, permitindo identificar padrões de participação, frequência digital e potenciais situações de risco acadêmico.

A classificação de risco adotada foi baseada no intervalo de dias sem acesso ao AVA:

• 🟢 Baixo risco: até 4 dias sem acesso;
• 🟡 Médio risco: entre 5 e 6 dias sem acesso;
• 🔴 Alto risco: 7 dias ou mais sem atividade.
"""
    )

    df["Risco"] = df["Risco"].astype(str)

    total = len(df)

    baixo = len(
        df[df["Risco"].str.contains("BAIXO", na=False)]
    )

    medio = len(
        df[df["Risco"].str.contains("MÉDIO", na=False)]
    )

    alto = len(
        df[df["Risco"].str.contains("ALTO", na=False)]
    )

    doc.add_heading(
        "Resumo Executivo",
        level=1
    )

    doc.add_paragraph(
        f"""
Total de registros analisados: {total}

🟢 Baixo risco: {baixo}
🟡 Médio risco: {medio}
🔴 Alto risco: {alto}
"""
    )

    doc.add_heading(
        "Tabela de Monitoramento",
        level=1
    )

    tabela = doc.add_table(
        rows=1,
        cols=4
    )

    tabela.style = "Table Grid"

    cabecalho = tabela.rows[0].cells

    cabecalho[0].text = "Aluno"
    cabecalho[1].text = "Último acesso"
    cabecalho[2].text = "Dias sem acesso"
    cabecalho[3].text = "Risco"

    for _, linha in df.iterrows():

        row = tabela.add_row().cells

        row[0].text = str(linha.get("Nome", ""))
        row[1].text = str(linha.get("Hora", ""))
        row[2].text = str(linha.get("Dias_sem_acesso", ""))
        row[3].text = str(linha.get("Risco", ""))

    doc.add_heading(
        "Conclusão Institucional",
        level=1
    )

    doc.add_paragraph(
        """
O monitoramento automatizado do Ambiente Virtual de Aprendizagem permitiu identificar estudantes em diferentes níveis de risco acadêmico, possibilitando ações institucionais de acompanhamento, permanência estudantil e intervenção pedagógica precoce.
"""
    )

    doc.save(nome_relatorio)

    print("\nRelatório Word criado com sucesso.")
    print(nome_relatorio)