import os

from login.moodle_login import fazer_login
from relatorios.baixar_acessos import abrir_relatorio_acessos


def menu_periodo():
    print("\n===================================")
    print("ATHENA AVA INTELLIGENCE")
    print("===================================\n")

    print("Selecione o período da análise:\n")
    print("1 - Últimos 7 dias")
    print("2 - Últimos 15 dias")
    print("3 - Últimos 30 dias")
    print("4 - Personalizado\n")

    opcao = input("Digite a opção desejada: ")

    if opcao == "1":
        return 7
    if opcao == "2":
        return 15
    if opcao == "3":
        return 30
    if opcao == "4":
        return int(input("\nDigite o número de dias desejado: "))

    print("\nOpção inválida.")
    return menu_periodo()


def selecionar_pdf():
    print("\nInforme o caminho completo do PDF da lista de alunos.\n")

    caminho_pdf = input("Caminho do PDF: ").strip().strip('"').strip("'")

    if not caminho_pdf:
        raise FileNotFoundError("Nenhum PDF foi informado.")

    if not os.path.exists(caminho_pdf):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_pdf}")

    print("\nPDF selecionado:")
    print(caminho_pdf)

    return caminho_pdf


def main():
    dias_analise = menu_periodo()
    caminho_pdf_alunos = selecionar_pdf()

    driver = fazer_login()

    abrir_relatorio_acessos(
        driver,
        dias_analise,
        caminho_pdf_alunos
    )

    input("\nPressione ENTER para finalizar...")


if __name__ == "__main__":
    main()
