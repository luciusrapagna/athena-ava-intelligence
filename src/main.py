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

    elif opcao == "2":
        return 15

    elif opcao == "3":
        return 30

    elif opcao == "4":

        dias = int(
            input(
                "\nDigite o número de dias desejado: "
            )
        )

        return dias

    else:

        print("\nOpção inválida.")
        return menu_periodo()


def main():

    dias_analise = menu_periodo()

    driver = fazer_login()

    abrir_relatorio_acessos(
        driver,
        dias_analise
    )

    input("\nPressione ENTER para finalizar...")


if __name__ == "__main__":
    main()