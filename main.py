import pyfiglet
from automation.file_selector import escolher_arquivo_excel
from automation.table_renderer import processar_tabela
from automation.cut_selector import selecionar_tipos_de_corte
from automation.priority_handler import definir_prioridade
from automation.create_plan import criar_novo_plano
from automation.add_row import adicionar_nova_linha
from automation.report_validator import validar_prazo
from tabulate import tabulate
from automation.action_selector import escolher_acao
from InquirerPy import inquirer
import os
import sys  # Para fechar o script com segurança


def main():
    # Limpar a tela do console
    os.system('cls' if os.name == 'nt' else 'clear')

    while True:
        # Pergunta o que o usuário deseja fazer
        f = pyfiglet.Figlet(font="basic", width=80)
        print('\n')
        print(f.renderText('plano   de producao'))
        acao = escolher_acao()

        if acao == "🚪: Sair":
            print("\n👋 Saindo do programa. Até logo!")
            sys.exit()  # Fecha o script com segurança

        if acao == "📥: Carregar Pedidos":
            # Gerar um novo plano
            arquivo = escolher_arquivo_excel()
            df_formatado, _ = processar_tabela(arquivo)

            tipos_corte = selecionar_tipos_de_corte(df_formatado)

            # Atribuir tipo de corte por PEDIDO
            df_formatado["TIPO DE CORTE"] = df_formatado["PEDIDO"].map(
                lambda pedido: tipos_corte[pedido]["tipo"]
            )

            print("\n🪚 Tabela com tipo de corte definido:\n")
            print(tabulate(df_formatado, headers='keys', tablefmt='grid', showindex=False))

            while True:
                df_priorizado = definir_prioridade(df_formatado)

                print("\n📋 Tabela final com prioridade:\n")
                print(tabulate(df_priorizado, headers='keys', tablefmt='grid', showindex=False))

                # Confirmação do usuário usando InquirerPy
                confirmacao = inquirer.select(
                    message="\nDeseja prosseguir ou reorganizar?",
                    choices=["Prosseguir", "Reorganizar"],
                    default="Prosseguir"
                ).execute()

                if confirmacao == "Prosseguir":
                    break
                elif confirmacao == "Reorganizar":
                    print("\n🔄 Reorganizando a tabela...\n")

            df_produzido = criar_novo_plano(df_priorizado)
            print("\n🗓️  Novo Plano Criado\n")
            print("\n📊  Relatório:\n")
            df_validado = validar_prazo(df_produzido)
            print(tabulate(df_validado, headers='keys', tablefmt='grid', showindex=False))

            # Perguntar ao usuário se deseja voltar ou sair
            proxima_acao = inquirer.select(
                message="\nO que você deseja fazer agora?",
                choices=["Voltar para o menu inicial", "Sair"],
                default="Voltar para o menu inicial"
            ).execute()

            if proxima_acao == "Sair":
                print("\n👋 Saindo do programa. Até logo!")
                sys.exit()  # Fecha o script com segurança


if __name__ == "__main__":
    main()