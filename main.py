import pyfiglet
from automation.file_selector import escolher_arquivo_excel
from automation.table_renderer import processar_tabela, validar_data_input
from automation.cut_selector import selecionar_tipos_de_corte
from automation.priority_handler import definir_prioridade
from automation.create_plan import criar_novo_plano
from automation.add_row import adicionar_nova_linha
from tabulate import tabulate
from automation.action_selector import escolher_acao  # Importando a função de escolha de ação
from utils.start_date_validator import validar_data_input  # Importando a validação de data


def main():
    # Pergunta o que o usuário deseja fazer
    f = pyfiglet.Figlet(font="basic", width=80)
    print('\n')
    print(f.renderText('plano   de producao'))
    acao = escolher_acao()

    if acao == "Adicionar um Pedido ao Plano de Produção":
        # Lógica para atualizar um plano de produção
        arquivo = escolher_arquivo_excel()
        df_formatado, _ = processar_tabela(arquivo)
        adicionar_nova_linha(arquivo)
        df_formatado, _ = processar_tabela(arquivo)

    elif acao == "Gerar planos dos Setores":
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

        df_priorizado = definir_prioridade(df_formatado)

        print("\n📋 Tabela final com prioridade:\n")
        print(tabulate(df_priorizado, headers='keys', tablefmt='grid', showindex=False))

        while True:
            try:
                inicio_data_str = input(f"Digite a data de inicio do Plano de Produção no formato DD/MM/AAAA: ")
                inicio_data = validar_data_input(inicio_data_str)
                print(f"✔️ Data de inicio do Plano de Produção: {inicio_data.strftime('%d/%m/%Y')}")
                break
            except ValueError as e:
                print(e)  # Exibe o erro se a data for inválida ou no passado
        
        criar_novo_plano(df_formatado, inicio_data)




if __name__ == "__main__":
    main()
