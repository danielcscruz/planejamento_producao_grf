import pyfiglet
from tabulate import tabulate
from InquirerPy import inquirer
import pandas as pd

from automation import gerar_relatorio_arquivo, criar_novo_plano, definir_ordem_manual, definir_prioridade, escolher_acao, escolher_arquivo_excel, preencher_producao, processar_tabela, selecionar_tipos_de_corte, validar_data_input, validar_prazo, escolher_arquivo_exportar
from automation.core.constants import DEFAULT_CONFIG_PATH

import os
import sys  # Para fechar o script com segurança


def main():
    # Limpar a tela do console
    os.system('cls' if os.name == 'nt' else 'clear')
    f = pyfiglet.Figlet(font="basic", width=80)
    print('\n')
    print(f.renderText('plano   de producao'))

    while True:
        # Pergunta o que o usuário deseja fazer
        # os.system('cls' if os.name == 'nt' else 'clear')

        acao = escolher_acao()

        if acao == "🚪: Sair":
            print("\n👋 Saindo do programa. Até logo!")
            sys.exit()  # Fecha o script com segurança
        
        if acao == "📊: Exportar Relatórios":
            arquivo_exportar = escolher_arquivo_exportar()

            if arquivo_exportar == None:
                continue
            
            gerar_relatorio_arquivo(arquivo_exportar)

        if acao == "⚙️ : Configurações ":
            try:
                # Lê o arquivo CSV usando pandas
                config_df = pd.read_csv(DEFAULT_CONFIG_PATH, encoding='utf-16')
                colunas = ["PARAMETRO", "VALOR", "UNIDADE", "DESCRICAO"]
                print(tabulate(config_df, headers=colunas, tablefmt="fancy_grid"))
                # Cria as opções para o menu
                opcoes = [
                    f"{row['PARAMETRO']} | {row['VALOR']} | {row['UNIDADE']} | {row['DESCRICAO']}"
                    for _, row in config_df.iterrows()
                ]
                opcoes.append("🔙 Voltar para o menu inicial")

                # Exibe o menu para o usuário
                escolha = inquirer.select(
                    message="Escolha qual parametro você deseja alterar:",
                    choices=opcoes,
                ).execute()
                
                if escolha == "🔙 Voltar para o menu inicial":
                    print("\n🔙 Retornando ao menu inicial...\n")
                    continue  # Volta para o início do loop principal

                # Extrai o parâmetro escolhido
                parametro_escolhido = escolha.split(" | ")[0]

                # Solicita um novo valor para o parâmetro
                while True:
                    try:
                        novo_valor = int(input(f"Digite o novo valor inteiro para {parametro_escolhido}: "))
                        break  # Sai do loop se o valor for válido
                    except ValueError:
                        print("❌ Entrada inválida. Por favor, insira um número inteiro.")

                # Atualiza o valor no DataFrame
                config_df.loc[config_df['PARAMETRO'] == parametro_escolhido, 'VALOR'] = novo_valor

                # Salva as alterações de volta no arquivo CSV
                config_df.to_csv(DEFAULT_CONFIG_PATH, index=False, encoding='utf-16')
                print(f"✅ O valor de {parametro_escolhido} foi atualizado para {novo_valor} com sucesso!")

            except Exception as e:
                print(f"Erro inesperado ao carregar o arquivo de configuração: {e}")

        if acao == "📥: Carregar Pedidos":
            # Gerar um novo plano
            arquivo = escolher_arquivo_excel()

            if arquivo == None:
                continue

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