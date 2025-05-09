from InquirerPy import inquirer

def definir_prioridade(df):
    escolha = inquirer.select(
        message="Como você quer definir a prioridade da produção?",
        choices=[
            "Não priorizar",
            "Priorizar por prazo de entrega",
            "Priorizar por quantidade de produção",
            "Definir manualmente a ordem de produção"
        ]
    ).execute()

    if escolha == "Não priorizar":
        # Retorna a tabela sem nenhuma modificação
        return df.reset_index(drop=True)

    elif escolha == "Priorizar por prazo de entrega":
        # Ordena por "ENTREGA" em ordem decrescente
        df_ordenado = df.sort_values(by="ENTREGA", ascending=True).reset_index(drop=True)
        return df_ordenado

    elif escolha == "Priorizar por quantidade de produção":
        # Ordena por "QUANTIDADE" em ordem decrescente
        df_ordenado = df.sort_values(by="QUANTIDADE", ascending=True).reset_index(drop=True)
        return df_ordenado

    elif escolha == "Definir manualmente a ordem de produção":
        # Chama a função para definir a ordem manualmente
        return definir_ordem_manual(df)


def definir_ordem_manual(df):
    print("\nPedidos disponíveis para ordenar:\n")
    for i, row in df.iterrows():
        print(f"{i+1}. Pedido {row['PEDIDO']} - {row['PRODUTO']} - Entrega: {row['ENTREGA']}")

    nova_ordem_str = inquirer.text(
        message="\nDigite os números na nova ordem desejada (ex: 3,1,2,5,4,6):\n"
    ).execute()

    try:
        nova_ordem = [int(i.strip()) - 1 for i in nova_ordem_str.split(",")]
        df_novo = df.iloc[nova_ordem].reset_index(drop=True)
        return df_novo
    except Exception as e:
        print(f"\nErro ao processar a nova ordem: {e}\n")
        print("\nMantendo a ordem original.\n")
        return df