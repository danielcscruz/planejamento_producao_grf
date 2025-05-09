from InquirerPy import inquirer
import pandas as pd


def definir_prioridade(df):
    escolha = inquirer.select(
        message="\nComo você quer definir a prioridade da produção?\n",
        choices=[
            "Ordenar automaticamente por prazo de entrega (crescente)",
            "Definir manualmente a ordem de produção"
        ]
    ).execute()

    if escolha.startswith("Ordenar"):
        df_ordenado = df.sort_values(by="ENTREGA", ascending=True).reset_index(drop=True)
        return df_ordenado
    else:
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
