from InquirerPy import inquirer

def escolher_acao():
    resposta = inquirer.select(
        message="O que você deseja fazer?",
        choices=["Adicionar um Pedido ao Plano de Produção", "Gerar planos dos Setores" ],
        default="Adicionar um Pedido ao Plano de Produção"
    ).execute()
    
    return resposta
