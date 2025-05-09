from InquirerPy import inquirer

def escolher_acao():
    resposta = inquirer.select(
        message="O que você deseja fazer?",
        choices=["📥: Carregar Pedidos", "➕: Adicionar um Pedido", "📊: Exportar Relatórios", "⚙️ : Configurações "  ],
        default="📥: Carregar Pedidos"
    ).execute()
    
    return resposta
