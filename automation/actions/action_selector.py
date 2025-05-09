from InquirerPy import inquirer

def escolher_acao():
    resposta = inquirer.select(
        message="O que você deseja fazer?",
        choices=["📥: Carregar Pedidos", "📊: Exportar Relatórios", "⚙️ : Configurações ", "🚪: Sair"  ],
        default="📥: Carregar Pedidos"
    ).execute()
    
    return resposta
