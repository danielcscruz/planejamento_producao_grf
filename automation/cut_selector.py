from InquirerPy import inquirer

def selecionar_tipos_de_corte(df_formatado):
    respostas = {}
    for _, row in df_formatado.iterrows():
        chave = row["PEDIDO"]
        descricao = f"{row['PEDIDO']} - {row['CLIENTE']} - {row['PRODUTO']}"
        tipo = inquirer.select(
            message=f"Escolha o tipo de corte para:\n{descricao}",
            choices=["Corte manual", "Corte a laser"],
            default="Corte manual"
        ).execute()
        respostas[chave] = {
            "tipo": tipo,
            "cliente": row["CLIENTE"],
            "produto": row["PRODUTO"]
        }
    return respostas
