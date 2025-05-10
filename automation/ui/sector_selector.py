from InquirerPy import inquirer
from automation.core.constants import SETOR_ORDEM

def selecionar_setor_inicio(df_priorizado):
    respostas = {}
    for _, row in df_priorizado.iterrows():
        chave = row["PEDIDO"]
        descricao = f"\n{row['PEDIDO']} - {row['CLIENTE']} - {row['PRODUTO']}"
        setor = inquirer.select(
            message=f"Escolha o setor de Inicio para:\n{descricao}",
            choices=SETOR_ORDEM,
            default=SETOR_ORDEM[0]
        ).execute()
        respostas[chave] = {
            "setor": setor,
            "cliente": row["CLIENTE"],
            "produto": row["PRODUTO"]
        }
    return respostas