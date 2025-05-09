import pandas as pd
from datetime import datetime

def validar_prazo(df_produzido: pd.DataFrame) -> pd.DataFrame:
    """
    Valida o prazo de entrega comparando as colunas 'ULTIMO DIA' e 'ENTREGA'.
    Adiciona uma nova coluna 'PRAZO' com os seguintes valores:
    - '✅' se 'ULTIMO DIA' < 'ENTREGA'
    - '⚠️' se 'ULTIMO DIA' == 'ENTREGA'
    - '❌' se 'ULTIMO DIA' > 'ENTREGA'

    Args:
        df_produzido (pd.DataFrame): DataFrame com as colunas 'ULTIMO DIA' e 'ENTREGA'.

    Returns:
        pd.DataFrame: DataFrame atualizado com a nova coluna 'PRAZO'.
    """
    def calcular_prazo(ultimo_dia, entrega):
        try:
            # Converte as datas para objetos datetime
            ultimo_dia_dt = datetime.strptime(ultimo_dia, "%d/%m/%Y")
            entrega_dt = datetime.strptime(entrega, "%d/%m/%Y")
            
            # Compara as datas e retorna o símbolo correspondente
            if ultimo_dia_dt < entrega_dt:
                return "✅"
            elif ultimo_dia_dt == entrega_dt:
                return "⚠️"
            else:
                return "❌"
        except Exception as e:
            # Retorna vazio em caso de erro na conversão
            return "❓"

    # Aplica a função de cálculo de prazo em cada linha do DataFrame
    df_produzido["PRAZO"] = df_produzido.apply(
        lambda row: calcular_prazo(row["ULTIMO DIA"], row["ENTREGA"]), axis=1
    )

    return df_produzido