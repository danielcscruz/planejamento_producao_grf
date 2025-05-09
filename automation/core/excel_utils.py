"""
Funções de utilidade para manipulação de planilhas Excel.
"""

from datetime import datetime
from openpyxl.worksheet.worksheet import Worksheet

def encontrar_coluna_por_data(ws: Worksheet, data):
    """
    Encontra a coluna na planilha que corresponde à data fornecida.
    Verifica tanto formato string quanto objeto datetime.
    
    Args:
        ws (Worksheet): Planilha do openpyxl.
        data (datetime): Data a ser localizada na planilha.
        
    Returns:
        int or None: Número da coluna se encontrada, None caso contrário.
    """
    data_str = data.strftime("%d/%m/%Y")
    
    for col in range(8, ws.max_column + 1):  # Começando da coluna H (8)
        cell_value = ws.cell(row=2, column=col).value
        
        # Verifica se o valor da célula é uma data
        if isinstance(cell_value, datetime):
            # Compara o valor da data
            if cell_value.date() == data.date():
                return col
        elif isinstance(cell_value, str):
            # Tenta converter para datetime se for string
            try:
                # Tenta primeiro o formato DD/MM/YYYY
                cell_date = datetime.strptime(cell_value, "%d/%m/%Y")
                if cell_date.date() == data.date():
                    return col
            except ValueError:
                try:
                    # Tenta o formato YYYY-MM-DD
                    cell_date = datetime.strptime(cell_value, "%Y-%m-%d")
                    if cell_date.date() == data.date():
                        return col
                except ValueError:
                    # Ignora se não conseguir converter
                    pass
    
    return None

def calcular_producao_planejada(ws: Worksheet, setor_nome: str, coluna: int, linha_limite: int):
    """
    Calcula a produção já planejada para um setor em uma data específica.
    
    Args:
        ws (Worksheet): Planilha do openpyxl.
        setor_nome (str): Nome do setor para calcular a produção.
        coluna (int): Coluna correspondente à data.
        linha_limite (int): Linha limite até onde calcular.
        
    Returns:
        int: Soma da produção planejada para o setor na data.
    """
    valor_planejado = 0
    for row in range(12, linha_limite):
        setor_cell_value = ws.cell(row=row, column=7).value  # Coluna G é a 7ª coluna
        if setor_cell_value == setor_nome:
            cell_value = ws.cell(row=row, column=coluna).value
            valor_planejado += int(cell_value or 0)
            
    return valor_planejado

def obter_limite_producao(ws: Worksheet, linha_limite: int):
    """
    Obtém o limite de produção diário para um setor.
    
    Args:
        ws (Worksheet): Planilha do openpyxl.
        linha_limite (int): Linha onde está o limite do setor.
        
    Returns:
        int: Valor limite de produção diária para o setor.
    """
    try:
        valor_limite_cell = ws.cell(row=linha_limite, column=5).value
        valor_limite_max = int(valor_limite_cell) if valor_limite_cell is not None else 0
        return valor_limite_max
    except (ValueError, TypeError):
        return 0