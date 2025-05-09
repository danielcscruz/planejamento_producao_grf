"""
Funções de utilidade para manipulação de arquivos.
"""

from datetime import datetime
from pathlib import Path

def salvar_nova_versao(caminho_original, workbook):
    """
    Salva a planilha em uma nova versão na pasta exp/
    
    Args:
        caminho_original (str): Caminho do arquivo original.
        workbook: Objeto workbook do openpyxl.
        
    Returns:
        str or None: Caminho da nova versão salva ou None em caso de erro.
    """
    try:
        # Garantir que a pasta exp/ existe
        pasta_exp = Path("exp")
        pasta_exp.mkdir(exist_ok=True)
        
        # Obter o nome do arquivo original
        nome_arquivo = Path(caminho_original).name
        base_nome = Path(caminho_original).stem
        extensao = Path(caminho_original).suffix
        
        # Gerar um nome de arquivo com timestamp
        timestamp = datetime.now().strftime("%Y_%m_%d__%H_%M_%S")
        novo_nome = f"{base_nome}_{timestamp}{extensao}"
        novo_caminho = pasta_exp / novo_nome
        
        # Salvar a planilha no novo caminho
        workbook.save(str(novo_caminho))
        return str(novo_caminho)
    except Exception as e:
        return None