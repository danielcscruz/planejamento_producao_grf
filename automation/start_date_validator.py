# /utils/start_date_validator.py
from datetime import datetime

def validar_data_input(data_str):
    try:
        # Tenta converter a string para um objeto datetime
        data = datetime.strptime(data_str, "%d/%m/%Y")
    except ValueError:
        return None  # Retorna None se a data não for válida
    
    # Verifica se a data não está no passado
    if data < datetime.today():
        raise ValueError(f"❌ A data {data_str} não pode ser no passado!")
    
    return data
