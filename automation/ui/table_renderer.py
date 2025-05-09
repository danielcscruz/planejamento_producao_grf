import pandas as pd
from datetime import datetime
from tabulate import tabulate
from openpyxl import load_workbook

def validar_data_input(data_str):
    try:
        return datetime.strptime(data_str, '%d/%m/%Y')
    except ValueError:
        return None

def processar_tabela(file_choice):
    print("Lendo o arquivo inicial...")
    arquivo_destino = f"{file_choice}.xlsx" if not str(file_choice).endswith('.xlsx') else file_choice
    print(f"\nAbrindo planilha: {arquivo_destino}")
    
    # Localiza linha do cabeçalho
    wb = load_workbook(arquivo_destino)
    ws = wb.active
    header_row = None
    for i, row in enumerate(ws.iter_rows(values_only=True), start=0):
        if row and "Pedido" in row:
            header_row = i
            break
    
    if header_row is None:
        raise ValueError("Coluna 'Pedido' não encontrada em nenhuma linha!")
    
    print(f"-> Cabeçalho encontrado na linha {header_row + 1}")
    
    df = pd.read_excel(arquivo_destino, header=header_row)
    df = df[df["Pedido"].notna()]
    df_unique = df.drop_duplicates(subset="Pedido", keep="first")
    
    colunas_desejadas = ["Pedido", "Entrega", "Cliente", "Produto", "QTD"]
    df_formatado = df_unique[colunas_desejadas].copy()
    df_formatado.columns = ['PEDIDO', 'ENTREGA', 'CLIENTE', 'PRODUTO', 'QUANTIDADE']
    
    # Converte PEDIDO para string
    df_formatado['PEDIDO'] = df_formatado['PEDIDO'].astype(str).str.strip()
    # Remove a parte decimal dos números (.0)
    df_formatado['PEDIDO'] = df_formatado['PEDIDO'].str.replace('.0$', '', regex=True)
    
    # Detecta datas inválidas
    datas_parseadas = pd.to_datetime(df_formatado['ENTREGA'].astype(str), format='%d/%m/%Y', errors='coerce')
    for idx, data in enumerate(df_formatado['ENTREGA']):
        if isinstance(data, (datetime, pd.Timestamp)):
            continue # Já é uma data válida
        
        # Ignora horas ou valores vazios
        if str(data).strip() in ['', 'nan', 'NaT']:
            print(f"\n⚠️ Data inválida detectada na linha {idx + header_row + 2} (valor: '{data}')")
            while True:
                nova_data_str = input(f"Digite uma nova data para '{data}' no formato DD/MM/AAAA: ")
                nova_data = validar_data_input(nova_data_str)
                if nova_data:
                    ws.cell(row=idx + header_row + 2, column=colunas_desejadas.index("Entrega") + 1).value = nova_data
                    print(f"✔️ Data corrigida para {nova_data_str}")
                    break
                else:
                    print("❌ Formato inválido. Tente novamente.")
    
    # Salva as correções no arquivo
    wb.save(arquivo_destino)
    wb.close()
    
    # Recarrega com datas corrigidas
    df = pd.read_excel(arquivo_destino, header=header_row)
    df = df[df["Pedido"].notna()]
    df_unique = df.drop_duplicates(subset="Pedido", keep="first")
    df_formatado = df_unique[colunas_desejadas].copy()
    df_formatado.columns = ['PEDIDO', 'ENTREGA', 'CLIENTE', 'PRODUTO', 'QUANTIDADE']
    
    # Converte PEDIDO para string e remove parte decimal
    df_formatado['PEDIDO'] = df_formatado['PEDIDO'].astype(str).str.strip()
    df_formatado['PEDIDO'] = df_formatado['PEDIDO'].str.replace('.0$', '', regex=True)
    
    df_formatado['ENTREGA'] = pd.to_datetime(df_formatado['ENTREGA'])
    hoje = pd.Timestamp(datetime.today().date())
    
    # Exibe o resultado
    df_formatado['ENTREGA'] = df_formatado['ENTREGA'].dt.strftime('%d/%m/%Y')
    print("\n✅ Dados formatados com sucesso:\n")
    print(tabulate(df_formatado, headers='keys', tablefmt='grid', showindex=False))
    
    # Exibe a coluna PEDIDO (agora como string sem parte decimal)
    print("\nColuna PEDIDO formatada:")
    
    produtos_unicos = df_formatado['PRODUTO'].unique().tolist()
    return df_formatado, produtos_unicos