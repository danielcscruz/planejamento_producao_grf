import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet
from datetime import datetime, timedelta

SETOR_ORDEM = [
    'PCP', 'Separação MP', 'Corte manual', 'Impressão',
    'Estampa', 'Corte laser', 'Costura', 'Arremate', 'Embalagem'
]

def obter_proximos_dias_uteis(data_inicio, dias_necessarios, calendario_path='data/_CALENDARIO.csv'):
    """
    Obtém uma lista de dias úteis a partir do próximo dia útil após a data de início.
    """
    print(f"DEBUG - Buscando {dias_necessarios} dias úteis a partir de {data_inicio}")
    df_cal = pd.read_csv(calendario_path)
    
    # Converte a coluna DATA para datetime
    df_cal['DATA'] = pd.to_datetime(df_cal['DATA'], format="%d/%m/%Y")
    
    # Filtra apenas dias úteis
    dias_uteis = df_cal[df_cal['VALOR'] == 'UTIL']['DATA'].tolist()
    print(f"DEBUG - Total de dias úteis no calendário: {len(dias_uteis)}")
    
    # Converte data_inicio para datetime se for string
    if isinstance(data_inicio, str):
        data_inicio = datetime.strptime(data_inicio, "%d/%m/%Y")
    
    # Avança para o próximo dia após a data de início
    data_proxima = data_inicio + timedelta(days=1)
    print(f"DEBUG - Próxima data após data_inicio: {data_proxima}")
    
    # Encontra a posição do próximo dia útil após a data de início
    proxima_data_util_idx = 0
    encontrou = False
    for i, data in enumerate(dias_uteis):
        if data >= data_proxima:
            proxima_data_util_idx = i
            encontrou = True
            print(f"DEBUG - Próximo dia útil encontrado: {data} no índice {i}")
            break
    
    if not encontrou:
        print(f"DEBUG - ALERTA: Nenhum dia útil encontrado após {data_proxima}!")
        return []
    
    # Retorna os próximos dias úteis necessários
    dias_selecionados = dias_uteis[proxima_data_util_idx:proxima_data_util_idx + dias_necessarios]
    print(f"DEBUG - Dias úteis selecionados: {dias_selecionados}")
    return dias_selecionados

def encontrar_coluna_por_data(ws, data):
    """
    Encontra a coluna na planilha que corresponde à data fornecida.
    Verifica tanto formato string quanto objeto datetime.
    """
    data_str = data.strftime("%d/%m/%Y")
    print(f"DEBUG - Procurando coluna para a data: {data_str}")
    
    for col in range(8, ws.max_column + 1):  # Começando da coluna H (8)
        cell_value = ws.cell(row=2, column=col).value
        # Limita o debug para não sobrecarregar o console
        if col % 10 == 0:  # Mostra debug a cada 10 colunas
            print(f"DEBUG - Verificando células... coluna atual {col}")
        
        # Verifica se o valor da célula é uma data
        if isinstance(cell_value, datetime):
            # Compara o valor da data
            if cell_value.date() == data.date():
                print(f"DEBUG - Data encontrada na coluna {col} (formato datetime)")
                return col
        elif isinstance(cell_value, str):
            # Tenta converter para datetime se for string
            try:
                # Tenta primeiro o formato DD/MM/YYYY
                cell_date = datetime.strptime(cell_value, "%d/%m/%Y")
                if cell_date.date() == data.date():
                    print(f"DEBUG - Data encontrada na coluna {col} (formato string DD/MM/YYYY)")
                    return col
            except ValueError:
                try:
                    # Tenta o formato YYYY-MM-DD
                    cell_date = datetime.strptime(cell_value, "%Y-%m-%d")
                    if cell_date.date() == data.date():
                        print(f"DEBUG - Data encontrada na coluna {col} (formato string YYYY-MM-DD)")
                        return col
                except ValueError:
                    # Ignora se não conseguir converter
                    pass
    
    print(f"DEBUG - ALERTA: Data {data_str} não encontrada em nenhuma coluna da planilha!")
    return None

def preencher_producao(ws: Worksheet, quantidade: int, setor: str, linha: int, 
                      data_inicio=None, calendario_path='data/_CALENDARIO.csv'):
    """
    Preenche a produção a partir do setor especificado, propagando para os demais setores na ordem.
    """
    print(f"\nDEBUG - Iniciando planejamento: quantidade={quantidade}, setor={setor}, linha={linha}")
    
    # Verifica se o arquivo de calendário existe
    print(f"DEBUG - Verificando calendário: {calendario_path}")
    try:
        df_cal = pd.read_csv(calendario_path)
        print(f"DEBUG - Calendário carregado com sucesso. {len(df_cal)} registros.")
    except Exception as e:
        print(f"DEBUG - ERRO ao carregar calendário: {e}")
        return
    
    df_cal['DATA'] = pd.to_datetime(df_cal['DATA'], format="%d/%m/%Y")
    
    # Verifica se o setor é válido
    if setor not in SETOR_ORDEM:
        print(f"DEBUG - ERRO: Setor '{setor}' não reconhecido.")
        raise ValueError(f"Setor '{setor}' não reconhecido. Deve ser um dos: {SETOR_ORDEM}")
    
    # Determina o índice do setor inicial na ordem de processamento
    setor_idx = SETOR_ORDEM.index(setor)
    setores_processar = SETOR_ORDEM[setor_idx:]
    print(f"DEBUG - Setores a processar: {setores_processar}")
    
    # Se data_inicio não for fornecida, usa a data atual
    if data_inicio is None:
        data_inicio = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        print(f"DEBUG - Data de início não fornecida, usando data atual: {data_inicio}")
    elif isinstance(data_inicio, str):
        data_inicio = datetime.strptime(data_inicio, "%d/%m/%Y")
        print(f"DEBUG - Data de início convertida de string: {data_inicio}")
    
    # Para rastrear a data atual de processamento para cada setor
    data_atual = data_inicio
    print(f"DEBUG - Data inicial para processamento: {data_atual}")
    
    # Processa cada setor na ordem
    for i, setor_nome in enumerate(setores_processar):
        print(f"\nDEBUG - Processando setor: {setor_nome} (índice {i})")
        
        # Se não for o primeiro setor, avança para o próximo dia útil
        if i > 0:
            print(f"DEBUG - Avançando para o próximo dia útil após {data_atual}")
            
            # Avança para o próximo dia útil - CORRIGIDO PARA EVITAR LOOP INFINITO
            data_atual += timedelta(days=1)
            dias_avancados = 1
            ultima_data_verificada = None
            
            while True:
                # Verifica se estamos verificando a mesma data repetidamente (sinal de loop)
                if ultima_data_verificada == data_atual.date():
                    print(f"DEBUG - ALERTA: Detectado possível loop na data {data_atual.date()}")
                    # Avança mais um dia para quebrar o padrão
                    data_atual += timedelta(days=1)
                
                ultima_data_verificada = data_atual.date()
                
                matching_rows = df_cal[(df_cal['DATA'].dt.date == data_atual.date())]
                if not matching_rows.empty:
                    print(f"DEBUG - Verificando data {data_atual.date()}: valor={matching_rows['VALOR'].iloc[0]}")
                    if matching_rows['VALOR'].iloc[0] == 'UTIL':
                        print(f"DEBUG - Dia útil encontrado: {data_atual.date()}")
                        break
                else:
                    print(f"DEBUG - Data {data_atual.date()} não encontrada no calendário!")
                
                data_atual += timedelta(days=1)
                dias_avancados += 1
                if dias_avancados > 30:  # Segurança para evitar loop infinito
                    print("DEBUG - ALERTA: Mais de 30 dias avançados sem encontrar dia útil! Interrompendo.")
                    break
            
            print(f"DEBUG - Novo dia útil para processamento: {data_atual}")
        
        # Calcula a linha do setor atual (linha do pedido + offset do setor)
        linha_setor = linha + SETOR_ORDEM.index(setor_nome) + 1
        print(f"DEBUG - Linha do setor {setor_nome}: {linha_setor}")
        
        # Obtém o limite de produção diário para este setor
        linha_limite = SETOR_ORDEM.index(setor_nome) + 3  # +3 para corresponder à posição na planilha
        try:
            valor_limite_cell = ws.cell(row=linha_limite, column=5).value
            print(f"DEBUG - Célula de limite na linha {linha_limite}, coluna 5: {valor_limite_cell}")
            valor_limite = int(valor_limite_cell) if valor_limite_cell is not None else 0
        except (ValueError, TypeError) as e:
            print(f"DEBUG - Erro ao converter limite: {e}")
            valor_limite = 0
        
        print(f"DEBUG - Limite diário para o setor {setor_nome}: {valor_limite}")
        
        qtd_restante = quantidade
        dias_necessarios = 0
        
        # Calcula quantos dias serão necessários para processar toda a quantidade
        if valor_limite > 0:
            dias_necessarios = (qtd_restante + valor_limite - 1) // valor_limite  # Arredonda para cima
        else:
            dias_necessarios = 1  # Se não há limite, assume que tudo pode ser feito em um dia
        
        print(f"DEBUG - Dias necessários para processar {qtd_restante} unidades: {dias_necessarios}")
        
        # Obtém os próximos dias úteis necessários
        dias_uteis = obter_proximos_dias_uteis(data_atual, dias_necessarios, calendario_path)
        
        if not dias_uteis:
            print(f"DEBUG - ALERTA: Nenhum dia útil encontrado para o setor {setor_nome}!")
            continue
        
        print(f"DEBUG - Distribuindo produção para o setor {setor_nome}")
        # Distribui a produção pelos dias úteis
        for dia_util in dias_uteis:
            # Encontra a coluna correspondente à data
            col = encontrar_coluna_por_data(ws, dia_util)
            
            if col is None:
                print(f"DEBUG - Pulando dia {dia_util} pois não foi encontrado na planilha")
                continue
            
            # Determina quanto produzir neste dia
            producao_dia = min(valor_limite, qtd_restante) if valor_limite > 0 else qtd_restante
            
            # Registra a produção na planilha
            print(f"DEBUG - Preenchendo: linha={linha_setor}, coluna={col}, valor={producao_dia}")
            current_value = ws.cell(row=linha_setor, column=col).value
            print(f"DEBUG - Valor atual na célula: {current_value}")
            
            try:
                ws.cell(row=linha_setor, column=col, value=producao_dia)
                print(f"DEBUG - Produção registrada com sucesso: {producao_dia} unidades")
            except Exception as e:
                print(f"DEBUG - ERRO ao registrar produção: {e}")
            
            qtd_restante -= producao_dia
            print(f"DEBUG - Quantidade restante: {qtd_restante}")
            
            # Se toda a quantidade foi distribuída, para
            if qtd_restante <= 0:
                print(f"DEBUG - Toda a quantidade foi distribuída para o setor {setor_nome}")
                break
        
        # Atualiza a data atual para o último dia utilizado neste setor
        if dias_uteis:
            data_atual = dias_uteis[-1]
            print(f"DEBUG - Data atual atualizada para o último dia utilizado: {data_atual}")
    
    print("DEBUG - Planejamento de produção concluído")