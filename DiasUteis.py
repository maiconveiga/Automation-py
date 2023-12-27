import pandas as pd
from datetime import date, timedelta
from openpyxl import Workbook
import os

def calcular_dias_uteis(ano, mes):
    primeiro_dia = date(ano, mes, 1)
    ultimo_dia = date(ano, mes, 1) + timedelta(days=32)
    ultimo_dia = ultimo_dia.replace(day=1) - timedelta(days=1)
    dias_uteis = 0

    # Loop pelos dias do mês
    for dia in range((ultimo_dia - primeiro_dia).days + 1):
        data = primeiro_dia + timedelta(days=dia)

        # Verificar se o dia é útil (não é sábado nem domingo)
        if data.weekday() < 5:
            dias_uteis += 1

    return dias_uteis

def gerar_excel():
    # Criar um DataFrame vazio
    df = pd.DataFrame(columns=['Ano', 'Mês', 'Quantidade de Dias Úteis'])

    # Loop pelos meses do ano de 2023
    for mes in range(1, 13):
        dias_uteis = calcular_dias_uteis(2023, mes)

        # Adicionar dados ao DataFrame
        df.loc[len(df)] = {'Ano': 2023, 'Mês': mes, 'Quantidade de Dias Úteis': dias_uteis}

    # Criar um arquivo Excel e salvar o DataFrame nele
    wb = Workbook()
    ws = wb.active

    # Adicionar cabeçalho
    cabecalho = ['Ano', 'Mês', 'Quantidade de Dias Úteis']
    ws.append(cabecalho)

    # Adicionar os dados ao Excel
    for row in df.itertuples(index=False, name=None):
        ws.append(row)

    # Determinar o caminho para a área de trabalho
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')

    # Salvar o arquivo Excel na área de trabalho
    arquivo_excel = os.path.join(desktop_path, 'dias_uteis_2023.xlsx')
    wb.save(arquivo_excel)

# Chamar a função para gerar o Excel
gerar_excel()
