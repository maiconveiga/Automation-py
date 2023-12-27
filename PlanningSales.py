import pandas as pd
import datetime as dt
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook

#pyinstaller --onefile PlanningSales.py
print('')
print('')
print('')
print('')
print("#########################################################")
print("#   Planning Sales - Coleta e tratamento de planilhas   #")
print("#########################################################")
print('')
##########################################################################
if(format(dt.date.today(), "%m") == '1'):
    mes = 'jan'
elif(format(dt.date.today(), "%m") == '2'):
    mes = 'fev'
elif(format(dt.date.today(), "%m") == '3'):
    mes = 'mar'
elif(format(dt.date.today(), "%m") == '4'):
    mes = 'abr'
elif(format(dt.date.today(), "%m") == '5'):
    mes = 'maio'
elif(format(dt.date.today(), "%m") == '6'):
    mes = 'jun'
elif(format(dt.date.today(), "%m") == '7'):
    mes ='jul'
elif(format(dt.date.today(), "%m") == '8'):
    mes = 'ago'
elif(format(dt.date.today(), "%m") == '9'):
    mes = 'set'
elif(format(dt.date.today(), "%m") == '10'):
    mes = 'out'
elif(format(dt.date.today(), "%m") == '11'):
    mes = 'nov'
elif(format(dt.date.today(), "%m") == '12'):
    mes = 'dez'

ano = format(dt.date.today(), "%y")
numero = int(format(dt.date.today(), "%y"))+int(format(dt.date.today(), "%m"))+16

nomeParts = str(numero) + ". "+ str(mes) + "-" + str(ano) + " Parts"
nomeRas = str(numero) + ". "+ str(mes) + " " + str(ano) + " RAS"

print('-------------------------------------')
print(f"{nomeParts} -> Sendo criada")
print('-------------------------------------')

tabelaParts = pd.read_excel(r"\\C051mb30\spool_rpw_prod\Parts Center\OS\ESFT0001.xlsx")
tabelaParts = tabelaParts.drop("UN", axis = 1)
tabelaParts = tabelaParts.drop("POLÍTICA", axis = 1)
tabelaParts = tabelaParts.drop("CLASSIF POLIT", axis = 1)
tabelaParts = tabelaParts.drop("DESC CLASSIF POLIT", axis = 1)
tabelaParts = tabelaParts.drop("COND PAGTO CLIENTE", axis = 1)
tabelaParts = tabelaParts.drop("COND PAGTO NF", axis = 1)
tabelaParts = tabelaParts.rename(columns={' QUANTIDADE': 'QUANTIDADE '})
tabelaParts = tabelaParts.rename(columns={'VALOR BRUTO': 'VALOR BRUTO '})
tabelaParts = tabelaParts.rename(columns={'VALOR LÍQUIDO': 'VALOR LÍQUIDO '})
tabelaParts = tabelaParts.rename(columns={'COFINS (RET)': 'COFINS (RET) '})
tabelaParts = tabelaParts.rename(columns={'INSS (RET)': 'INSS (RET) '})
tabelaParts = tabelaParts.rename(columns={'ISS (RET)': 'ISS (RET) '})
tabelaParts = tabelaParts.rename(columns={'CUSTO MÉDIO DO ITEM': 'CUSTO MÉDIO DO ITEM '})
tabelaParts = tabelaParts.rename(columns={'CENTRO DE CUSTO': 'CENTRO DE CUSTO '})
tabelaParts = tabelaParts.rename(columns={'CÓDIGO CEST': 'CÓDIGO CEST '})
tabelaParts = tabelaParts.rename(columns={'SITUAÇÃO NOTA FISCAL': 'SITUAÇÃO NOTA FISCAL '})
tabelaParts = tabelaParts.rename(columns={'NR PED TRANSF': 'NR PED TRANSF.'})
tabelaParts["Período"] = mes+"/20"+ano

colunas = tabelaParts.columns.to_list()


tabelaParts = tabelaParts[['Período','ESTAB', 'CLIENTE', 'NOME', 'CIDADE', 'ESTADO', 'TIPO CLIENTE', 'VENDEDOR', 'DESC. VENDEDOR', 'NOME VENDEDOR', 'DATA', 'SERIE', 'NOTA FISCAL', 'PEDIDO', 'DATA PEDIDO', 'DEPÓSITO', 'ITEM', 'DESCRIÇÃO DO ITEM', 'TIPO DO ITEM', 'ACCONTING CLASS', 'CLASSIFICAÇÃO FISCAL', 'GRUPO ESTOQUE', 'DESC. GRUPO ESTOQUE', 'FAMILIA', 'DESCRIÇÃO FAMILIA', 'UNEG', 'CANAL DE VENDA', 'CFOP', 'DESCRIÇÃO CFOP', 'SINIIT', 'RTENTRAD', 'RTVENDA', 'RTNAC', 'RTMERC', 'QUANTIDADE ', 'VALOR BRUTO ', 'VALOR LÍQUIDO ', 'ICMS', 'DIFAL DESTINO', 'FECP DIFAL', 'ICMS ST', 'FCP ST', 'IPI', 'PIS', 'COFINS', 'COFINS (RET) ', 'INSS (RET) ', 'ISS', 'ISS (RET) ', 'IRRF (RET)', 'CSLL (RET)', 'PIS (RET)', 'CUSTO MÉDIO DO ITEM ', 'CONTA CONTÁBIL', 'CENTRO DE CUSTO ', 'STANDARD', 'CURRENT', 'PREÇO BASE', 'TERRIT', 'TXCLS', 'USUÁRIO NF', 'PROJETO', 'TIPO FRETE', 'ORIGEM', 'DESTINO MERCADORIA', 'TRANSPORTADORA', 'NOME TRANSPORTADOR', 'CÓDIGO CEST ', 'SITUAÇÃO NOTA FISCAL ', 'CHAVE ACESSO NF-e', 'NR PEDIDO', 'DT ENTREGA ORIG', 'CPF/CNPJ', 'DT RECONHECIMENTO', 'TABELA PREÇO', 'DT IMPLANTAÇÃO CLIENTE', 'LEAD TIME', 'FAMILIA COMERCIAL', 'DESCRIÇÃO FAMILIA COMERCIAL', 'DT NECESSIDADE', 'NR PED TRANSF.']]


writer = pd.ExcelWriter(rf"C:\ETL PlanningSales\{nomeParts}.xlsx")
tabelaParts.to_excel(writer, sheet_name='New Parts', index=False)
writer.close()


print('-------------------------------------')
print(f"{nomeParts} -> Criada")
print('-------------------------------------')

print('')
print('')

print('-------------------------------------')
print(f"{nomeRas} -> Sendo criada")
print('-------------------------------------')
tabelaRAS = pd.read_excel(r"\\C051mb30\spool_rpw_prod\Parts Center\OS\RAS121AC.xlsx")
tamanho = len(tabelaRAS.index)
tabelaRAS = tabelaRAS.drop(tamanho -1)
tabelaRAS = tabelaRAS.rename(columns={'VALOR LIQUIDO ORIGINAL': 'VALOR LÍQUIDO ORIGINAL'})
tabelaRAS = tabelaRAS.rename(columns={'CLASSIFICACAO': 'CLASSIFICAÇÃO'})

writer = pd.ExcelWriter(rf"C:\ETL PlanningSales\{nomeRas}.xlsx")
tabelaRAS.to_excel(writer, sheet_name='RAS', index=False)
writer.close()
print('-------------------------------------')
print(f"{nomeRas} -> Criada")
print('-------------------------------------')
print('')

##############################################################################