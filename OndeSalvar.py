import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
print('executando')

def salvar_arquivo_excel(conteudo, caminho):
    print('Salvar excel')
    # Cria um novo workbook e adiciona uma folha de trabalho
    workbook = Workbook()
    sheet = workbook.active
    
    # Adiciona o conteúdo à folha de trabalho
    for linha in conteudo:
        sheet.append(linha)
    
    # Salva o arquivo no caminho fornecido
    workbook.save(caminho)

def selecionar_caminho():
    print('Selecionar')
    # Cria uma janela para seleção de arquivo
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    
    # Pede ao usuário que selecione o local de salvamento
    caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    
    # Se o usuário cancelar a seleção, retorne None
    if not caminho_arquivo:
        return None
    
    return caminho_arquivo

def main():
    print('Main')
    # Chama a função para selecionar o caminho do arquivo
    caminho_arquivo = selecionar_caminho()
    
    # Se o usuário cancelar a seleção, saia do programa
    if caminho_arquivo is None:
        print("Operação cancelada pelo usuário.")
        return
    
    # Conteúdo que você deseja salvar no arquivo Excel
    conteudo_excel = [
        ["Nome", "Idade", "Cidade"],
        ["João", 25, "São Paulo"],
        ["Maria", 30, "Rio de Janeiro"],
        ["Carlos", 22, "Belo Horizonte"]
    ]
    
    # Salva o arquivo Excel no local escolhido pelo usuário
    salvar_arquivo_excel(conteudo_excel, caminho_arquivo)
    
    print(f"Arquivo salvo com sucesso em: {caminho_arquivo}")

if __name__ == "__main__":
    print('Main validação')
    main()