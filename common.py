import openpyxl
import os
from tabulate import tabulate
import pandas as pd


produtosCadastrados = openpyxl.load_workbook('Produtos.xlsx')

produtosCadastrado = produtosCadastrados.active

pd.set_option('future.no_silent_downcasting', True)

def menu():
    os.system('cls')
    
    print("Bem vindo")
    print("[1]Vendas")
    print("[2]Cadastrar cliente")
    print("[3]Cadastrar vendendor")
    print("[4]Menu de produtos")

    opcaoMenu = int(input("Selecione uma opção: "))
    
    return opcaoMenu




def imprimirTabelas(planilhas, verifPlanilha):
        planilha = pd.read_excel(planilhas, header=None, skiprows= 1)
        cabeçalho = planilha.iloc[0].fillna('').infer_objects(copy=False).tolist()
        planilha.columns = cabeçalho
        planilha = planilha.iloc[1:]
        planilha = planilha.fillna('').infer_objects(copy=False)
        
        tabela = tabulate(planilha, headers='keys', tablefmt='pretty', showindex=False)
        
        if verifPlanilha == 1:
            titulo = "PRODUTOS"
        elif verifPlanilha == 2:
            titulo = "VENDEDOR"
        elif verifPlanilha == 3:
            titulo = "CLIENTE"
        elif verifPlanilha == 4:
            titulo = "VENDAS"
        
        tableWidth = 50
        centeredTitulo = titulo.center(tableWidth)
        print(centeredTitulo)
        print(tabela)

def verificarNumeros(verifNumeros):

    if not verifNumeros.isdigit():
            os.system('cls')
            print("Entrada invalida... favor digitar somente números")
            verifNumeros = False
            return verifNumeros

def procurarCelula(planilhas, celulaProcurada):
    celula = None

    for row in range(1, planilhas.max_row + 1):  # Percorre todas as linhas
        for column in range(1, planilhas.max_column + 1):  # Percorre todas as colunas
            if planilhas.cell(row=row, column=column).value == celulaProcurada:
                celula = (row, column)  # Armazena a posição (linha, coluna)
                return celula
            
    

if __name__ == '__main__':
    while True:

        opcaoMenu = menu()

        if opcaoMenu == 1:
            pass

        elif opcaoMenu == 2 or opcaoMenu == 3:
            from pessoas import Pessoa
            Pessoa.menuPessoas(opcaoMenu)

        elif opcaoMenu == 4:
            from produto import Produto
            Produto.menuProdutos()
            
        else:
            print("Opção invalida....")
            continue














