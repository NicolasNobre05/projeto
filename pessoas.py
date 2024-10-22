import os
import openpyxl
from menus import Menu
import random
from datetime import datetime



class Pessoa:
    
    def __init__(self, nome,idade, dataAniversario, contato, cadastro):
        self.nome = nome
        self.idade = idade
        self.dataAniversario = dataAniversario
        self.contato = contato
        self.cadastro = cadastro

    def exibirPessoas(self):
        print("---------------------------------------")
        print(f"Nome: {self.nome}\nIdade: {self.idade}\nData de Aniversário: {self.dataAniversario}\nQuantidade: {self.contato}\nCadastro: {self.cadastro}")
        print("---------------------------------------")

    def menuPessoas(opcaoMenu):
        os.system('cls')
        
        if opcaoMenu == 2:
            opcaoPessoa ="CLIENTES"
            pessoasCadastradas = openpyxl.load_workbook('Clientes.xlsx')
            dfPessoasCadastradas = 'Clientes.xlsx'
            pessoaCadastrada = pessoasCadastradas
            pessoaCadastrada.active
        else:
            opcaoPessoa = "VENDEDORES"
            pessoasCadastradas = openpyxl.load_workbook('Vendedores.xlsx')
            dfPessoasCadastradas = 'Vendedores.xlsx'
            pessoaCadastrada = pessoasCadastradas
            pessoaCadastrada.active
        

        print(f"----------MENU {opcaoPessoa}----------")
        
        opcaoPessoa = opcaoPessoa.lower()
        print(f"[1]{opcaoPessoa.capitalize()} cadastrados")
        print(f"[2]Excluir {opcaoPessoa}")
        print(f"[3]Cadastrar novos {opcaoPessoa}")
        print("[4]Voltar ao menu principal")
    
        print("---------------------------------")

        escolhaMenuPessoa = int(input("Selecione uma opção: "))

        if escolhaMenuPessoa == 1:
            os.system('cls')
            
        elif escolhaMenuPessoa == 2:
            os.system('cls')

        elif escolhaMenuPessoa == 3:
            os.system('cls')
            Pessoa.opcaoCadastroPessoa(opcaoMenu, opcaoPessoa)
        elif escolhaMenuPessoa == 4:
            os.system('cls')
            Menu.menuPrincipal()
        else:
            from common import menu
            os.system('cls')
            "Opção invalida..."
            menu()
        
    def opcaoCadastroPessoa(opcaoMenu,opcaoPessoa):
        pessoaCadastrada,pessoasCadastradas = Pessoa.planilhas(opcaoMenu)
        
        print(f"Cadastre o {opcaoPessoa}")
        
        nomePessoa = str(input("Nome: "))
        
        os.system('cls')

        dataAniversario = str(input("Data de nascimento:"))
        
        idade = Pessoa.validadorIdade(dataAniversario)

        os.system('cls')

        contato = int(input("Digite o telefone para contato: [Somente números]"))
        
        def geradorCadastro(opcaoMenu):
            from common import procurarCelula
            numero = ''.join(random.choices('0123456789', k=11))
            planilhas = pessoaCadastrada
            celulaProcurada = numero
            celula = procurarCelula(planilhas, celulaProcurada)
            if opcaoMenu == 2:
                cadastro = "CL" + numero
                if celula != None:
                    geradorCadastro(opcaoMenu)
            else:
                if celula != None:
                    geradorCadastro(opcaoMenu)
                cadastro = "VR" + numero
            
            return cadastro
        
        cadastro = geradorCadastro(opcaoMenu)    


        pessoa = Pessoa(nomePessoa, dataAniversario, idade, contato, cadastro)

        pessoa_dados = {
            'Nome' : pessoa.nome,
            'Data de aniversário' : pessoa.dataAniversario,
            'Idade' : pessoa.idade,
            'Contato': pessoa.contato,
            'Cadastro' : pessoa.cadastro

        }

        pessoaCadastrada.append(list(pessoa_dados.values()))
        if opcaoMenu == 2:
            pessoasCadastradas.save('Clientes.xlsx')
        else:
            pessoasCadastradas.save('Vendedores.xlsx')


    def validadorIdade(dataAniversario):
            
        dataAtual = datetime.now()

        try:
            dataAniversario = datetime.strptime(dataAniversario,"%d/%m/%Y")
        except:
            os.system('cls')
            print("Formato incorreto! Use o formato dd/mm/aaaa.")
            Pessoa.opcaoCadastroPessoa()

        if dataAtual >= dataAniversario:
            diferencaData = dataAtual - dataAniversario
            idade = diferencaData.days // 365
            return idade
        else:
            idade = False
            os.system('cls')
            print("Data invalida...")
            Pessoa.opcaoCadastroPessoa()

    def planilhas(opcaoMenu):
        if opcaoMenu == 2:
            pessoasCadastradas = openpyxl.load_workbook('Clientes.xlsx')
            dfPessoasCadastradas = 'Clientes.xlsx'
            pessoaCadastrada = pessoasCadastradas.active
            return pessoaCadastrada, pessoasCadastradas
        else:
            pessoasCadastradas = openpyxl.load_workbook('Vendedores.xlsx')
            dfPessoasCadastradas = 'Vendedores.xlsx'
            pessoaCadastrada = pessoasCadastradas.active
            return pessoaCadastrada, pessoasCadastradas
        







class Cliente(Pessoa):
    
    def __init__(self, nome, idade, dataAniversario, contato, cadastro, compra):
        super().__init__(nome, idade, dataAniversario, contato, cadastro)
        self.compras = compra
    
    def exibirCliente(self):
        super().exibirPessoas()
        print(f"\nCompras: {self.compras}")

class Vendedor(Pessoa):
    
    def __init__(self, nome, idade, dataAniversario, contato, cadastro, vendas):
        super().__init__(nome, idade, dataAniversario, contato, cadastro)
        self.vendas = vendas
    
    def exibirCliente(self):
        super().exibirPessoas()
        print(f"\nVendas: {self.vendas}")