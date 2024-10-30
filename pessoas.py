import os
import openpyxl
from menus import Menu
import random
from datetime import datetime



class Pessoa:
    
    #FUNÇÕES PRINCIPAIS
    
    def __init__(self, nome,idade, dataAniversario, contato, cadastro):
        self.nome = nome
        self.idade = idade
        self.dataAniversario = dataAniversario
        self.contato = contato
        self.cadastro = cadastro

    def exibirPessoas(self):
        print("---------------------------------------")
        print(f"\033[1;32mNome: \033[0m{self.nome}")
        print(f"\033[1;32mIdade: \033[0m{self.idade}")
        print(f"\033[1;32mData de Aniversário: \033[0m{self.dataAniversario}")
        print(f"\033[1;32mContato: \033[0m{self.contato}")
        print(f"\033[1;32mCadastro: \033[0m{self.cadastro}")
        

    def menuPessoas(opcaoMenu):
        os.system('cls')
        
        if opcaoMenu == 2:
            opcaoPessoa ="CLIENTES"
        else:
            opcaoPessoa = "VENDEDORES"

        print(f"\033[1;34m----------MENU {opcaoPessoa.upper()}----------\033[0m")
        print(f"\033[1;33m[1] {opcaoPessoa.capitalize()} cadastrados\033[0m")
        print(f"\033[1;33m[2] Excluir {opcaoPessoa}\033[0m")
        print(f"\033[1;33m[3] Cadastrar novos {opcaoPessoa}\033[0m")
        print("\033[1;33m[4] Voltar ao menu principal\033[0m")
        print("\033[1;34m---------------------------------\033[0m")

        escolhaMenuPessoa = int(input("Selecione uma opção: "))

        if escolhaMenuPessoa == 1:
            os.system('cls')
            Pessoa.opcaoProdCadastros(opcaoMenu)
        elif escolhaMenuPessoa == 2:
            os.system('cls')
            Pessoa.opcaoExlcuirPessoa(opcaoMenu, opcaoPessoa)
        elif escolhaMenuPessoa == 3:
            os.system('cls')
            Pessoa.opcaoCadastroPessoa(opcaoMenu, opcaoPessoa)
        elif escolhaMenuPessoa == 4:
            os.system('cls')
            Menu.menuPrincipal()
        else:
            from common import main
            os.system('cls')
            "Opção invalida..."
            main()
        

    def opcaoProdCadastros(opcaoMenu):
        _,_, dfPessoasCadastradas = Pessoa.planilhas(opcaoMenu)
        Pessoa.imprimirPlanilha(dfPessoasCadastradas, opcaoMenu)
        Menu.menuPessoas(opcaoMenu)

    def opcaoExlcuirPessoa(opcaoMenu, opcaoPessoa):
        from common import procurarCelula, menu
        
        pessoaCadastrada,pessoasCadastradas, dfPessoasCadastradas = Pessoa.planilhas(opcaoMenu)
        
        Pessoa.imprimirPlanilha(opcaoMenu, dfPessoasCadastradas)

        escolhaDeletarPessoa = input(f"escolha o {opcaoPessoa} que deseja excluir: ")

        celulaProcurada = escolhaDeletarPessoa
        planilhas = pessoaCadastrada
        
        escolhaDeletarPessoa = procurarCelula(planilhas, celulaProcurada)
        
        if escolhaDeletarPessoa == None:
            os.system('cls')
            print(f"{opcaoPessoa} não encontrado...")
            menu()

        pessoaCadastrada.delete_rows(escolhaDeletarPessoa)
        if opcaoMenu == 2:
            pessoasCadastradas.save('Clientes.xlsx')
        else:
            pessoasCadastradas.save('Vendedores.xlsx')
        

    def opcaoCadastroPessoa(opcaoMenu,opcaoPessoa):
        pessoaCadastrada,pessoasCadastradas, _ = Pessoa.planilhas(opcaoMenu)
        
        print(f"\033[1;36mCadastre o {opcaoPessoa}\033[0m")
    
        nomePessoa = str(input("\033[1;32mNome: \033[0m")).lower()
        

        dataAniversario = str(input("\033[1;32mData de nascimento: \033[0m"))
        
        idade = Pessoa.validadorIdade(dataAniversario)
        
        
        contato = Pessoa.validadorTelefone()
        
        os.system('cls')
        
        cadastro = Pessoa.geradorCadastro(opcaoMenu, pessoaCadastrada)

        pessoa = Pessoa(nomePessoa, idade,dataAniversario, contato, cadastro)

        pessoa_dados = {
            'Nome' : pessoa.nome,
            'Idade' : pessoa.idade,
            'Data de aniversário' : pessoa.dataAniversario,
            'Contato': pessoa.contato,
            'Cadastro' : pessoa.cadastro

        }

        pessoaCadastrada.append(list(pessoa_dados.values()))
        if opcaoMenu == 2:
            pessoasCadastradas.save('Clientes.xlsx')
        else:
            pessoasCadastradas.save('Vendedores.xlsx')
        
        if opcaoMenu == 2:
            cliente = Cliente(nomePessoa, idade,dataAniversario, contato, cadastro, compra= None)
            cliente.exibirCliente()
            
        else:
            vendedor = Vendedor(nomePessoa, idade,dataAniversario, contato, cadastro, vendas= None)
            vendedor.exibirVendedor()
        print("---------------------------------------")



#FUNÇÕES AUXILIARES

    def geradorCadastro(opcaoMenu, pessoaCadastrada):
            from common import procurarCelula
            numero = ''.join(random.choices('0123456789', k=5))
            planilhas = pessoaCadastrada
            celulaProcurada = numero
            celula = procurarCelula(planilhas, celulaProcurada)
            if opcaoMenu == 2:
                cadastro = "CL" + numero
                if celula != None:
                    Pessoa.geradorCadastro(opcaoMenu, pessoaCadastrada)
            else:
                if celula != None:
                    Pessoa.geradorCadastro(opcaoMenu, pessoaCadastrada)
                cadastro = "VR" + numero
            
            return cadastro

    def validadorTelefone():
        contato = str((input("\033[1;32mDigite o telefone para contato [Somente números]: \033[0m")))
        
        if contato.isdigit():
            if len(contato) == 11:
                return contato
            else:
                print("\033[31mTelefone inválido....\033[0m") 
                Pessoa.validadorTelefone()
        else:
            print("\033[31mTelefone inválido....\033[0m") 
            Pessoa.validadorTelefone()

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

    def imprimirPlanilha(dfPessoasCadastradas, opcaoMenu):
        from common import imprimirTabelas
        planilhas = dfPessoasCadastradas
        if opcaoMenu == 2:
            verifPlanilha = 2
        else:
            verifPlanilha = 3
        imprimirTabelas(planilhas, verifPlanilha)

    def planilhas(opcaoMenu):
        if opcaoMenu == 2:
            pessoasCadastradas = openpyxl.load_workbook('Clientes.xlsx')
            dfPessoasCadastradas = 'Clientes.xlsx'
            pessoaCadastrada = pessoasCadastradas.active
            return pessoaCadastrada, pessoasCadastradas, dfPessoasCadastradas
        else:
            pessoasCadastradas = openpyxl.load_workbook('Vendedores.xlsx')
            dfPessoasCadastradas = 'Vendedores.xlsx'
            pessoaCadastrada = pessoasCadastradas.active
            return pessoaCadastrada, pessoasCadastradas, dfPessoasCadastradas
        



class Cliente(Pessoa):
    
    def __init__(self, nome, idade, dataAniversario, contato, cadastro, compra):
        super().__init__(nome, idade, dataAniversario, contato, cadastro)
        self.compras = compra
    
    def exibirCliente(self):
        super().exibirPessoas()
        print(f"\033[1;32mCompras: \033[0m{self.compras}")

class Vendedor(Pessoa):
    
    def __init__(self, nome, idade, dataAniversario, contato, cadastro, vendas):
        super().__init__( nome, idade, dataAniversario, contato, cadastro)
        self.vendas = vendas
    
    def exibirVendedor(self):
        super().exibirPessoas()
        print(f"\033[1;32mVendas: \033[0m{self.vendas}")