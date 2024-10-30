import os
import openpyxl
from menus import Menu
import random

produtosCadastrados = openpyxl.load_workbook('Produtos.xlsx')
dfProdutosCadastrados = 'Produtos.xlsx'

produtosCadastrado = produtosCadastrados.active  


class Produto:

    #FUNÇÕES PRINCIPAIS
    
    def __init__(self, nome, tipo, codigo,preco, quantidade, estoqueMinimo, estoqueMaximo):
        self.nome = nome
        self.tipo = tipo
        self.codigo = codigo
        self.preco = preco
        self.quantidade = quantidade
        self.estoqueMinimo = estoqueMinimo
        self.estoqueMaximo = estoqueMaximo

    def exibirProduto(self):
            print("\033[1;34m" + "-" * 39 + "\033[0m")
            print(f"\033[1;32mNome: \033[0m{self.nome}")
            print(f"\033[1;32mTipo: \033[0m{self.tipo}")
            print(f"\033[1;32mCódigo: \033[0m{self.codigo}")
            print(f"\033[1;32mQuantidade: \033[0m{self.quantidade}")
            print("\033[1;34m" + "-" * 39 + "\033[0m")
    
    def menuProdutos():
        os.system('cls')
        print("\033[1;34m" + "----------MENU PRODUTOS----------" + "\033[0m")
        print("\033[1;33m" + "[1] Produtos cadastrados" + "\033[0m")
        print("\033[1;33m" + "[2] Excluir produto" + "\033[0m")
        print("\033[1;33m" + "[3] Cadastrar novo produto" + "\033[0m")
        print("\033[1;33m" + "[4] Voltar ao menu principal" + "\033[0m")
        print("\033[1;34m" + "---------------------------------" + "\033[0m")
        
        escolhaMenuProdutos = int(input("Selecione uma opção: "))
        
        if escolhaMenuProdutos == 1:
            os.system('cls')
            Produto.opcaoProdCadastrados()
        elif escolhaMenuProdutos == 2:
            os.system('cls')
            Produto.opcaoExcluirProd()
        elif escolhaMenuProdutos == 3:
            os.system('cls')
            Produto.opcaoCadastroProd()
        elif escolhaMenuProdutos == 4:
            os.system('cls')
            Menu.menuPrincipal()
        else:
            from common import main 
            os.system('cls')
            "Opção invalida..."
            main()
    
    def opcaoProdCadastrados():
        Produto.imprimirPlanilha()
        Menu.menuProdutos()

    def opcaoExcluirProd():
        from common import procurarCelula, main
        Produto.imprimirPlanilha()
        
        escolhaDeletarLinha = input("Escolhe o produto que deseja excluir: ")
        
        celulaProcurada = escolhaDeletarLinha
        planilhas = produtosCadastrado
        
        linhaProdDeletar = procurarCelula(planilhas, celulaProcurada)

        if linhaProdDeletar == None:
            os.system('cls')
            print("Produto não encontrado...")
            main()
            
            

        produtosCadastrado.delete_rows(linhaProdDeletar)
        produtosCadastrados.save('Produtos.xlsx')
        os.system('cls')
        print(f"linha {escolhaDeletarLinha} deletada.")
        Produto.imprimirPlanilha()
        escolhaDeletarLinha = input("Deseja deletar outro produto? S/N  ")
        escolhaDeletarLinha = escolhaDeletarLinha.upper()
        if escolhaDeletarLinha == 'S':
            os.system('cls')
            Produto.opcaoExcluirProd()
        os.system('cls')
        Menu.menuProdutos()

    def opcaoCadastroProd():
        
        print("\033[1;34m" + "Cadastre seu produto:" + "\033[0m")

        nomeProduto = input("Nome: ").lower()

        os.system('cls')

        print("\033[1;34m" + "Selecione o tipo do produto:" + "\033[0m")
        print("\033[1;32m" + "[1] Alimento" + "\033[0m")
        print("\033[1;32m" + "[2] Higiene" + "\033[0m")
        print("\033[1;32m" + "[3] Outros" + "\033[0m")

        tipoProduto = int(input("Tipo: "))
    
        codigoProduto = Produto.geradorCadastro(produtosCadastrado, tipoProduto)

        if tipoProduto == 1:
            tipoProduto = "Alimento"
        elif tipoProduto == 2:
            tipoProduto = "Higiene"
        elif tipoProduto == 3:
            tipoProduto = "Outros"
        else:
            print("Opção invalida....")
            Produto.opcaoCadastroProd()


        os.system('cls')

        precoProduto = input("Preço unitário (R$): ").strip()
        precoProduto = precoProduto.replace(',','.')
        precoProduto = float(precoProduto)

        if precoProduto < 0:
            print("Valor invalido....")
            Produto.opcaoCadastroProd()
        

        os.system('cls')

        quantidadeProduto = input("Quantidade (PC): ")

        estoqueMinimo = input("Estoque Mínimo:")
        estoqueMaximo = input("Estoque Maximo:")
        
        verifNumeros = quantidadeProduto

        Produto.verificarValoresProd(verifNumeros)

        os.system('cls')

        produto = Produto(nomeProduto, tipoProduto, codigoProduto, precoProduto, quantidadeProduto, estoqueMinimo, estoqueMaximo)

        produto_dados = {
            'Nome': produto.nome,
            'Tipo': produto.tipo,
            'Preço': produto.preco,
            'Quantidade': produto.quantidade,
            'codigo': produto.codigo,
            'EstoqueMinimo' : produto.estoqueMinimo,
            'EstoqueMaximo' : produto.estoqueMaximo
            
        }
        
        produtosCadastrado.append(list(produto_dados.values()))

        produtosCadastrados.save('Produtos.xlsx')
        produto.exibirProduto() 
    
    ##FUNÇÕES AUXILIARES

    def imprimirPlanilha():
        from common import imprimirTabelas
        planilhas = dfProdutosCadastrados
        verifPlanilha = 4
        imprimirTabelas(planilhas, verifPlanilha)
        
    def verificarValoresProd(verifNumeros):
        from common import verificarNumeros
        if verificarNumeros(verifNumeros) == False:
            Produto.opcaoCadastroProd()
    
    def geradorCadastro(produtosCadastrado, tipoProduto):
            from common import procurarCelula

            numero = ''.join(random.choices('0123456789', k=5))
            planilhas = produtosCadastrado
            celulaProcurada = numero
            celula = procurarCelula(planilhas, celulaProcurada)
            
            if tipoProduto == 1:
                cadastro = "AL" + numero
                if celula != None:
                    Produto.geradorCadastro( produtosCadastrado)
            elif tipoProduto == 2:
                cadastro = "HI" + numero
                if celula != None:
                    Produto.geradorCadastro( produtosCadastrado)
            elif tipoProduto == 3:
                cadastro = "OU" + numero
                if celula != None:
                    Produto.geradorCadastro( produtosCadastrado)
            
            return cadastro