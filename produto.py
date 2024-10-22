import os
import openpyxl
from menus import Menu


produtosCadastrados = openpyxl.load_workbook('Produtos.xlsx')
dfProdutosCadastrados = 'Produtos.xlsx'

produtosCadastrado = produtosCadastrados.active


class Produto:
    def __init__(self, nome, tipo, codigo,preco, quantidade):
        self.nome = nome
        self.tipo = tipo
        self.codigo = codigo
        self.preco = preco
        self.quantidade = quantidade

    def exibirProduto(self):
            print("---------------------------------------")
            print(f"Nome: {self.nome}\nTipo: {self.tipo}\nCódigo: {self.codigo}\nQuantidade: {self.quantidade}")
            print("---------------------------------------")

    def menuProdutos():
        os.system('cls')
        print("----------MENU PRODUTOS----------")

        print("[1]Produtos cadastrados")
        print("[2]Excluir produto")
        print("[3]Cadastrar novo produto")
        print("[4]Voltar ao menu principal")

        print("---------------------------------")

        escolhaMenuProdutos = int(input("Selecione uma opção:"))
        
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
            from common import menu
            os.system('cls')
            "Opção invalida..."
            menu()
    
    def opcaoProdCadastrados():
        Produto.imprimirPlanilha()
        Menu.menuProdutos()

    def opcaoExcluirProd():
        from common import procurarCelula
        Produto.imprimirPlanilha()
        
        escolhaDeletarLinha = input("Escolhe o produto que deseja excluir: ")
        
        celulaProcurada = escolhaDeletarLinha
        planilhas = produtosCadastrado
        
        linhaProdDeletar = procurarCelula(planilhas, celulaProcurada)

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
        
        print("Cadastre seu produto: ")

        nomeProduto = str(input("Nome: "))

        os.system('cls')

        print("Tipo [1] Alimento")
        print("Tipo [2] Higiene")
        print("Tipo [3] Outros")

        tipoProduto = int(input("Tipo: "))

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

        codigoProduto = (input("Código: "))
        
        verifNumeros = codigoProduto

        Produto.verificarValoresProd(verifNumeros)

        os.system('cls')

        precoProduto = input("Preço unitário (R$): ")

        verifNumeros = precoProduto
        
        Produto.verificarValoresProd(verifNumeros)

        os.system('cls')

        quantidadeProduto = input("Quantidade (PC): ")
        
        verifNumeros = quantidadeProduto

        Produto.verificarValoresProd(verifNumeros)

        os.system('cls')

        produto = Produto(nomeProduto, tipoProduto, codigoProduto, precoProduto, quantidadeProduto)

        produto_dados = {
            'Nome': produto.nome,
            'Tipo': produto.tipo,
            'Preço': produto.preco,
            'Quantidade': produto.quantidade,
            'codigo': produto.codigo
            
        }
        
        produtosCadastrado.append(list(produto_dados.values()))

        produtosCadastrados.save('Produtos.xlsx')
        produto.exibirProduto() 
        Menu.menuProdutos()
    
    def imprimirPlanilha():
        from common import imprimirTabelas
        planilhas = dfProdutosCadastrados
        verifPlanilha = 1
        imprimirTabelas(planilhas, verifPlanilha)
        
    def verificarValoresProd(verifNumeros):
        from common import verificarNumeros
        if verificarNumeros(verifNumeros) == False:
            Produto.opcaoCadastroProd()