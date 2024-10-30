from common import main

class Menu:
    
    def comum():
        opcaoMenus = input("Deseja voltar ao menu ? S/N  ")
        opcaoMenus = opcaoMenus.upper()
        print(opcaoMenus)
        if opcaoMenus != 'S':
            print("Código encerrado")
            exit()

    def menuPrincipal():
        Menu.comum()
        main()

    def menuProdutos():
        from produto import Produto
        Menu.comum()
        Produto.menuProdutos()

    def menuPessoas(opcaoMenu):
        from pessoas import Pessoa
        Menu.comum()
        Pessoa.menuPessoas(opcaoMenu)
    
    def menuVenda():
        from vendas import Venda
        Menu.comum()
        Venda.menuVendas()