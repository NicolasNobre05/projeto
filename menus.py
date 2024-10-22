from common import menu

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
        menu()

    def menuProdutos():
        from produto import Produto
        Menu.comum()
        Produto.menuProdutos()
    