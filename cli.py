from core import criarPlanilha, preencherPlanilha

while True:
    print("===== EXCEL MANAGER CLI VERSION =====")
    print("1. CRIAR PLANILHA")
    print("2. PREENCHER PLANILHA")
    opt = int(input("> "))

    if opt == 1:
        try:
            data = input("Escreva os dados separando por VIRGULA ( ex: nome,idade,joao,maria,32,32): ")
            valores = [v.strip() for v in data.split(",")]

            colunas = int(input("Quantas colunas deve ter: "))

            nome_arquivo = input("Qual o nome desejado? ( Vazio = Resultado.xlsx ):")
        except ValueError:
            print("Número invalido")
            exit()

        criarPlanilha(valores, colunas, nome_arquivo)
        break
    elif opt == 2:
        try:
            path = input("Caminho do arquivo: ")

            data = input("Escreva os dados separando por VIRGULA ( ex: joao,32, maria, 32): ")
            valores = [v.strip() for v in data.split(",")]

            colunas = int(input("Quantas colunas deve ter: "))

            linha_inicial = 2

            preencherPlanilha(path, valores, colunas, linha_inicial)
            break
        except ValueError:
            print("Valores invalidos")
            exit()
    else:
        print("Opções invalidas")
