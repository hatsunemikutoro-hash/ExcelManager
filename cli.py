from core import criarPlanilha, preencherPlanilha, contarLinhas, buscarRegistro, lerCabecalhos, listarPlanilha

while True:
    print("===== EXCEL MANAGER CLI VERSION =====")
    print("1. CRIAR PLANILHA")
    print("2. PREENCHER PLANILHA")
    print("3. CONTAR LINHAS")
    print("4. FILTRAR PALAVRA")
    print("5. LER CABECALHOS")
    print("6. LISTAR PLANILHA")
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
    elif opt == 3:
        try:
            path = input("Caminho do arquivo: ")
            print(f"O arquivo tem {contarLinhas(path)} linhas.")
        except Exception as e:
            print("Algum erro ocorreu.")
    elif opt == 4:
        try:
            path = input("Caminho do arquivo: ")
            filtro = input("Que palavra deseja buscar?: ")
            print(buscarRegistro(path, filtro))
        except Exception as e:
            print("Ocorreu um erro.")
    elif opt == 5:
        try:
            path = input("Caminho do arquivo: ")
            print(lerCabecalhos(path))
        except Exception as e:
            print("Ocorreu um erro.")
    elif opt == 6:
        try:
            path = input("Caminho do arquivo: ")
            print(listarPlanilha(path))
        except Exception as e:
            print("Ocorreu um erro.")
    else:
        print("Opções invalidas")
