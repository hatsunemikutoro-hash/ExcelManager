from core import criarPlanilha

data = input("Escreva os dados separando por VIRGULA: ")
valores = [v.strip() for v in data.split(",")]

try:
    colunas = int(input("Quantas colunas deve ter: "))

    nome_arquivo = input("Qual o nome desejado? ( Vazio = Resultado.xlsx :")
except ValueError:
    print("Número invalido")
    exit()

criarPlanilha(valores, colunas, nome_arquivo)