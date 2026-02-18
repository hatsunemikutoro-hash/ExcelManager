from openpyxl import Workbook
from openpyxl import load_workbook

def criarPlanilha(lista, colunas, nome_arquivo):

    wb = Workbook()
    ws = wb.active

    if not nome_arquivo:
        nome_arquivo = "resultado"

    linha=1
    coluna=1

    for item in lista:
        ws.cell(row=linha, column=coluna, value=item)
        coluna += 1

        if coluna > colunas:
            coluna = 1
            linha += 1

    wb.save(nome_arquivo + ".xlsx")
    return nome_arquivo + ".xlsx"

def preencherPlanilha(path, lista, colunas, linha_inicial=2):
    wb = load_workbook(path)
    ws = wb.active

    linha = linha_inicial
    coluna = 1

    for item in lista:
        ws.cell(row=linha, column=coluna, value=item)
        coluna += 1

        if coluna > colunas:
            coluna = 1
            linha += 1

    wb.save(path)
    return path