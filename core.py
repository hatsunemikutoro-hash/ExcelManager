import openpyxl
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

def contarLinhas(arquivo):
    try:
        wb = openpyxl.load_workbook(arquivo)
        aba = wb.active

        total = aba.max_row

        return total
    except Exception as e:
        return f"erro ao abrir o arquivo{arquivo}"

def lerCabecalhos(arquivo):
    try:
        wb = openpyxl.load_workbook(arquivo)
        aba = wb.active


        #next pega a primeira linha
        primeira_linha = next(aba.iter_rows(min_row=1, max_row=1, values_only=True))

        cabecalho_filtrado = [str(item) for item in primeira_linha if item is not None]

        if not cabecalho_filtrado:
            return "A planilha parece estar vazia ou não possui cabeçalhos."

        return " | ".join(cabecalho_filtrado)
    except Exception as e:
        return f"erro ao abrir o arquivo{arquivo}"

def buscarRegistro(arquivo, filtro):
    try:
        wb = openpyxl.load_workbook(arquivo)
        ws = wb.active

        for linha in ws.iter_rows():
            for celula in linha:
                value = celula.value
                coluna = celula.column_letter

                if filtro in value:
                    return f"{value} encontrado na coluna {coluna}"

        return f"O termo '{filtro}' não foi encontrado em nenhuma célula."
    except Exception as e:
        return f"Erro ao abrir o arquivo {arquivo}"

def listarPlanilha(arquivo):
    try:
        wb = openpyxl.load_workbook(arquivo)

        return wb.sheetnames
    except Exception as e:
        return f"Erro: {e}"