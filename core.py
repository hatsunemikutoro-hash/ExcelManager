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
# ja coloquei

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
# ja coloquei
def contarLinhas(arquivo):
    try:
        wb = openpyxl.load_workbook(arquivo)
        ws = wb.active

        for row in range(ws.max_row, 0, -1):
            if any(ws.cell(row=row, column=col).value is not None for col in range(1, ws.max_column + 1)):
                return row
        return 0
    except Exception as e:
        return f"erro ao abrir o arquivo{arquivo}"


def lerCabecalhos(arquivo):
    try:
        # data_only=True evita ler fórmulas, pega só o resultado
        wb = openpyxl.load_workbook(arquivo, data_only=True)
        aba = wb.active

        # Usamos um gerador para não travar se a planilha for gigante
        gerador_linhas = aba.iter_rows(min_row=1, max_row=1, values_only=True)

        # O try/except aqui evita o erro de 'StopIteration' se a aba estiver vazia
        try:
            primeira_linha = next(gerador_linhas)
        except StopIteration:
            return "A planilha está completamente vazia."

        # Filtra e limpa os dados
        cabecalho_filtrado = [str(item).strip() for item in primeira_linha if item is not None]

        if not cabecalho_filtrado:
            return "Nenhum cabeçalho encontrado na primeira linha."

        # Retorna a string formatada pronta para o seu Output da UI
        return " | ".join(cabecalho_filtrado)

    except Exception as e:
        return f"Erro ao abrir o arquivo {arquivo}: {e}"

def buscarRegistro(arquivo, filtro):
    try:
        wb = openpyxl.load_workbook(arquivo)
        ws = wb.active

        for linha in ws.iter_rows():
            for celula in linha:
                value = celula.value

                value_str = str(value) if value is not None else ""
                coluna = celula.column_letter

                if filtro.lower() in value_str.lower():
                    return f"{value_str} encontrado na coluna {coluna}"

        return f"O termo '{filtro}' não foi encontrado em nenhuma célula."
    except Exception as e:
        return f"{e}"
# ja coloquei

def listarPlanilha(arquivo):
    try:
        wb = openpyxl.load_workbook(arquivo)

        return wb.sheetnames
    except Exception as e:
        return f"Erro: {e}"