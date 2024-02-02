from openpyxl import load_workbook  # load_workbook permite abrir e carregar um arquivo Excel existente para manipulação.
folderPath = "C:\\Users\\ramon\\Documents\\GitHub\\analiseDadosManutencao\\Análise planilha\\"
fileName = "ListaGeral.xlsx"

workbook = load_workbook(folderPath + fileName)
mainTabSheet = workbook["planilhaGeral"]

colecao = []

for i in range(0,1):
    row_values = mainTabSheet.iter_rows(min_row=0,max_row=10,min_col=0,max_col=0,values_only=True)
    row_values_list = list(row_values)
    colecao.append(row_values_list)
    print(i)
print(colecao)
workbook.close()
print("---fim do programa------")


