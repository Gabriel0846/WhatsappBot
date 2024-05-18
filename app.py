import openpyxl

workbook = openpyxl.load_workbook("clientes.xlsx")
pagina_clientes = workbook["Planilha1"]

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    print(nome)
    print(telefone)
    print(vencimento)