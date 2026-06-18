from openpyxl import Workbook

#1 - Criar um novo arquivo Excel
wb = Workbook()
#2 - Selecionar a planilha ativa
ws = wb.active
ws.title = "Dados"

#3 - Adicionar dados à planilha
ws.append(["Nome", "Idade", "Cidade"])
ws.append(["Alice", 30, "São Paulo"])
ws.append(["Bob", 25, "Rio de Janeiro"])
ws.append(["Carol", 28, "Belo Horizonte"])

#4 - Salvar o arquivo Excel
wb.save("pronto/novo_arquivo.xlsx")
print('Arquivo Excel criado com sucesso!')