import xlrd


# Abre o arquivo xls
workbook = xlrd.open_workbook('produtos.xls')

# Especifica a folha
sheet = workbook.sheet_by_index(0)

# Cria o dicionario
data_dict = {}

# Variavel para controle de fileiras
n = 0

# Armazena as keys e values em um dict
while True:
    try:
        column_headers = sheet.row_values(n)
    except:
        break
    data_dict[column_headers[0]] = column_headers[1]
    n += 1

# Cria um txt com o dicionario
with open("dict.txt", 'w+') as values:
    values.write(str(data_dict))