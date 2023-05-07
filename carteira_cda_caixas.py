import pandas as pd
import datetime
import xlsxwriter

# Definir data atual
data_atual = datetime.date.today().strftime('%d%m%Y')

# Lê o arquivo xlsx e retorna um DataFrame
carteira = pd.read_excel('carteira.xlsx', sheet_name=1)

# Filtra as linhas com "ECN" na coluna "LOJA" e valores "0" e "10" na coluna "ESTADO"
filtros = (carteira['LOJA'] == 'ECN') & (carteira['ESTADO'].isin([0, 10]))
carteira_filtrada = carteira.loc[filtros]

# Cria uma nova coluna "CAIXAS" que é o resultado da divisão da coluna "UNIDADE" pela coluna "PCB"
carteira_filtrada['CAIXAS'] = carteira_filtrada['UNIDADE'] / carteira_filtrada['PCB']

# Agrupa por "GRUPO" e soma as caixas de cada grupo
grupo_caixas = carteira_filtrada.groupby('GRUPO')['CAIXAS'].sum()

# Criar um novo DataFrame com os cabeçalhos renomeados
df = pd.DataFrame({'DEPTO': grupo_caixas.index, 'CAIXAS': grupo_caixas.values})

# Substituir valores ausentes por 0
df['CAIXAS'] = df['CAIXAS'].fillna(0)

# Calcular o total geral
total_geral = df['CAIXAS'].sum()

# Adicionar uma linha adicional com o total geral
df = df.append({'DEPTO': 'Total', 'CAIXAS': total_geral}, ignore_index=True)

# Criar um arquivo Excel e um objeto de planilha
workbook = xlsxwriter.Workbook('grupo_caixas.xlsx')
worksheet = workbook.add_worksheet()

# Definir a largura da coluna A como 2.70 cm
worksheet.set_column('A:A', 13.4)

# Definir um formato para remover as casas decimais
formato_decimal = workbook.add_format({'num_format': '0'})

# Escrever o título "CARTEIRA CDA" na primeira linha
worksheet.write('A1', f'CARTEIRA CDA: {data_atual}')

# Escrever o DataFrame no arquivo Excel com o formato de número personalizado
worksheet.write_column('A2', df['DEPTO'])
worksheet.write_column('B2', df['CAIXAS'], formato_decimal)

# Fechar o arquivo Excel
workbook.close()
