import pandas as pd

# Carregar o arquivo Excel
excel_file_path = 'Caminho da planilha que deseja inserir'
column_name = 'Identificador'

data_frame = pd.read_excel(excel_file_path)

for index, row in data_frame.iterrows():
    cell_value = row[column_name]
    print(cell_value)