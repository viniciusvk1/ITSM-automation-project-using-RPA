import pyautogui as automation
import time
import pandas as pd

excel_file_path = 'C:\\Users\\vinicius.almeida\\OneDrive - DISTRIBUIDORA DE MEDICAMENTOS SANTA CRUZ LTDA\\Documentos\\PycharmProjects\\RPA\\Trial Version\\First Test using Pandas and Pyautogui\\Test0.1.xlsx'
column_name = 'Teste'

data_frame = pd.read_excel(excel_file_path)

# Abrir o Bloco de Notas
automation.hotkey('winleft', 's')
automation.typewrite('bloco de notas')
time.sleep(0)
automation.press('enter')
time.sleep(2)

# Iterar sobre os códigos e escrevê-los no Bloco de Notas
for index, row in data_frame.iterrows():
    identificador = row['Teste']

    automation.write(identificador)
    automation.press('enter')  # Pressione Enter para criar uma nova linha

# Fechar o Bloco de Notas
automation.hotkey('alt', 'f4')

time.sleep(2)

automation.press('tab')

time.sleep(2)

automation.press('enter')