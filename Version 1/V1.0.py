import pyautogui as automation
import pandas as pd
import time

excel_file_path = 'C:\\Users\\vinicius.almeida\\OneDrive - DISTRIBUIDORA DE MEDICAMENTOS SANTA CRUZ LTDA\\Documentos\\PycharmProjects\\RPA\\Version 1\\ExcelTestSSO.xlsx'
column_name = 'Identificador'

data_frame = pd.read_excel(excel_file_path)

# Abre o navegador e clica na caixa de busca dos chamados
automation.moveTo(x=408, y=1050)
automation.click(x=408, y=1050)
automation.moveTo(x=156, y=402)
time.sleep(0.4)
automation.click(x=156, y=402)

for index, row in data_frame.iterrows():
    identificador = str(row['Identificador'])

    automation.write(identificador)
    automation.press('enter')

    automation.moveTo(x=580, y=354)
    automation.click(x=580, y=354)

    automation.moveTo(x=605, y=392)
    time.sleep(1.8)
    automation.click(x=605, y=392)

    automation.moveTo(x=1054, y=356)
    time.sleep(0.5)
    automation.click(x=1054, y=356)

    automation.moveTo(x=493, y=471)
    automation.click(x=493, y=471)
    automation.typewrite("Chamado cancelado por X motivo a pedido de Fulano")

    automation.moveTo(x=394, y=250)
    automation.click(x=394, y=250)

    automation.moveTo(x=273, y=404)
    automation.click(x=273, y=404)

    time.sleep(0.7)
    automation.moveTo(x=156, y=402)
    automation.click(x=156, y=402)