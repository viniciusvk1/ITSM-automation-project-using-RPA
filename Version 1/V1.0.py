import pyautogui as automation
import pandas as pd
import time

excel_file_path = 'C:\\Users\\vinicius.almeida\\OneDrive - DISTRIBUIDORA DE MEDICAMENTOS SANTA CRUZ LTDA\\Documentos\\PycharmProjects\\RPA\\Trial Version\\First Test using Pandas and Pyautogui\\Test0.1.xlsx'
column_name = 'Teste'

data_frame = pd.read_excel(excel_file_path)

automation.moveTo(x=408, y=1050)
time.sleep(0)
automation.click(x=408, y=1050)

automation.moveTo(x=156, y=402)
automation.click(x=156, y=402)
