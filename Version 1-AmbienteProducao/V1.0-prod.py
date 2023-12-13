import pyautogui as automation
import pandas as pd
import time

excel_file_path = 'C:\\Users\\vinicius.almeida\\OneDrive - DISTRIBUIDORA DE MEDICAMENTOS SANTA CRUZ LTDA\\Documentos\\PycharmProjects\\ITSM-automation-project-using-RPA\\Version 1-AmbienteProducao\\ExcelTestSSO.xlsx'
column_name = 'Identificador'

data_frame = pd.read_excel(excel_file_path)

#Aguarda 5 segundos para iniciar o script
time.sleep(5)

#Seleciona o campo de busca para colar o identificador
automation.click(x=221, y=334)

#Cola os identificadores no campo de busca
for index, row in data_frame.iterrows():
    identificador = str(row['Identificador'])

    automation.write(identificador)
    automation.press('enter')

    #Move para o botão "Mais"
    time.sleep(1)
    automation.click(x=667, y=282)

    #Move para o botão "Alterar Situação"
    time.sleep(2)
    automation.click(x=712, y=341)

    #Aguarda a tela de justificativa abrir
    time.sleep(2)

    #Seleciona a opção "Cancelar"
    automation.click(x=1094, y=379)

    #Clica no campo de justificativa e digitar o motivo
    automation.click(x=756, y=627)
    automation.typewrite("Chamado cancelado por obsolescência!")

    #Clica no botão "Salvar e Sair"
    automation.click(x=401, y=269)

    #Aguarda um tempo para executar a proxima ação
    time.sleep(1.5)

    #Limpa a caixa de busca para colar o proximo identificador
    automation.click(x=326, y=348)

    #Clica novamente na caixa de busca para colar o proximo código
    automation.click(x=136, y=349)

    time.sleep(3)