import pyautogui as automation

"""
Esse projeto recebe a finalidade de melhorar o processo de finalização de chamados no software de ITSM da SoftExpert.

    Manualmente a finalização de um chamado demanda em média 13 segundos para ser executado manualmente, utilizando
esse script o tempo de execução ganha uma nova media de 5-7 segundos dependendo da máquina onde esta sendo usado o bot.

O motivo principal: Após cancelar mais de 250 chamados manualmente, percebi que sendo uma ação repetitiva pode ser melhorada
de forma em quase 50% do tempo da atividade, assim demandando menos tempo e solucionando de melhor forma tal problema.

automationITSM v.0.0.1 é somente um teste/alpha onde a lógica do script está sendo criada e implementada.

 - Para a próxima versão ao invés de adicionar manualmente o codigo no script, será possivel importar uma planilha com as informações
 já preenchidas, assim automatizando 100% a atividade.

"""


# Inicio do script
# Encontra o icone do navegador Google Chrome na barra de tarefas
automation.sleep(0)
# print(automation.position())

# A posição sendo encontrada anteriormente
# inicia o script levando o cursor do mouse em cima do icone do navegador e clicando onde esta selecionado
automation.moveTo(x=408, y=1050)
automation.sleep(0)
automation.click(x=408, y=1050)

# Seguinte continuidade é necessário o ITSM estar aberto em navegador na seguinte localidade:
# Minhas Tarefas > Acompanhamento > Workflow (Gestão de Workflow)

# O proximo script sera responsavel por encontrar a caixa de pesquisa na aba citada anteriormente
automation.sleep(0)
#print(automation.position())

# O script clica na caixa de busca de chamados mapeado anteriormente
automation.moveTo(x=156, y=402)
automation.click(x=156, y=402)

# Como teste o programa cola um codigo de exemplo: SS00447433
automation.typewrite("WCL01022162")
automation.press('enter')

# essa parte sera responsavel pelo cancelamento do chamado em teste
# Primeiramente sera mapeado a posicação da tela onde se encontra as ações
automation.sleep(0)
#print(automation.position())

# Movendo o mouse ate o botão "mais" e clicando
automation.moveTo(x=559, y=350)
automation.click(x=559, y=350)

# proximo passo sera clicar em "Alterar Situação"

automation.sleep(0)
# print(automation.position())

# Executando o click na posição mapeada anteriormente
automation.moveTo(x=619, y=392)
automation.sleep(3)
automation.click(x=619, y=392)

# apos o click deve selecionar o motivo (no caso será cancelamento)
# Mapeando a ação
automation.sleep(1.3)
# print(automation.position())

automation.moveTo(x=1052, y=356)
automation.click(x=1052, y=356)

# Agora adicionando uma descrição ao cancelamento (motivo)

#Mapeando campo "Motivo"
automation.sleep(0)
# print(automation.position())

automation.moveTo(x=490, y=456)
automation.click(x=490, y=456)

# Como teste adicionando um motivo generico

automation.typewrite("Cancelado a pedido de Fulano.")

# Salvando o cancelamento

automation.sleep(0)
# print(automation.position())

automation.moveTo(x=391, y=253)
automation.click(x=391, y=253)

# Limpando a caixa de buscas apos o cancelamentso

automation.moveTo(x=273, y=403)
automation.click(x=273, y=403)