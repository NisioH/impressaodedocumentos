import pandas as pd
import pyautogui
import time
import datetime

pyautogui.FAILSAFE = False


time.sleep(2)
# Abrir programa AirMaster Posição do mouse: Point(x=347, y=746)
pyautogui.click(305, 744) 

time.sleep(60)


pyautogui.click(10, 14)
print("Clicou no menu iniciar")

time.sleep(1)
pyautogui.click(46, 44)
time.sleep(1)

pyautogui.write(str('admin'), interval=0.1)
time.sleep(1)
pyautogui.press('tab')
pyautogui.write(str('senha'), interval=0.1)
time.sleep(1)
pyautogui.click(699, 436)
time.sleep(3)

# Comerçar a tirar o relatório da Termometria
pyautogui.click(508, 11)
time.sleep(1)
pyautogui.click(38, 63)
time.sleep(1)


# Obter data atual
hoje = datetime.date.today()
ontem = hoje - datetime.timedelta(days=1)

# Se hoje for segunda-feira (weekday == 0)
if hoje.weekday() == 0:
    sabado = hoje - datetime.timedelta(days=2)
    domingo = hoje - datetime.timedelta(days=1)
    data_inicial = sabado.strftime('%d%m%Y')
    data_final = domingo.strftime('%d%m%Y')
else:
    data_inicial = ontem.strftime('%d%m%Y')
    data_final = ontem.strftime('%d%m%Y')

# Espera para garantir que o campo esteja pronto
time.sleep(2)

# Digitar data inicial
pyautogui.write(data_inicial, interval=0.1)

pyautogui.press('tab')  # ou outro comando para ir ao próximo campo

# Digitar data final
pyautogui.write(data_final, interval=0.1)

# Atualizar 
pyautogui.click(519, 174)
time.sleep(1)

#Selecionar todas as unidades
pyautogui.click(247, 354)
time.sleep(1)

#Selecionar todas datas
pyautogui.click(551, 356)
time.sleep(1)

#Selecionar todas as horas
pyautogui.click(853, 353)
time.sleep(1)

#Clicar em PDF para gerar relatório
pyautogui.click(364, 87)
time.sleep(50)

# Selecionar Área de trabalho
pyautogui.click(388, 257)
time.sleep(1)

# Selecionar pasta Termometria
pyautogui.doubleClick(671, 329)
time.sleep(1)

# Salvar
pyautogui.click(979, 519)
time.sleep(5)

pyautogui.press('enter')
time.sleep(1)

#Fechar a janela
pyautogui.click(1132, 26)
time.sleep(1)

pyautogui.click(109, 11)
time.sleep(1)

pyautogui.click(287, 73)
time.sleep(1)

pyautogui.click(568, 485)
   