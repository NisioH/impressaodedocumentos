import pandas as pd
import pyautogui
import time
import datetime

pyautogui.FAILSAFE = False


time.sleep(2)
# Abrir programa AirMaster Posição do mouse: Point(x=347, y=746)
pyautogui.click(697, 1052) 

time.sleep(60)


pyautogui.click(17, 17)
print("Clicou no menu iniciar")

time.sleep(1)
pyautogui.click(45, 40)
time.sleep(1)

pyautogui.write(str('admin'), interval=0.1)
time.sleep(1)
pyautogui.press('tab')
pyautogui.write(str('senha'), interval=0.1)
time.sleep(1)
pyautogui.click(961, 590)
time.sleep(3)

# Comerçar a tirar o relatório da Termometria
pyautogui.click(500, 13)
time.sleep(1)
pyautogui.click(41, 65)
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
pyautogui.click(798, 331)
time.sleep(1)

#Selecionar todas as unidades
pyautogui.click(524, 510)
time.sleep(1)

#Selecionar todas datas
pyautogui.click(827, 510)
time.sleep(1)

#Selecionar todas as horas
pyautogui.click(1130, 510)
time.sleep(1)

#Clicar em PDF para gerar relatório
pyautogui.click(664, 239)
time.sleep(50)

# Selecionar Área de trabalho
pyautogui.click(653, 405)
time.sleep(2)

# Selecionar pasta Termometria
""" pyautogui.doubleClick(645, 345)
time.sleep(1) """

# Salvar
pyautogui.click(1264, 669)
time.sleep(5)

pyautogui.press('enter')
time.sleep(1)

#Fechar a janela
pyautogui.click(1409, 185)
time.sleep(1)

pyautogui.click(115, 16)
time.sleep(1)

""" pyautogui.click(287, 73)
time.sleep(1)

pyautogui.click(568, 485) """
   