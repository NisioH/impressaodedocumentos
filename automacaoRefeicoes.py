import pandas as pd
import pyautogui
import time

df_automacao = pd.read_excel('ControleRefeições.xlsx', sheet_name='Safrista Algodoeira')

df_automacao['Data'] = pd.to_datetime(df_automacao['Data'])

time.sleep(5)

for index, row in df_automacao.iterrows():
    try:
        pyautogui.click(1512, 167) 
        time.sleep(3)

        pyautogui.click(725, 287)
        pyautogui.write(row['Data'].strftime('%d/%m/%Y'), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab', presses=2) 

        pyautogui.write(str(row['Cantina']), interval=0.2)
        time.sleep(3)
        pyautogui.press('tab') 

        pyautogui.write(str(row['Colaborador']), interval=0.2)
        time.sleep(3)
        pyautogui.press('tab')

        pyautogui.write(str(row['Cafe']), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab')

        pyautogui.write(str(row['AlmocoBuffet']), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab')

        pyautogui.write(str(row['AlmocoMarmita']), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab')

        pyautogui.write(str(row['Janta']), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab')

        pyautogui.write(str(row['Lanche']), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab', presses=2)

        pyautogui.press('enter') 
        time.sleep(15)

        print(f"Linha {index} processada com sucesso!")

    except Exception as e:
        print(f"Erro na linha {index}: {e}")