import pandas as pd
import pyautogui
import time


df_automacao = pd.read_excel(
    r'C:\Users\fazin\OneDrive\Documents\2023\SOLICITAÇÕES - 2025.xlsx',
    sheet_name='Fevereiro'
)


time.sleep(5)

for index, row in df_automacao.iterrows():
    try:
        
        
        pyautogui.click(1198, 167) 
        time.sleep(1)

        pyautogui.click(338, 200)
        pyautogui.write(str(row['Nº SOLICITAÇÃO']), interval=0.1)
        time.sleep(1)
        pyautogui.press('tab')

        pyautogui.write(str(row['DESCRIÇAO DA SOLICITAÇÃO / PEDIDO SERVIÇO, PEÇAS E PRODUTOS']), interval=0.1)
        time.sleep(1)
        pyautogui.press('tab')

        pyautogui.write(str(row['SOLICITADO']), interval=0.1)
        time.sleep(1)
        pyautogui.press('tab')

        pyautogui.write(str(row['SAFRA']), interval=0.1)
        time.sleep(1)
        pyautogui.press('tab')

        pyautogui.write(str(row['C. CUSTO']), interval=0.1)
        time.sleep(1)
        pyautogui.press('tab')

        pyautogui.write(str(row['STATUS']), interval=0.1)
        time.sleep(1)
        pyautogui.press('tab')

        pyautogui.scroll(-1000)
        time.sleep(1)

        if pd.notna(row['DATA']):
            data_formatada = row['DATA'].strftime('%d/%m/%Y')
            pyautogui.write(data_formatada, interval=0.3)
        else:
            pyautogui.write('', interval=0.3)
        time.sleep(1)

        pyautogui.click(341, 494)
        if pd.notna(row['DATA RECEBIDO']):
            recebido_formatado = row['DATA RECEBIDO'].strftime('%d/%m/%Y')
            pyautogui.write(recebido_formatado, interval=0.1)
        else:
            pyautogui.write('', interval=0.1)
        time.sleep(1)

        pyautogui.click(357, 572)
        pyautogui.write(str(row['Fornecedor']), interval=0.1)
        time.sleep(1)
        pyautogui.press('tab')

        pyautogui.write(str(row['Nota fiscal']), interval=0.1)
        time.sleep(1)
        pyautogui.press('tab')

        pyautogui.click(353, 693)
        time.sleep(1)
    
        pyautogui.click(1229, 167)
        time.sleep(3)

        print(f"Linha {index} processada com sucesso!")

    except Exception as e:
        print(f"Erro na linha {index}: {e}")





   