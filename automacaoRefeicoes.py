import pandas as pd
import pyautogui
import time

# Carrega o arquivo Excel
df_automacao = pd.read_excel('ControleRefeições.xlsx', sheet_name='Colaborador Algodoeira')

# Garante que a coluna 'Data' seja lida como data pelo Pandas para evitar erros no strftime
df_automacao['Data'] = pd.to_datetime(df_automacao['Data'])

# Tempo para você clicar na tela do sistema antes 01/02/2026do script começar
time.sleep(5)

for index, row in df_automacao.iterrows():
    try:
        # Clica no botão de "Novo" ou "Incluir" 2
        pyautogui.click(1512, 167) 
        time.sleep(3)

        # Clica no primeiro campo (Data) e digita
        pyautogui.click(725, 287)
        pyautogui.write(row['Data'].strftime('%d/%m/%Y'), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab', presses=2) # Avança 2 campos
    
        # Cantina
        pyautogui.write(str(row['Cantina']), interval=0.2)
        time.sleep(3)
        pyautogui.press('tab') # Avança 2 campos (Correção aqui)

        # Colaborador
        pyautogui.write(str(row['Colaborador']), interval=0.2)
        time.sleep(3)
        pyautogui.press('tab')

        # Café  0dor
        pyautogui.write(str(row['Cafe']), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab')

        # Almoço Buffet (Correção do .strftime removido daqui)
        pyautogui.write(str(row['AlmocoBuffet']), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab')

        # Almoço Marmita
        pyautogui.write(str(row['AlmocoMarmita']), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab')

        # Janta
        pyautogui.write(str(row['Janta']), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab')

        # Lanche
        pyautogui.write(str(row['Lanche']), interval=0.1)
        time.sleep(3)
        pyautogui.press('tab', presses=2) # Avança 2 campos

        # Clica no botão "Salvar" ou "Gravar"
        pyautogui.press('enter') 
        time.sleep(15)

        print(f"Linha {index} processada com sucesso!")

    except Exception as e:
        print(f"Erro na linha {index}: {e}")