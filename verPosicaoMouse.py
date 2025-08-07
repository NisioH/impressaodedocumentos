import pyautogui
import time


time.sleep(5)  # Te dá 5 segundos para posicionar o cursor19/c02
print("Posição do mouse:", pyautogui.position())
import keyboard

keyboard.write("Descrição com ç, á, ê, ó, ã, etc.", delay=0.1)


pyautogui.write("çãõáéíóú âêîôû ÀÈÌÒÙ")



 