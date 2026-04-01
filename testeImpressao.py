import os
import win32api
import tkinter as tk
from tkinter import filedialog, messagebox


def imprimir_arquivo(caminho_arquivo):
    try:
        print(f"Imprimindo: {caminho_arquivo}")
        win32api.ShellExecute(
            0,
            "print",
            caminho_arquivo,
            None,
            ".",
            0
        )
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao imprimir {caminho_arquivo}: {str(e)}")

def imprimir_pasta(caminho_pasta):
    if not os.path.isdir(caminho_pasta):
        messagebox.showerror("Erro", "Pasta inválida.")
        return
    
    arquivos = os.listdir(caminho_pasta)
    for arquivo in arquivos:
        caminho_arquivo = os.path.join(caminho_pasta, arquivo)
        if os.path.isfile(caminho_arquivo):
            imprimir_arquivo(caminho_arquivo)

def escolher_pasta():
    pasta = filedialog.askdirectory(title="Seelecione a pasta com os arquivos")
    if pasta:
        imprimir_pasta(pasta)

def escolher_arquivo():
    arquivo = filedialog.askopenfilename(title="Selecione o arquivo")
    if arquivo:
        imprimir_arquivo(arquivo)

root = tk.Tk()
root.title("Automação de Impressão")
root.geometry("350x250")

lbl = tk.Label(root, text="Escolha uma opção para imprimir:")
lbl.pack(pady=10)

btn_pasta = tk.Button(root, text="Imprimir todos arquivos da pasta", command=escolher_pasta, font=("Arial", 14))
btn_pasta.pack(pady=10)

btn_arquivo = tk.Button(root, text="Imprimir apenas um arquivo", command=escolher_arquivo, font=("Arial", 14))
btn_arquivo.pack(pady=10)

root.mainloop()
                             