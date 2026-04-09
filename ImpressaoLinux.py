import os
import subprocess
import flet as ft


def main(page: ft.Page):
    # Configurações da janela
    page.title = "Automação de Impressão - Linux"
    page.window_width = 450
    page.window_height = 300
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.theme_mode = ft.ThemeMode.LIGHT  # ou DARK, como preferir

    # --- Configurações de Impressão ---
    EXTENSOES_PERMITIDAS = {".pdf", ".jpg", ".jpeg", ".png", ".txt", ".odt", ".docx"}

    # --- Funções de Notificação Visual ---
    def mostrar_erro(mensagem):
        snack = ft.SnackBar(ft.Text(mensagem, color=ft.colors.WHITE), bgcolor=ft.colors.ERROR)
        page.overlay.append(snack)
        snack.open = True
        page.update()

    def mostrar_sucesso(mensagem):
        snack = ft.SnackBar(ft.Text(mensagem, color=ft.colors.WHITE), bgcolor=ft.colors.GREEN_700)
        page.overlay.append(snack)
        snack.open = True
        page.update()

    # --- Lógica de Impressão ---
    def imprimir_arquivo(caminho_arquivo):
        try:
            print(f"Imprimindo: {caminho_arquivo}")
            # Envia o arquivo para o CUPS (impressora padrão do Linux)
            subprocess.run(['lpr', caminho_arquivo], check=True)
            mostrar_sucesso(f"Enviado para impressão:\n{os.path.basename(caminho_arquivo)}")
            return True

        except subprocess.CalledProcessError as e:
            mostrar_erro(f"Falha ao imprimir {os.path.basename(caminho_arquivo)}.")
        except FileNotFoundError:
            mostrar_erro("Comando 'lpr' não encontrado. Verifique se o CUPS está instalado.")
        except Exception as e:
            mostrar_erro(f"Erro inesperado: {str(e)}")
        return False

    def imprimir_pasta(caminho_pasta):
        if not os.path.isdir(caminho_pasta):
            mostrar_erro("Pasta inválida.")
            return

        # Listagem ordenada alfabeticamente
        arquivos = sorted(os.listdir(caminho_pasta))
        arquivos_impressos = 0

        for arquivo in arquivos:
            caminho_arquivo = os.path.join(caminho_pasta, arquivo)
            extensao = os.path.splitext(arquivo)[1].lower()

            if os.path.isfile(caminho_arquivo) and extensao in EXTENSOES_PERMITIDAS:
                if imprimir_arquivo(caminho_arquivo):
                    arquivos_impressos += 1

        if arquivos_impressos > 0:
            mostrar_sucesso(f"{arquivos_impressos} arquivo(s) enviados para a fila!")
        else:
            mostrar_erro("Nenhum arquivo válido (PDF, Imagem, TXT) encontrado na pasta.")

    # --- Tratamento das Janelas de Seleção ---
    def ao_escolher_arquivo(e: ft.FilePickerResultEvent):
        if e.files and len(e.files) > 0:
            arquivo = e.files[0]
            extensao = os.path.splitext(arquivo.name)[1].lower()
            if extensao in EXTENSOES_PERMITIDAS:
                imprimir_arquivo(arquivo.path)
            else:
                mostrar_erro(f"Extensão {extensao} não suportada.")

    def ao_escolher_pasta(e: ft.FilePickerResultEvent):
        if e.path:
            imprimir_pasta(e.path)

    # Configuração dos Seletores (FilePickers)
    seletor_arquivo = ft.FilePicker(on_result=ao_escolher_arquivo)
    seletor_pasta = ft.FilePicker(on_result=ao_escolher_pasta)

    # No Flet, os FilePickers precisam ser adicionados ao "overlay" da página
    page.overlay.extend([seletor_arquivo, seletor_pasta])

    # --- Construção da Interface ---
    page.add(
        ft.Icon(name=ft.icons.PRINT, size=50, color=ft.colors.BLUE),
        ft.Text("Escolha uma opção para imprimir:", size=18, weight=ft.FontWeight.BOLD),
        ft.Divider(height=20, color="transparent"),  # Espaçamento

        ft.ElevatedButton(
            text="Imprimir todos arquivos da pasta",
            icon=ft.icons.FOLDER_OPEN,
            width=350,
            height=50,
            on_click=lambda _: seletor_pasta.get_directory_path(dialog_title="Selecione a pasta com os arquivos")
        ),

        ft.ElevatedButton(
            text="Imprimir apenas um arquivo",
            icon=ft.icons.INSERT_DRIVE_FILE,
            width=350,
            height=50,
            on_click=lambda _: seletor_arquivo.pick_files(dialog_title="Selecione o arquivo")
        )
    )


# Executa o aplicativo
ft.run(main)
