import os
import subprocess
import flet as ft
import tempfile
import shutil


def main(page: ft.Page):
    # Configurações da janela (Atualizado para a API nova do Flet)
    page.title = "Automação de Impressão - Linux"
    page.window.width = 450
    page.window.height = 300
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.theme_mode = ft.ThemeMode.LIGHT  # ou DARK, como preferir

    # --- Configurações de Impressão ---
    EXTENSOES_PERMITIDAS = {".pdf", ".jpg", ".jpeg", ".png", ".txt", ".odt", ".odf", ".docx"}

    # --- Funções de Notificação Visual ---
    def mostrar_erro(mensagem):
        snack = ft.SnackBar(ft.Text(mensagem, color=ft.Colors.WHITE), bgcolor=ft.Colors.ERROR)
        page.overlay.append(snack)
        snack.open = True
        page.update()

    def mostrar_sucesso(mensagem):
        snack = ft.SnackBar(ft.Text(mensagem, color=ft.Colors.WHITE), bgcolor=ft.Colors.GREEN_700)
        page.overlay.append(snack)
        snack.open = True
        page.update()

    # --- Lógica de Impressão ---
    def imprimir_arquivo(caminho_arquivo):
        extensao = os.path.splitext(caminho_arquivo)[1].lower()

        try:
            print(f"-> Processando arquivo: {caminho_arquivo}")

            # Se for documento do Office, convertemos para PDF primeiro
            if extensao in [".docx", ".doc", ".odt"]:
                print("-> Arquivo Office detectado. Convertendo para PDF invisivelmente...")

                pasta_origem = os.path.dirname(caminho_arquivo)

                # Pede ao LibreOffice para gerar um PDF na mesma pasta do arquivo original
                subprocess.run([
                    'libreoffice',
                    '--headless',
                    '--convert-to', 'pdf',
                    caminho_arquivo,
                    '--outdir', pasta_origem
                ], check=True)

                # Monta o caminho do novo arquivo PDF gerado
                nome_base = os.path.splitext(os.path.basename(caminho_arquivo))[0]
                caminho_pdf = os.path.join(pasta_origem, f"{nome_base}.pdf")

                try:
                    print(f"-> Conversão concluída! Imprimindo o gerado: {caminho_pdf}")
                    subprocess.run(['lpr', caminho_pdf], check=True)
                finally:
                    # Deleta o PDF fantasma para não sujar a máquina do usuário
                    if os.path.exists(caminho_pdf):
                        os.remove(caminho_pdf)
                        print("-> PDF temporário apagado com sucesso.")

            # Para PDFs, imagens e TXT originais, o lpr funciona nativamente
            else:
                print("-> Arquivo simples detectado. Usando lpr...")
                subprocess.run(['lpr', caminho_arquivo], check=True)

            mostrar_sucesso(f"Enviado para impressão:\n{os.path.basename(caminho_arquivo)}")
            return True

        except subprocess.CalledProcessError as e:
            mostrar_erro(f"Falha ao imprimir {os.path.basename(caminho_arquivo)}.")
            print(f"Erro no subprocesso: {e}")
        except FileNotFoundError:
            if extensao in [".docx", ".doc", ".odt"]:
                mostrar_erro("LibreOffice não encontrado. Instale-o no terminal.")
            else:
                mostrar_erro("Comando 'lpr' não encontrado.")
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

    # --- Tratamento das Janelas de Seleção (ATUALIZADO) ---
    async def ao_escolher_arquivo(e):
        # Chama o FilePicker de forma assíncrona
        arquivos = await ft.FilePicker().pick_files(dialog_title="Selecione o arquivo")

        if arquivos and len(arquivos) > 0:
            arquivo = arquivos[0]
            extensao = os.path.splitext(arquivo.name)[1].lower()
            if extensao in EXTENSOES_PERMITIDAS:
                imprimir_arquivo(arquivo.path)
            else:
                mostrar_erro(f"Extensão {extensao} não suportada.")

    async def ao_escolher_pasta(e):
        # Chama o FilePicker de forma assíncrona
        pasta = await ft.FilePicker().get_directory_path(dialog_title="Selecione a pasta com os arquivos")

        if pasta:
            imprimir_pasta(pasta)

    # --- Construção da Interface ---
    page.add(
        ft.Icon(ft.Icons.PRINT, size=50, color=ft.Colors.BLUE),
        ft.Text("Escolha uma opção para imprimir:", size=18, weight=ft.FontWeight.BOLD),
        ft.Divider(height=20, color="transparent"),  # Espaçamento

        ft.ElevatedButton(
            "Imprimir todos arquivos da pasta",
            icon=ft.Icons.FOLDER_OPEN,
            width=350,
            height=50,
            on_click=ao_escolher_pasta  # Passa a função async diretamente
        ),

        ft.ElevatedButton(
            "Imprimir apenas um arquivo",
            icon=ft.Icons.INSERT_DRIVE_FILE,
            width=350,
            height=50,
            on_click=ao_escolher_arquivo  # Passa a função async diretamente
        )
    )


# Executa o aplicativo
ft.app(target=main)