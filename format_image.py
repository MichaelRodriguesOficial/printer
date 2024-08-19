# format_image.py

import os
import tempfile
import win32print
import shutil
from PIL import Image, ImageWin
import win32ui
import win32con

def print_image(image_path, printer_name, add_output_text):
    printer_handle = None  # Inicializa a variável printer_handle

    try:
        # Verificar se o arquivo é uma imagem suportada
        if not image_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif')):
            add_output_text(f"O arquivo {image_path} não é uma imagem suportada.")
            return

        # Verificar as preferências da impressora
        printer_defaults = {"DesiredAccess": win32print.PRINTER_ACCESS_USE | win32print.PRINTER_ALL_ACCESS}
        printer_handle = win32print.OpenPrinter(printer_name, printer_defaults)

        # Criar um diretório temporário
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_file_path = os.path.join(temp_dir, os.path.basename(image_path))

            # Copiar o arquivo para o diretório temporário
            shutil.copyfile(image_path, temp_file_path)

            # Imprimir o arquivo usando um dispositivo de impressão
            image = Image.open(temp_file_path)
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)

            # Obtém as dimensões da página
            width = hdc.GetDeviceCaps(win32con.HORZRES)
            height = hdc.GetDeviceCaps(win32con.VERTRES)

            # Converte a imagem para o formato de impressão
            bitmap = ImageWin.Dib(image)
            
            # Ajusta o tamanho da imagem para a página
            image_width, image_height = image.size
            scale_x = width / image_width
            scale_y = height / image_height
            scale = min(scale_x, scale_y)

            new_width = int(image_width * scale)
            new_height = int(image_height * scale)
            
            # Ajusta o ponto de origem para centralizar a imagem na página
            x_offset = (width - new_width) // 2
            y_offset = (height - new_height) // 2

            # Inicia o trabalho de impressão
            hdc.StartDoc('Print Job')
            hdc.StartPage()
            
            # Desenha a imagem na página
            bitmap.draw(hdc.GetHandleOutput(), (x_offset, y_offset, x_offset + new_width, y_offset + new_height))

            # Finaliza o trabalho de impressão
            hdc.EndPage()
            hdc.EndDoc()
            hdc.DeleteDC()

            add_output_text(f"{os.path.basename(image_path)} | ✅\n")

    except Exception as e:
        add_output_text(f"{os.path.basename(image_path)} | {e} | ❌\n")

    finally:
        if printer_handle is not None:
            win32print.ClosePrinter(printer_handle)
