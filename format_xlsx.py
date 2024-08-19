# format_xlsx.py

import os
import tempfile
import win32print
import shutil
from comtypes import client

def print_xlsx(file_path, printer_name, add_output_text):
    printer_handle = None  # Inicializa a variável printer_handle

    try:
        # Verificar se o arquivo é XLSX ou XLS
        if not file_path.lower().endswith(('.xlsx', '.xls', '.csv')):
            add_output_text(f"O arquivo {file_path} não é um arquivo XLSX ou XLS.")
            return

        # Verificar as preferências da impressora
        printer_defaults = {"DesiredAccess": win32print.PRINTER_ACCESS_USE | win32print.PRINTER_ALL_ACCESS}
        printer_handle = win32print.OpenPrinter(printer_name, printer_defaults)

        # Criar um diretório temporário
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_file_path = os.path.join(temp_dir, os.path.basename(file_path))

            # Copiar o arquivo para o diretório temporário
            shutil.copyfile(file_path, temp_file_path)

            # Imprimir o arquivo usando o Excel
            excel = client.CreateObject('Excel.Application')
            excel.Visible = False  # Não mostrar o Excel
            workbook = excel.Workbooks.Open(temp_file_path)
            workbook.PrintOut()
            workbook.Close()
            excel.Quit()

        add_output_text(f"{os.path.basename(file_path)} | ✅\n")

    except Exception as e:
        add_output_text(f"{os.path.basename(file_path)} | {e} | ❌\n")

    finally:
        if printer_handle is not None:
            win32print.ClosePrinter(printer_handle)
