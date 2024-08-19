import os
import time
import tempfile
import win32print
import win32api
from PyPDF4 import PdfFileReader, PdfFileWriter

def print_pdf(file_path, printer_name, add_output_text):
    printer_handle = None  # Inicializa a variável printer_handle

    try:
        # Verificar se o arquivo é PDF
        if not file_path.lower().endswith('.pdf'):
            add_output_text(f"O arquivo {file_path} não é um arquivo PDF.")
            return

        # Abrir o arquivo PDF
        with open(file_path, 'rb') as file:
            # Verificar as preferências da impressora
            printer_defaults = {"DesiredAccess": win32print.PRINTER_ACCESS_USE | win32print.PRINTER_ALL_ACCESS}
            printer_handle = win32print.OpenPrinter(printer_name, printer_defaults)

            # Ler o PDF
            pdf = PdfFileReader(file)
            num_pages = pdf.numPages

            # Criar um diretório temporário
            with tempfile.TemporaryDirectory() as temp_dir:
                combined_pdf_path = os.path.join(temp_dir, f"{os.path.splitext(os.path.basename(file_path))[0]}.pdf")
                with open(combined_pdf_path, "wb") as temp_combined_file:
                    temp_pdf_writer = PdfFileWriter()
                    for page_num in range(num_pages):
                        # Adicionar cada página ao PdfFileWriter
                        temp_pdf_writer.addPage(pdf.getPage(page_num))
                    # Escrever todas as páginas para um único arquivo temporário
                    temp_pdf_writer.write(temp_combined_file)

                # Imprimir o arquivo PDF combinado usando win32api para chamar o comando de impressão do Windows
                print(f"Imprimindo arquivo temporário combinado: {combined_pdf_path}")

                # Verificar se a impressora está conectada e disponível
                printer_info = win32print.GetPrinter(printer_handle, 2)
                if printer_info['Status'] != 0:  # 0 indica que a impressora está pronta e sem erros
                    add_output_text(f"⚠️ A Impressora {printer_name} não está disponível ou pronta. ⚠️\n\n Verifique se a impressora está ligada ou conectada na rede.")
                    return

                # Enviar o arquivo PDF para a impressora
                result = win32api.ShellExecute(
                    0,  # Handle da janela pai, 0 significa sem janela
                    "print",  # Comando a ser executado
                    combined_pdf_path,  # Arquivo a ser impresso
                    f'/d:"{printer_name}"',  # Argumentos para o comando, especificando a impressora
                    ".",  # Diretório de trabalho
                    0  # Modo de exibição, 0 significa oculto
                )

                # Aguardar para garantir que o processo de impressão seja iniciado
                if result <= 32:
                    #add_output_text(f"Erro ao tentar imprimir. Código de erro: {result}")
                    add_output_text(f"{os.path.basename(file_path)} | Páginas: {num_pages} | {result} | ❌\n")
                else:
                    # Espera de alguns segundos para garantir que a impressão seja processada
                    time.sleep(3)

            add_output_text(f"{os.path.basename(file_path)} | Páginas: {num_pages} | ✅\n")

    #except Exception as e:
    #    add_output_text(f"{os.path.basename(file_path)} | Páginas: {num_pages} | {e} | ❌\n")

    finally:
        if printer_handle is not None:
            win32print.ClosePrinter(printer_handle)
