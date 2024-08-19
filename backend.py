import subprocess
import patoolib
import os
import shutil
import tkinter as tk
import format_docx, format_xlsx, format_image, format_pdf

# Função para verificar se o arquivo é um arquivo compactado
def is_archive(file_path):
    return file_path.lower().endswith(('.zip', '.rar', '.rar4'))

# Função para extrair e imprimir arquivos
def extract_and_print_all(file_path, selected_printer, add_output_text, output_text):
    try:
        output_text.delete(1.0, tk.END)  # Limpa a saída anterior

        if not selected_printer:
            add_output_text("Impressora não selecionada.")
            return

        if is_archive(file_path):
            temp_folder = "temp_extracted"
            os.makedirs(temp_folder, exist_ok=True)

            try:
                patoolib.extract_archive(file_path, outdir=temp_folder)
                extracted_files = sorted(os.listdir(temp_folder))

                for extracted_file in extracted_files:
                    full_path = os.path.join(temp_folder, extracted_file)
                    file_extension = os.path.splitext(full_path)[1].lower()

                    if file_extension in ('.doc', '.docx'):
                        format_docx.print_docx(full_path, selected_printer, add_output_text)
                    elif file_extension in ('.xls', '.xlsx', '.csv'):
                        format_xlsx.print_xlsx(full_path, selected_printer, add_output_text)  # Correção aqui
                    elif file_extension in ('.png', '.jpg', '.jpeg', '.bmp', '.gif'):
                        format_image.print_image(full_path, selected_printer, add_output_text)
                    elif file_extension in ('.pdf'):
                        format_pdf.print_pdf(full_path, selected_printer, add_output_text)
                    else:
                        add_output_text(f"Não foi possível imprimir o arquivo {extracted_file}: extensão não suportada.")
                        
            except Exception as e:
                add_output_text(f"Erro ao extrair o arquivo: {e}")
            finally:
                shutil.rmtree(temp_folder, ignore_errors=True)
        else:
            file_extension = os.path.splitext(file_path)[1].lower()
            if file_extension in ('.doc', '.docx'):
                format_docx.print_docx(file_path, selected_printer, add_output_text)
            elif file_extension in ('.xls', '.xlsx', '.csv'):
                format_xlsx.print_xlsx(file_path, selected_printer, add_output_text)
            elif file_extension in ('.png', '.jpg', '.jpeg', '.bmp', '.gif'):
                format_image.print_image(file_path, selected_printer, add_output_text)
            elif file_extension in ('.pdf'):
                format_pdf.print_pdf(file_path, selected_printer, add_output_text)
            else:
                add_output_text(f"Não foi possível imprimir o arquivo {file_path}: extensão não suportada.")

    except Exception as e:
        add_output_text(f"Erro: {e}")

# Modifique a função para aceitar quatro argumentos
def browse_folder_and_print(folder_path, selected_printer, add_output_text):
    temp_folder = "temp_extracted"
    os.makedirs(temp_folder, exist_ok=True)

    try:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)

                if is_archive(file_path):
                    patoolib.extract_archive(file_path, outdir=temp_folder)
                else:
                    file_extension = os.path.splitext(file_path)[1].lower()

                    if file_extension in ('.doc', '.docx'):
                        format_docx.print_docx(file_path, selected_printer, add_output_text)
                    elif file_extension in ('.xls', '.xlsx', '.csv'):
                        format_xlsx.print_xlsx(file_path, selected_printer, add_output_text)  # Correção aqui
                    elif file_extension in ('.png', '.jpg', '.jpeg', '.bmp', '.gif'):
                        format_image.print_image(file_path, selected_printer, add_output_text)
                    elif file_extension in ('.pdf'):
                        format_pdf.print_pdf(file_path, selected_printer, add_output_text)
                    else:
                        add_output_text(f"Não foi possível imprimir o arquivo {file_path}: extensão não suportada.")

        for root, dirs, files in os.walk(temp_folder):
            for file in files:
                file_path = os.path.join(root, file)
                if not is_archive(file_path):
                    file_extension = os.path.splitext(file_path)[1].lower()

                    if file_extension in ('.doc', '.docx'):
                        format_docx.print_docx(file_path, selected_printer, add_output_text)
                    elif file_extension in ('.xls', '.xlsx', '.csv'):
                        format_xlsx.print_xlsx(file_path, selected_printer, add_output_text)  # Correção aqui
                    elif file_extension in ('.png', '.jpg', '.jpeg', '.bmp', '.gif'):
                        format_image.print_image(file_path, selected_printer, add_output_text)
                    elif file_extension in ('.pdf'):
                        format_pdf.print_pdf(file_path, selected_printer, add_output_text)
                    else:
                        add_output_text(f"Não foi possível imprimir o arquivo {file_path}: extensão não suportada.")

    except Exception as e:
        add_output_text(f"Erro: {e}")

    finally:
        shutil.rmtree(temp_folder, ignore_errors=True)

def show_printer_properties(selected_printer, add_output_text):
    if selected_printer:
        try:
            subprocess.run(["rundll32.exe", "printui.dll,PrintUIEntry", "/e", f"/n{selected_printer}"])
        except Exception as e:
            add_output_text(f"Erro ao abrir as configurações da impressora: {e}")
    else:
        add_output_text("⚠️ Nenhuma impressora selecionada. ⚠️\n\nPor favor selecione uma impressora antes de abrir as configurações.\n")
