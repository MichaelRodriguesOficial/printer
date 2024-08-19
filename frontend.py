import tkinter as tk
from tkinter import filedialog, ttk
from PIL import Image, ImageTk
import threading
import backend
import win32print

# Funão para carregar ícones como imagens
def load_icon(filename):
    image = Image.open(filename)
    image = image.resize((24, 24), Image.BICUBIC)
    return ImageTk.PhotoImage(image)

# Função para abrir uma janela de diálogo de seleção de arquivo
def browse_file():
    selected_printer = printer_combobox.get()
    if not selected_printer:
        add_output_text("⚠️ Nenhuma impressora selecionada. ⚠️\n\nPor favor selecione uma impressora antes de selecionar o arquivo.\n")
        return
    
    file_path = filedialog.askopenfilename()
    if file_path:
        # Executar a extração e impressão em uma thread separada
        threading.Thread(target=lambda: backend.extract_and_print_all(file_path, selected_printer, add_output_text, output_text)).start()

def browse_folder_and_print():
    selected_printer = printer_combobox.get()
    if not selected_printer:
        add_output_text("⚠️ Nenhuma impressora selecionada. ⚠️\n\nPor favor selecione uma impressora antes de selecionar a pasta.\n")
        return

    folder_path = filedialog.askdirectory()
    if folder_path:
        threading.Thread(target=lambda: backend.browse_folder_and_print(folder_path, selected_printer, add_output_text)).start()


def add_output_text(text):
    output_text.config(state=tk.NORMAL)
    output_text.insert(tk.END, text + "\n")
    output_text.see(tk.END)
    output_text.config(state=tk.DISABLED)

def show_printer_properties():
    selected_printer = printer_combobox.get()
    backend.show_printer_properties(selected_printer, add_output_text)

# Dicionário para armazenar as cores globais
colors = {
    "bg_color": "#F5F5F5",
    "bt_color": "#4CAF50",
    "text_color": "#333333",
    "hover_color": "#45a049"
}

# Função para alternar para o tema claro
def light_theme():
    colors.update({
        "bg_color": "#F5F5F5",
        "bt_color": "#4CAF50",
        "text_color": "#333333"
    })
    update_theme()

# Função para alternar para o tema escuro
def dark_theme():
    colors.update({
        "bg_color": "#333333",
        "bt_color": "#4CAF50",
        "text_color": "#FFFFFF"
    })
    update_theme()

# Função para atualizar o tema
def update_theme():
    app.configure(bg=colors["bg_color"])
    left_frame.configure(style="TFrame")
    right_frame.configure(style="TFrame")
    properties_button.configure(style="TButton")
    browse_button.configure(style="TButton")
    print_folder_button.configure(style="TButton")
    image_label.configure(background=colors["bg_color"])
    printer_label.configure(style="TLabel")
    output_text.configure(bg=colors["bg_color"], fg=colors["text_color"])
    update_styles()  # Atualiza os estilos do ttk

# Função para configurar o estilo
def update_styles():
    style = ttk.Style()
    style.theme_use("clam")

    style.configure("TFrame", background=colors["bg_color"])
    style.configure("TLabel", background=colors["bg_color"], foreground=colors["text_color"], font=("Roboto", 10))
    style.configure("TButton", background=colors["bt_color"], foreground="white", font=("Roboto", 10, "bold"), borderwidth=0)
    style.map("TButton", background=[("active", colors["hover_color"])], foreground=[("active", "white")])
    style.configure("TCombobox", background=colors["bg_color"], foreground=colors["text_color"], font=("Roboto", 10))

# Cria a janela
app = tk.Tk()
app.title("App de Impressão CAV SUL")
app.geometry("670x410")
app.resizable(False, False)  # Não permite redimensionamento
app.iconbitmap('img/favicon.ico')

# Cria um Frame para organizar os elementos à esquerda
left_frame = ttk.Frame(app)
left_frame.grid(row=0, column=0, padx=20, pady=20, sticky="n")

# Criar um Frame para os botões de tema
theme_buttons_frame = ttk.Frame(left_frame)
theme_buttons_frame.grid(row=0, column=0, pady=5, padx=5, sticky="ew")

# Carrega os ícones
sun_icon = load_icon("img/sol.png")
moon_icon = load_icon("img/lua.png")

# Criar os botões para alternar entre os temas
light_theme_button = ttk.Button(theme_buttons_frame, image=sun_icon, command=light_theme, style="TButton")
light_theme_button.grid(row=0, column=0, padx=(0, 60), sticky="e")

dark_theme_button = ttk.Button(theme_buttons_frame, image=moon_icon, command=dark_theme, style="TButton")
dark_theme_button.grid(row=0, column=1, padx=(65, 0), sticky="w")

# Cria um Frame para organizar a imagem no left_frame
image_frame = ttk.Frame(left_frame)
image_frame.grid(row=1, column=0, pady=15)

# Adiciona um ícone
image_label = ttk.Label(image_frame, text="🖨", font=("Arial", 64), background=colors["bg_color"])
image_label.grid(row=0, sticky="ew")

# Obtém a lista de impressoras disponíveis no sistema
available_printers = win32print.EnumPrinters(2)
printer_names = [printer[2] for printer in available_printers]

printer_label = ttk.Label(left_frame, text="Selecione uma Impressora", font=("Roboto", 12, "bold"))
printer_label.grid(row=2, column=0, sticky="ew", pady=10)

# Cria uma lista suspensa (combobox) para selecionar a impressora
printer_combobox = ttk.Combobox(left_frame, values=printer_names, state="readonly", font=("Roboto", 10))
printer_combobox.grid(row=3, column=0, pady=5, sticky="ew")

# Cria um botão para abrir as configurações da impressora
properties_button = ttk.Button(left_frame, text="Configuração de Impressão", command=show_printer_properties)
properties_button.grid(row=4, column=0, pady=7, sticky="ew")

# Cria um botão para selecionar o arquivo
browse_button = ttk.Button(left_frame, text="Selecionar Arquivo", command=browse_file)
browse_button.grid(row=5, column=0, pady=7, sticky="ew")

# Cria um botão para selecionar a pasta e imprimir arquivos
print_folder_button = ttk.Button(left_frame, text="Selecionar Pasta", command=browse_folder_and_print)
print_folder_button.grid(row=6, column=0, pady=7, sticky="ew")

# Cria um Frame para organizar o visor de saída à direita
right_frame = ttk.Frame(app)
right_frame.grid(row=0, column=1, pady=20, sticky="n")

# Cria um subframe para o visor de saída com o gerenciador de geometria grid
output_frame = ttk.Frame(right_frame)
output_frame.pack()

# Cria uma barra de rolagem vertical
scrollbar = ttk.Scrollbar(output_frame, orient="vertical")

# Cria um widget Text para exibir o conteúdo com a barra de rolagem
output_text = tk.Text(output_frame, wrap=tk.WORD, width=55, height=23, yscrollcommand=scrollbar.set,
                      bg=colors["bg_color"], fg=colors["text_color"], font=("Roboto", 10),
                      borderwidth=2, relief="groove", state="disabled")
output_text.grid(row=0, column=0)

# Vincula a barra de rolagem ao widget Text
scrollbar.config(command=output_text.yview)
scrollbar.grid(row=0, column=1, sticky="ns")

light_theme()  # Inicializa com tema claro

app.mainloop()
