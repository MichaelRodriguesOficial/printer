# App de Impressão CAV SUL

Este é um aplicativo de interface gráfica (GUI) desenvolvido em Python que permite aos usuários selecionar e imprimir arquivos de diferentes formatos (como .docx, .xlsx, .pdf, .png, entre outros) em uma impressora escolhida. O aplicativo também suporta a extração de arquivos compactados (.zip, .rar) e a impressão dos arquivos contidos neles.

## Funcionalidades

- **Seleção de Arquivo**: O usuário pode selecionar um único arquivo para impressão.
- **Seleção de Pasta**: O usuário pode selecionar uma pasta, e todos os arquivos suportados dentro dela serão impressos.
- **Suporte a Arquivos Compactados**: O aplicativo pode extrair e imprimir arquivos dentro de arquivos .zip e .rar.
- **Configurações de Impressora**: O usuário pode escolher a impressora desejada e acessar as configurações da impressora.
- **Temas**: O aplicativo oferece suporte a temas claro e escuro.

## Tecnologias Utilizadas

- **Python 3.x**
- **Tkinter**: Biblioteca padrão do Python para criação de GUIs.
- **PIL (Pillow)**: Utilizado para manipulação de imagens.
- **patoolib**: Utilizado para extração de arquivos compactados.
- **win32print**: Utilizado para gerenciar impressoras no ambiente Windows.

## Estrutura do Projeto

- `frontend.py`: Código responsável pela interface gráfica e interação do usuário.
- `backend.py`: Código responsável pela lógica de extração e impressão dos arquivos.
- `format_docx.py`, `format_xlsx.py`, `format_image.py`, `format_pdf.py`: Módulos para lidar com a formatação e impressão de diferentes tipos de arquivos.

## Como Executar

1. **Instale as dependências**:
   - Certifique-se de ter o Python 3.x instalado.
   - Instale as bibliotecas necessárias com o comando:
     ```bash
     pip install pillow patool pywin32
     ```

2. **Execute o aplicativo**:
   - Navegue até o diretório do projeto e execute o arquivo `frontend.py`:
     ```bash
     python frontend.py
     ```

## Screenshots

Aqui estão algumas capturas de tela do aplicativo em ação:

### Tela Principal

![Tela Principal Claro](https://github.com/MichaelRodriguesOficial/printer/blob/main/imagens/Claro.png?raw=true)
![Tela Principal Escuro](https://github.com/MichaelRodriguesOficial/printer/blob/main/imagens/Escuro.png?raw=true)

### Avisos

![Avisos](https://github.com/MichaelRodriguesOficial/printer/blob/main/imagens/Avisos.png?raw=true)

### Imprimindo

![Imprimindo](https://github.com/MichaelRodriguesOficial/printer/blob/main/imagens/Imprimindo.png?raw=true)

## Contribuições

Contribuições são bem-vindas! Se você encontrar algum problema ou tiver alguma sugestão, sinta-se à vontade para abrir uma issue ou enviar um pull request.

## Transformando o código em .exe com o PyInstaller

     ```bash
     pip install pyinstaller
     ```
     ```bash
     pyinstaller --noconfirm --onefile --windowed --icon "C:\Github\Printer\img\favicon.ico" --add-data "C:\Github\Printer\img\sol.png;images" --add-data "C:\Github\Printer\img\lua.png;images" --hidden-import pkg_resources --hidden-import pkg_resources.extern "C:/Github/Printer/frontend.py"
     ```
     
## Licença

Este projeto é licenciado sob a [MIT License](LICENSE).

## Autor

Desenvolvido por Michael Rodrigues.

