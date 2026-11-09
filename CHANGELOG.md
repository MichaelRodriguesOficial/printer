# Changelog

## [2.1.0] — 2026-11

Impressão mais rápida, principalmente em lotes de PDF.

### Desempenho
- **PDF vetorial:** o PDF é desenhado direto na impressora (texto e vetores), como fazem o Foxit e o
  Chrome, em vez de virar uma imagem de 300 dpi por página. Para 20 páginas de texto, os dados
  enviados à impressora caíram de 1 GB para 16 MB num driver PostScript clássico e de 56 MB para
  8 MB no driver PCL6 do Windows. No PostScript, o tempo total ficou 4x menor.
- **Conexão com a impressora reaproveitada** durante o lote, em vez de reaberta a cada arquivo.
- Fotos grandes são reduzidas ao tamanho impresso antes do envio, e imagens em tons de cinza
  seguem em 8 bits.
- O resumo do lote mostra o tempo total ("Concluído em 12 s"), e o log registra o tempo de cada arquivo.

### Novidades
- Chave **"PDF rápido (vetorial)"** na barra lateral, ligada por padrão. Se alguma impressora
  imprimir PDFs com defeito, basta desligar para voltar ao modo imagem.
- `tools/benchmark_print.py` compara os modos numa impressora. O CI roda esse benchmark
  em seis drivers do Windows a cada push.

## [2.0.0] — 2026-11

Reescrita completa: nova interface, impressão confiável na impressora escolhida e testes
automatizados no Windows.

### Novidades
- Interface nova em CustomTkinter, com tema claro, escuro ou do sistema e identidade azul da CAV SUL.
- **Fila de impressão**: os arquivos aparecem numa lista antes de imprimir, com status por item
  (Na fila, Imprimindo, Impresso, Erro, Ignorado, Cancelado) e número de páginas.
- **Arrastar e soltar** arquivos, pastas e compactados na janela.
- **Cópias** (1 a 99), botão **Cancelar** e reimpressão só dos itens com erro.
- A impressora padrão do Windows já vem selecionada, e a lista pode ser atualizada sem reabrir o app.
- O app mostra o status da impressora (offline, sem papel, porta aberta...) e pede confirmação antes de imprimir.
- Lembra a última impressora, o número de cópias e o tema.
- Novos formatos: TIFF com várias páginas, WebP, RTF, ODT, TXT, XLSM, XLSB, ODS, 7z e TAR.
- PDFs impressos com o PDFium, sem depender do leitor de PDF instalado. Campos de formulário preenchidos também saem.
- Sem Microsoft Office, Word e Excel são impressos pelo LibreOffice.
- Log em `%LOCALAPPDATA%\PrinterCAV\logs`, acessível pelo link "Abrir logs".

### Correções
- Word, Excel e PDF agora saem **na impressora escolhida** (antes saíam na padrão do Windows).
- O Word não cancela mais a impressão ao fechar (`Background=False`), e diálogos invisíveis não travam mais o app.
- Macros de documentos do Office ficam sempre desabilitadas durante a impressão.
- Não exige mais permissão de administrador na impressora (`PRINTER_ACCESS_USE`).
- Um PDF com erro não interrompe mais o lote inteiro, e o status ✅/❌ ficou correto.
- `extensão in ('.pdf')` comparava texto em vez de tupla: arquivos sem extensão iam para o fluxo de PDF.
- Zips com subpastas e zips aninhados funcionam, e dois zips com arquivos de mesmo nome não se sobrescrevem mais.
- Nomes acentuados de zips criados pelo Windows Explorer são lidos corretamente.
- Extração protegida contra *path traversal* e "zip bomb".
- `.rar` e `.7z` encontram o 7-Zip instalado mesmo fora do PATH.
- PNG com transparência sai com fundo branco (antes saía preto), e fotos de celular saem na rotação correta (EXIF).
- A interface não é mais atualizada a partir de outra thread, o que causava travamentos aleatórios.
- O executável encontra os ícones e recursos de qualquer pasta (`sys._MEIPASS`).

### Projeto
- Código organizado no pacote `printer_app` (núcleo separado da interface).
- Adicionados `requirements.txt`, `pyproject.toml` (ruff/pytest), arquivo `LICENSE` e script de build.
- Testes automatizados no GitHub Actions em Windows, com impressão real numa impressora virtual,
  autoteste do `.exe` e publicação do executável nas releases.

## [1.0.0] — 2024-08
- Primeira versão.
