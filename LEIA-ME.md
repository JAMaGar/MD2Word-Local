# MD2Word Local — Markdown para Word, sem internet

O **MD2Word Local** converte arquivos Markdown (`.md`) em documentos do Word (`.docx`, `.doc`) e em PDF. A conversão acontece **inteiramente no seu computador**: o programa não envia nada para a internet, não usa serviços online e não precisa de conta ou chave de acesso.

## 1. Privacidade

- O texto dos seus arquivos nunca sai da máquina.
- Imagens com endereço da internet (`https://...`) **não são baixadas**; no documento aparece apenas a indicação `[Imagem: ...]`.
- A internet é usada **uma única vez**, durante a geração do programa (`build.bat`), para baixar as bibliotecas de código aberto do repositório oficial do Python (PyPI). Depois de instalado, o programa funciona offline.
- O código-fonte completo está na pasta `src` (cerca de 1.800 linhas de Python) e pode ser lido e auditado.

## 2. Como gerar e instalar (uma vez só)

Pré-requisito: **Python 3.9 ou superior** instalado no Windows. Se não tiver, instale pelo site python.org ou rode no terminal `winget install Python.Python.3.13`.

1. Descompacte a pasta `MD2Word` em qualquer lugar (por exemplo, `Documentos`).
2. Dê dois cliques em **`build.bat`**. Ele cria um ambiente isolado, baixa as bibliotecas, gera o `MD2Word.exe` e, se o **Inno Setup 6** (gratuito) estiver instalado, cria o instalador. Se não estiver, o script oferece instalá-lo pelo `winget`.
3. Execute **`instalador\MD2Word_Local_Setup_1.0.0.exe`**. A instalação não exige administrador e permite escolher:
   - atalho na Área de Trabalho;
   - item **“Converter para Word (.docx)”** no menu do botão direito dos arquivos `.md`;
   - o MD2Word na lista **“Abrir com”**.

Sem o Inno Setup, o programa portátil fica em `dist\MD2Word\MD2Word.exe` e pode ser usado diretamente.

Para apenas testar sem gerar o `.exe`, use `executar_sem_instalar.bat`.

> Como o executável não tem assinatura digital, o Windows SmartScreen pode exibir o aviso “O Windows protegeu o computador”. Clique em **Mais informações → Executar assim mesmo**. Isso é normal para programas gerados localmente.

## 3. Como usar

| Forma de uso | O que fazer |
|:--|:--|
| Janela do programa | Abra o MD2Word, adicione ou arraste arquivos `.md`, escolha os formatos e clique em **Converter** (ou `Ctrl+Enter`). |
| Botão direito | Clique com o botão direito no `.md` → **Converter para Word (.docx)**. O `.docx` é criado ao lado do arquivo e aberto em seguida. No Windows 11, o item fica em **Mostrar mais opções**. |
| Abrir com | Botão direito no `.md` → **Abrir com → MD2Word Local**: abre a janela com o arquivo já na lista. |
| Colar texto | Na aba **Colar Markdown**, cole o texto (por exemplo, uma resposta copiada de um chat) e clique em **Converter**. |
| Linha de comando | `python src\conversor.py arquivo.md -f docx -f pdf` |

As opções escolhidas na janela (formatos, fonte, papel, margens etc.) ficam salvas e também valem para o atalho do botão direito.

## 4. O que é convertido

- Títulos `#` a `######` com os estilos nativos **Título 1–6** do Word, que aparecem no Painel de Navegação e no Sumário automático.
- **Negrito**, *itálico*, ~~tachado~~, `código`, links clicáveis, H<sub>2</sub>O, x<sup>2</sup>, <u>sublinhado</u> e <mark>destaque</mark>.
- Listas com marcadores, numeradas (inclusive começando em outro número), aninhadas e listas de tarefas `[ ]` / `[x]`.
- Tabelas com alinhamento por coluna; o cabeçalho se repete em cada página.
- Citações `>`, blocos de código com fundo cinza e linhas horizontais `---`.
- Notas de rodapé `[^1]`, convertidas em **notas de rodapé reais do Word**.
- Equações `$...$` e `$$...$$`, convertidas em **equações editáveis do Word** quando o Microsoft Office está instalado (o programa usa o arquivo `MML2OMML.XSL` que acompanha o Office). Sem o Office, as equações entram como texto aproximado.
- Imagens locais (caminho relativo ao `.md`) e imagens embutidas em base64, redimensionadas para caber na página.
- Quebra de página: escreva `<!-- pagebreak -->` em uma linha separada.

## 5. Formatos de saída

| Formato | Requisito |
|:--|:--|
| `.docx` | Nenhum: é gerado pelo próprio programa. |
| `.doc` (Word 97-2003) | Microsoft Word instalado (ou LibreOffice). |
| `.pdf` | Microsoft Word instalado (ou LibreOffice). |

Para `.doc` e `.pdf`, o programa abre o Word de forma invisível, salva no formato escolhido e o fecha.

## 6. Usar um modelo próprio (ex.: ABNT, timbre da instituição)

Em **Modelo .docx (opcional)**, escolha um documento do Word cujos estilos devem ser usados. O conteúdo do modelo é descartado; ficam os estilos (Normal, Título 1, 2…), as margens, o tamanho do papel, os cabeçalhos e os rodapés. Para ajustar a aparência, edite os estilos nesse arquivo no próprio Word.

## 7. Limitações conhecidas

- Imagens SVG não são suportadas pelo Word via este método; converta para PNG.
- Diagramas Mermaid aparecem como bloco de código.
- Links internos `[texto](#secao)` aparecem formatados, mas não saltam para a seção.
- Equações LaTeX muito específicas (ambientes `align`, macros próprias) podem cair no modo texto.

## 8. Alternativa por linha de comando

Se você já tem o **Pandoc** instalado, o comando `pandoc arquivo.md -o arquivo.docx` também converte localmente. O MD2Word acrescenta interface gráfica, menu de contexto, papel A4 e idioma português já configurados, saída `.doc`/`.pdf` via Word e a aba de colar texto.

## 9. Desinstalar

**Configurações → Aplicativos → MD2Word Local → Desinstalar.** Os itens do menu de contexto são removidos junto. As preferências ficam em `%APPDATA%\MD2Word\config.json` e podem ser apagadas manualmente.
