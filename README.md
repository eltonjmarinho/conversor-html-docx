# Conversor de HTML para DOCX

Este projeto converte arquivos HTML em documentos DOCX, aplicando estilos a partir de um arquivo de referência.

## Funcionalidades

- Converte múltiplos arquivos `.html` de uma pasta de entrada (`arq_html`).
- Salva os arquivos convertidos `.docx` em uma pasta de saída (`arq_docx`).
- Utiliza um arquivo `custom-reference.docx` para aplicar estilos consistentes aos documentos gerados.
- Baixa e instala o Pandoc automaticamente se não for encontrado no sistema.

## Como Usar

### 1. Pré-requisitos

- Python 3

### 2. Instalação

Clone este repositório e instale as dependências:

```bash
pip install -r requirements.txt
```

### 3. Estrutura de Pastas

- Coloque todos os seus arquivos HTML na pasta `arq_html`.
- O arquivo `custom-reference.docx` deve estar na raiz do projeto para ser usado como modelo de estilo.

### 4. Execução

Para iniciar a conversão, execute o script principal:

```bash
python main.py
```

Os arquivos convertidos serão salvos na pasta `arq_docx`.
