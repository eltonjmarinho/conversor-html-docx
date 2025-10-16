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

Recomenda-se o uso de um ambiente virtual para isolar as dependências do projeto.

1. **Crie e ative um ambiente virtual:**

   - No Windows:
     ```bash
     python -m venv .venv
     .venv\Scripts\activate
     ```

   - No macOS & Linux:
     ```bash
     python3 -m venv .venv
     source .venv/bin/activate
     ```

2. **Instale as dependências:**

   Com o ambiente virtual ativado, instale as bibliotecas necessárias a partir do arquivo `requirements.txt`:

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
