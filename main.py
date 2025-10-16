import os
import pypandoc
import shutil

def convert_html_to_docx(input_dir, output_dir, reference_docx):
    """
    Converte todos os arquivos HTML no diretório de entrada para o formato DOCX
    e os salva no diretório de saída, usando um arquivo de referência do Word.
    """
    # Garante que o diretório de saída exista
    os.makedirs(output_dir, exist_ok=True)

    try:
        # Verifica se o pandoc está instalado
        pypandoc.get_pandoc_path()
    except OSError:
        print("Pandoc não encontrado. Tentando baixar e instalar...")
        pypandoc.download_pandoc()

    # Lista todos os arquivos HTML no diretório de entrada
    try:
        html_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".html")]
    except FileNotFoundError:
        print(f"Erro: O diretório de entrada '{input_dir}' não foi encontrado.")
        return

    if not html_files:
        print(f"Nenhum arquivo HTML encontrado em '{input_dir}'.")
        return

    print(f"Encontrados {len(html_files)} arquivos HTML para converter.")

    # Argumentos para o Pandoc
    pandoc_args = [
        "--standalone"
    ]
    if reference_docx and os.path.exists(reference_docx):
        pandoc_args.extend(["--reference-doc", reference_docx])
        print(f"Usando o arquivo de referência: {reference_docx}")
    else:
        print("Arquivo de referência não encontrado. Usando estilos padrão do Pandoc.")


    for filename in html_files:
        html_path = os.path.join(input_dir, filename)
        docx_filename = os.path.splitext(filename)[0] + ".docx"
        docx_path = os.path.join(output_dir, docx_filename)

        print(f"Convertendo '{html_path}' para '{docx_path}'...")

        try:
            # Converte o arquivo usando pypandoc
            pypandoc.convert_file(
                html_path,
                "docx",
                outputfile=docx_path,
                extra_args=pandoc_args
            )
            print("Conversão bem-sucedida!")

        except Exception as e:
            print(f"Erro ao converter o arquivo '{html_path}': {e}")

    print("\nConversão de todos os arquivos concluída!")

if __name__ == "__main__":
    INPUT_DIR = "arq_html"
    OUTPUT_DIR = "arq_docx"
    REFERENCE_DOCX = "custom-reference.docx"
    convert_html_to_docx(INPUT_DIR, OUTPUT_DIR, REFERENCE_DOCX)