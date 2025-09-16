from pathlib import Path
import argparse
from compiler import DocCompiler

def main():
    parser = argparse.ArgumentParser(description="Compilar Google Docs referenciados em um DOCX.")
    parser.add_argument("-s", "--source", type=str, default="source.docx", help="Caminho do arquivo DOCX de origem com hyperlinks")
    parser.add_argument("-o", "--output", type=str, default="compilado_final.docx", help="Nome do arquivo DOCX compilado final")
    parser.add_argument("--temp-dir", type=str, default="temp_docs", help="Diretório temporário para DOCX baixados")

    args = parser.parse_args()

    origem = Path(args.source)
    final = Path(args.output)
    temp_dir = Path(args.temp_dir)

    if not origem.exists():
        print(f"Arquivo de origem não encontrado: {origem}")
        return

    compiler = DocCompiler()
    compiler.compile_to_docx(origem, final, temp_dir)

if __name__ == "__main__":
    main()
