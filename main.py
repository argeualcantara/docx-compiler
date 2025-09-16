import argparse
from pathlib import Path
from compiler import DocCompiler

def main():
    parser = argparse.ArgumentParser(description="Compile Google Docs referenced in a DOCX.")
    parser.add_argument("-s", "--source", type=str, default="source.docx", help="Path to the source DOCX file containing hyperlinks")
    parser.add_argument("-o", "--output", type=str, default="compiled_final.docx", help="Name of the final compiled DOCX file")
    parser.add_argument("--temp-dir", type=str, default="temp_docs", help="Temporary directory for downloaded DOCX files")

    args = parser.parse_args()

    source = Path(args.source)
    final = Path(args.output)
    temp_dir = Path(args.temp_dir)

    if not source.exists():
        print(f"Source file not found: {source}")
        return

    compiler = DocCompiler()
    compiler.compile_to_docx(source, final, temp_dir)

if __name__ == "__main__":
    main()
