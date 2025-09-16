from pathlib import Path
from docx import Document
import sys
sys.path.append(str(Path(__file__).resolve().parent))

from utils.docx_utils import DocxUtils
from doc_handlers.google_docs_handler import GooogleDocsHandler

class DocCompiler:
    def __init__(self):
        self.docx_utils = DocxUtils()  # sua classe de utilitários
        self.google_handler = GooogleDocsHandler()
    
    def detectar_tipo_documento_link(self, link: str) -> str:
        """
        Retorna o tipo de documento baseado no link.
        Exemplo atual: apenas Google Docs.
        """
        if "docs.google.com/document/d/" in link:
            return "google_doc"
        # Aqui você pode adicionar outros tipos no futuro, ex: word online, pdf, etc.
        return "desconhecido"

    def compile_to_docx(self, origem: Path, final_path: Path, temp_dir: Path):
        temp_dir.mkdir(exist_ok=True)

        docx_utils = DocxUtils()

        links = docx_utils.extrair_hyperlinks_por_linha(origem)
        print(f"{len(links)} links encontrados.")

        doc_final = Document()
        doc_final.add_heading("Documento Compilado", level=1)

        for idx, link in enumerate(links, start=1):
            file_type = self.detectar_tipo_documento_link(link)
            handler = None

            if file_type == "google_doc":
                handler = self.google_handler
            else:
                print("Tipo de arquivo não suportado, por enquanto apenas links do google docs podem ser utilizados")
                continue

            print(f"\nProcessando link {idx}/{len(links)}: {link}")

            doc_id = handler.extrair_doc_id(link)
            if not doc_id:
                print(f"ID do tipo {file_type} não encontrado na URL: {link}")
                continue

            temp_doc_path = temp_dir / f"{doc_id}.docx"

            try:
                handler.baixar_doc(doc_id, temp_doc_path)
                print(f"Doc baixado: {temp_doc_path.name}")

                # Copia conteúdo mantendo texto, formatação e imagens
                docx_utils.copiar_docx_com_imagens(temp_doc_path, doc_final)
                print(f"Conteúdo do Doc copiado para o docx compilado.")

                # Quebra de página após cada documento, exceto o último
                if idx < len(links):
                    doc_final.add_page_break()

            except Exception as e:
                print(f"Erro ao processar {link}: {e}")
                doc_final.add_paragraph(f"Não foi possível processar o link {link}: {e}")

        doc_final.save(final_path)
        print(f"\nProcesso finalizado! Documento compilado salvo em: {final_path}")
