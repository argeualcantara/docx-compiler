from docx import Document
from docx.shared import Cm
from pathlib import Path

class DocxUtils:
    def __init__(self):
        pass

    def extrair_hyperlinks_por_linha(self, docx_path: Path) -> list[str]:
        """
        Extrai todos os hyperlinks do DOCX, na ordem em que aparecem no documento.
        Retorna uma lista de URLs.
        """
        doc = Document(docx_path)
        urls = []

        # Mapeia todas as relações do tipo hyperlink
        rel_dict = {rel.rId: rel.target_ref for rel in doc.part.rels.values() if "hyperlink" in rel.reltype}

        # Percorre cada parágrafo
        for line in doc.paragraphs:
            # Cada hyperlink no XML do parágrafo
            for hyperlink in line._p.findall('.//w:hyperlink', doc.part._element.nsmap):
                rId = hyperlink.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rId and rId in rel_dict:
                    urls.append(rel_dict[rId])

        return urls

    def copiar_docx_com_imagens(self, src_path: Path, dest_doc: Document):
        src_doc = Document(src_path)

        for line in src_doc.paragraphs:
            # Novo parágrafo no destino
            new_para = dest_doc.add_paragraph(style=line.style)

            # Copia runs
            for run in line.runs:
                new_run = new_line.add_run(run.text)
                new_run.bold = run.bold
                new_run.italic = run.italic
                new_run.underline = run.underline
                new_run.font.size = run.font.size
                new_run.font.name = run.font.name

            # Copiar imagens (inline)
            for drawing in line._p.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/picture}pic'):
                blip = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                if blip is not None:
                    rEmbed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if rEmbed in src_doc.part.related_parts:
                        image_part = src_doc.part.related_parts[rEmbed]
                        image_bytes = image_part.blob

                        # Salva temporariamente
                        temp_img_path = Path("temp_image.png")
                        with open(temp_img_path, "wb") as f:
                            f.write(image_bytes)

                        # Insere imagem no documento destino
                        dest_doc.add_picture(str(temp_img_path), width=Cm(12))
                        temp_img_path.unlink()
