import re
import requests
from docx import Document
from pathlib import Path

class GooogleDocsHandler:
    def __init__(self):
        pass

    def extrair_doc_id(self, url: str) -> str:
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        return match.group(1) if match else None


    def baixar_doc(self, doc_id: str, dest_path: Path):
        url = f"https://docs.google.com/document/d/{doc_id}/export?format=docx"
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        with open(dest_path, "wb") as f:
            f.write(response.content)
