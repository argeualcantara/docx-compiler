import re
import requests
from docx import Document
from pathlib import Path
from .document_handler_interface import DocumentHandler

class GoogleDocsHandler(DocumentHandler):
    def __init__(self):
        pass

    def extract_doc_id(self, url: str) -> str:
        """
        Extract the Google Doc ID from a Google Docs URL.

        Args:
            url (str): The full URL of the Google Doc.

        Returns:
            str: The extracted document ID if found, otherwise None.

        Example:
            doc_id = extract_doc_id("https://docs.google.com/document/d/ABC123XYZ456/edit")
            print(doc_id)  # Output: ABC123XYZ456
        """
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        return match.group(1) if match else None


    def download_doc(self, doc_id: str, dest_path: Path):
        """
        Download a Google Doc as a DOCX file given its document ID.

        Args:
            doc_id (str): The Google Docs document ID to download.
            dest_path (Path): The file path where the downloaded DOCX will be saved.

        Behavior:
            - Constructs the Google Docs export URL.
            - Sends an HTTP GET request with a user-agent header.
            - Saves the DOCX content to the specified destination path.
            - Raises an exception if the request fails.

        Example:
            download_doc("ABC123XYZ456", Path("downloaded_doc.docx"))
        """
        url = f"https://docs.google.com/document/d/{doc_id}/export?format=docx"
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        with open(dest_path, "wb") as f:
            f.write(response.content)
