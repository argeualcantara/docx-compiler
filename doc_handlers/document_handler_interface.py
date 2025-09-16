from abc import ABC, abstractmethod
from pathlib import Path


class DocumentHandler(ABC):
    """
    Interface for document handlers.

    Defines the required methods for handling documents from external sources,
    such as extracting IDs from URLs and downloading documents.
    """

    @abstractmethod
    def extract_doc_id(self, url: str) -> str:
        """
        Extract the document ID from a given URL.

        Args:
            url (str): The full document URL.

        Returns:
            str: The extracted document ID if found, otherwise None.
        """
        pass

    @abstractmethod
    def download_doc(self, doc_id: str, dest_path: Path):
        """
        Download the document as a DOCX file given its document ID.

        Args:
            doc_id (str): The document ID to download.
            dest_path (Path): The file path where the downloaded DOCX will be saved.
        """
        pass
