"""Base backend interface for HWP operations."""

from abc import ABC, abstractmethod
from typing import Any


class BaseBackend(ABC):
    """Abstract base class for HWP backends."""

    @property
    @abstractmethod
    def name(self) -> str:
        """Backend name."""
        pass

    @property
    @abstractmethod
    def supported_formats(self) -> list[str]:
        """List of supported file formats."""
        pass

    @property
    @abstractmethod
    def is_connected(self) -> bool:
        """Check if backend is connected/ready."""
        pass

    # ===================
    # Connection
    # ===================

    @abstractmethod
    def connect(self, **kwargs: Any) -> bool:
        """Connect to the backend."""
        pass

    @abstractmethod
    def disconnect(self) -> bool:
        """Disconnect from the backend."""
        pass

    # ===================
    # Document Operations
    # ===================

    @abstractmethod
    def create_document(self) -> bool:
        """Create a new document."""
        pass

    @abstractmethod
    def open_document(self, path: str) -> bool:
        """Open an existing document."""
        pass

    @abstractmethod
    def save_document(self) -> bool:
        """Save the current document."""
        pass

    @abstractmethod
    def save_document_as(self, path: str, format: str = "HWPX") -> bool:
        """Save document to a new path."""
        pass

    @abstractmethod
    def close_document(self) -> bool:
        """Close the current document."""
        pass

    # ===================
    # Text Operations
    # ===================

    @abstractmethod
    def insert_text(self, text: str) -> bool:
        """Insert text at current position."""
        pass

    @abstractmethod
    def get_text(self) -> str:
        """Get all text from document."""
        pass

    @abstractmethod
    def find_text(self, text: str) -> bool:
        """Find text in document."""
        pass

    @abstractmethod
    def replace_text(self, find: str, replace: str, replace_all: bool = False) -> int:
        """Replace text in document."""
        pass

    @abstractmethod
    def insert_paragraph(self) -> bool:
        """Insert a new paragraph."""
        pass

    @abstractmethod
    def move_cursor(self, position: str) -> bool:
        """Move cursor to position."""
        pass

    # ===================
    # Table Operations
    # ===================

    @abstractmethod
    def create_table(self, rows: int, cols: int) -> bool:
        """Create a table."""
        pass

    @abstractmethod
    def get_cell_text(self, row: int, col: int) -> str:
        """Get text from table cell."""
        pass

    @abstractmethod
    def set_cell_text(self, row: int, col: int, text: str) -> bool:
        """Set text in table cell."""
        pass

    @abstractmethod
    def insert_row(self) -> bool:
        """Insert a row in table."""
        pass

    @abstractmethod
    def delete_row(self) -> bool:
        """Delete a row in table."""
        pass

    # ===================
    # Formatting
    # ===================

    @abstractmethod
    def set_font(
        self,
        name: str | None = None,
        size: int | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
    ) -> bool:
        """Set font properties."""
        pass

    @abstractmethod
    def set_alignment(self, align: str) -> bool:
        """Set paragraph alignment."""
        pass

    # ===================
    # Export
    # ===================

    @abstractmethod
    def export_pdf(self, path: str) -> bool:
        """Export document as PDF."""
        pass

    # ===================
    # Utility
    # ===================

    @abstractmethod
    def get_document_info(self) -> dict[str, Any]:
        """Get document information."""
        pass
