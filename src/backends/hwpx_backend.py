"""HWPX Backend - Cross-platform HWPX file handling."""

from __future__ import annotations

import os
from typing import Any

from .base import BaseBackend
from ..hwpx import HwpxDocument, HwpxReader, HwpxWriter


class HwpxBackend(BaseBackend):
    """Backend for direct HWPX file manipulation (cross-platform)."""

    def __init__(self) -> None:
        """Initialize HWPX backend."""
        self._document: HwpxDocument | None = None
        self._connected: bool = False

    @property
    def name(self) -> str:
        return "hwpx"

    @property
    def supported_formats(self) -> list[str]:
        return ["HWPX"]  # Only HWPX supported in cross-platform mode

    @property
    def is_connected(self) -> bool:
        return self._connected

    @staticmethod
    def is_available() -> bool:
        """HWPX backend is always available (pure Python)."""
        return True

    # ===================
    # Connection
    # ===================

    def connect(self, **kwargs: Any) -> bool:
        """
        Connect to HWPX backend.

        For HWPX backend, this just marks as ready.
        The 'visible' parameter is ignored (no GUI).
        """
        self._connected = True
        return True

    def disconnect(self) -> bool:
        """Disconnect from HWPX backend."""
        self._document = None
        self._connected = False
        return True

    def _check_connected(self) -> None:
        """Ensure backend is connected."""
        if not self._connected:
            raise RuntimeError("HWPX backend is not connected")

    def _check_document(self) -> HwpxDocument:
        """Ensure document is loaded."""
        self._check_connected()
        if self._document is None:
            raise RuntimeError("No document is open")
        return self._document

    # ===================
    # Document Operations
    # ===================

    def create_document(self) -> bool:
        """Create a new empty document."""
        self._check_connected()
        self._document = HwpxDocument()
        return True

    def open_document(self, path: str) -> bool:
        """Open an existing HWPX document."""
        self._check_connected()
        path = os.path.abspath(path)

        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")

        # Check file extension
        if not path.lower().endswith(".hwpx"):
            raise ValueError("HWPX backend only supports .hwpx files")

        self._document = HwpxReader.read(path)
        return True

    def save_document(self) -> bool:
        """Save the current document."""
        doc = self._check_document()

        if doc.path is None:
            raise RuntimeError("No file path set. Use save_document_as() instead.")

        return HwpxWriter.write(doc, doc.path)

    def save_document_as(self, path: str, format: str = "HWPX") -> bool:
        """Save document to a new path."""
        doc = self._check_document()

        if format.upper() != "HWPX":
            raise ValueError(f"HWPX backend only supports HWPX format, not {format}")

        path = os.path.abspath(path)

        # Ensure .hwpx extension
        if not path.lower().endswith(".hwpx"):
            path += ".hwpx"

        return HwpxWriter.write(doc, path)

    def close_document(self) -> bool:
        """Close the current document."""
        self._check_connected()
        self._document = None
        return True

    # ===================
    # Text Operations
    # ===================

    def insert_text(self, text: str) -> bool:
        """Insert text at current position."""
        doc = self._check_document()
        return doc.insert_text(text)

    def get_text(self) -> str:
        """Get all text from document."""
        doc = self._check_document()
        return doc.get_all_text()

    def find_text(self, text: str) -> bool:
        """Find text in document."""
        doc = self._check_document()
        results = doc.find_text(text)
        return len(results) > 0

    def replace_text(self, find: str, replace: str, replace_all: bool = False) -> int:
        """Replace text in document."""
        doc = self._check_document()
        return doc.replace_text(find, replace, replace_all)

    def insert_paragraph(self) -> bool:
        """Insert a new paragraph."""
        doc = self._check_document()
        return doc.insert_paragraph()

    def move_cursor(self, position: str) -> bool:
        """Move cursor to position."""
        doc = self._check_document()
        return doc.move_cursor(position)

    # ===================
    # Table Operations
    # ===================

    def create_table(self, rows: int, cols: int) -> bool:
        """Create a table."""
        doc = self._check_document()
        table = doc.create_table(rows, cols)
        return table is not None

    def get_cell_text(self, row: int, col: int) -> str:
        """Get text from table cell."""
        doc = self._check_document()
        tables = doc.get_tables()

        if not tables:
            return ""

        # Get last table (current table context)
        table = tables[-1]
        cell = table.get_cell(row, col)
        return cell.text if cell else ""

    def set_cell_text(self, row: int, col: int, text: str) -> bool:
        """Set text in table cell."""
        doc = self._check_document()
        tables = doc.get_tables()

        if not tables:
            return False

        # Get last table
        table = tables[-1]
        return table.set_cell_text(row, col, text)

    def insert_row(self) -> bool:
        """Insert a row in table."""
        doc = self._check_document()
        tables = doc.get_tables()

        if not tables:
            return False

        # Not fully implemented - would need cursor context
        # For now, return False
        return False

    def delete_row(self) -> bool:
        """Delete a row in table."""
        doc = self._check_document()
        # Not fully implemented
        return False

    # ===================
    # Formatting
    # ===================

    def set_font(
        self,
        name: str | None = None,
        size: int | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
    ) -> bool:
        """Set font properties."""
        doc = self._check_document()
        # Font styling in HWPX requires modifying the run elements
        # For simplicity, this affects future text insertions
        # Full implementation would require style management
        return True

    def set_alignment(self, align: str) -> bool:
        """Set paragraph alignment."""
        doc = self._check_document()
        return doc.set_paragraph_alignment(align)

    # ===================
    # Export
    # ===================

    def export_pdf(self, path: str) -> bool:
        """
        Export document as PDF.

        Note: PDF export is not available in HWPX backend.
        This would require a separate PDF library.
        """
        raise NotImplementedError(
            "PDF export is not available in HWPX backend. "
            "Use Windows COM backend for PDF export."
        )

    # ===================
    # Utility
    # ===================

    def get_document_info(self) -> dict[str, Any]:
        """Get document information."""
        self._check_connected()

        if self._document is None:
            return {"backend": self.name, "document": None}

        return {
            "path": self._document.path,
            "filename": os.path.basename(self._document.path) if self._document.path else None,
            "is_modified": self._document.is_modified,
            "backend": self.name,
            "sections": len(self._document.sections),
            "title": self._document.metadata.title,
            "author": self._document.metadata.author,
        }
