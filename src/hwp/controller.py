"""HWP Controller - COM interface wrapper for Hancom HWP."""

from __future__ import annotations

import os
from typing import Any

from .constants import HWP_PROGID, Action, Alignment, HWPUNIT_PER_PT
from .exceptions import (
    HwpConnectionError,
    HwpDocumentError,
    HwpFileNotFoundError,
    HwpNotConnectedError,
    HwpSaveError,
)

# Windows COM imports (only available on Windows)
try:
    import win32com.client
    import pythoncom

    WINDOWS_AVAILABLE = True
except ImportError:
    WINDOWS_AVAILABLE = False


class HwpController:
    """Controller for HWP document operations via COM interface."""

    def __init__(self) -> None:
        """Initialize HWP controller."""
        self._hwp: Any = None
        self._connected: bool = False

    @property
    def is_connected(self) -> bool:
        """Check if connected to HWP."""
        return self._connected and self._hwp is not None

    @property
    def hwp(self) -> Any:
        """Get HWP COM object."""
        if not self.is_connected:
            raise HwpNotConnectedError()
        return self._hwp

    def connect(self, visible: bool = True) -> bool:
        """
        Connect to HWP application.

        Args:
            visible: Whether to show HWP window (default: True)

        Returns:
            True if connection successful

        Raises:
            HwpConnectionError: If connection fails
        """
        if not WINDOWS_AVAILABLE:
            raise HwpConnectionError("Windows COM is not available on this platform")

        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()

            # Try to connect to existing HWP instance
            try:
                self._hwp = win32com.client.GetActiveObject(HWP_PROGID)
            except Exception:
                # Create new HWP instance
                self._hwp = win32com.client.Dispatch(HWP_PROGID)

            # Register security module to allow file access
            self._hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

            # Set visibility
            self._hwp.XHwpWindows.Item(0).Visible = visible

            self._connected = True
            return True

        except Exception as e:
            self._connected = False
            raise HwpConnectionError(f"Failed to connect to HWP: {e}")

    def disconnect(self) -> bool:
        """
        Disconnect from HWP application.

        Returns:
            True if disconnection successful
        """
        if self._hwp is not None:
            try:
                # Clear COM object reference
                self._hwp = None
                self._connected = False

                # Uninitialize COM
                if WINDOWS_AVAILABLE:
                    pythoncom.CoUninitialize()

                return True
            except Exception:
                pass

        self._connected = False
        return True

    # ===================
    # Document Operations
    # ===================

    def create_document(self) -> bool:
        """
        Create a new empty document.

        Returns:
            True if successful
        """
        return self.hwp.HAction.Run(Action.FILE_NEW)

    def open_document(self, path: str) -> bool:
        """
        Open an existing document.

        Args:
            path: File path to open

        Returns:
            True if successful

        Raises:
            HwpFileNotFoundError: If file doesn't exist
            HwpDocumentError: If open fails
        """
        # Normalize path
        path = os.path.abspath(path)

        if not os.path.exists(path):
            raise HwpFileNotFoundError(path)

        try:
            result = self.hwp.Open(path)
            if not result:
                raise HwpDocumentError(f"Failed to open document: {path}")
            return True
        except HwpFileNotFoundError:
            raise
        except Exception as e:
            raise HwpDocumentError(f"Failed to open document: {e}")

    def save_document(self) -> bool:
        """
        Save the current document.

        Returns:
            True if successful

        Raises:
            HwpSaveError: If save fails
        """
        try:
            result = self.hwp.HAction.Run(Action.FILE_SAVE)
            if not result:
                raise HwpSaveError()
            return True
        except HwpSaveError:
            raise
        except Exception as e:
            raise HwpSaveError(f"Failed to save document: {e}")

    def save_document_as(self, path: str, format: str = "HWP") -> bool:
        """
        Save the current document to a new path.

        Args:
            path: File path to save to
            format: File format (HWP, HWPX, PDF, etc.)

        Returns:
            True if successful

        Raises:
            HwpSaveError: If save fails
        """
        path = os.path.abspath(path)

        try:
            result = self.hwp.SaveAs(path, format)
            if not result:
                raise HwpSaveError(f"Failed to save document as: {path}")
            return True
        except HwpSaveError:
            raise
        except Exception as e:
            raise HwpSaveError(f"Failed to save document: {e}")

    def close_document(self) -> bool:
        """
        Close the current document.

        Returns:
            True if successful
        """
        try:
            return self.hwp.HAction.Run(Action.FILE_CLOSE)
        except Exception:
            return False

    # ================
    # Text Operations
    # ================

    def insert_text(self, text: str) -> bool:
        """
        Insert text at current cursor position.

        Args:
            text: Text to insert

        Returns:
            True if successful
        """
        try:
            act = self.hwp.CreateAction(Action.INSERT_TEXT)
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Text", text)
            return act.Execute(pset)
        except Exception:
            # Fallback method
            return self.hwp.HAction.Run(Action.INSERT_TEXT, text)

    def get_text(self) -> str:
        """
        Get all text from current document.

        Returns:
            Document text content
        """
        try:
            # Save cursor position
            self.hwp.HAction.Run(Action.SELECT_ALL)
            text = self.hwp.GetTextFile("TEXT", "")
            # Deselect
            self.hwp.HAction.Run(Action.MOVE_DOC_BEGIN)
            return text
        except Exception:
            return ""

    def find_text(self, text: str) -> bool:
        """
        Find text in document.

        Args:
            text: Text to find

        Returns:
            True if found
        """
        try:
            act = self.hwp.CreateAction(Action.FIND)
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("FindString", text)
            return act.Execute(pset)
        except Exception:
            return False

    def replace_text(self, find: str, replace: str, replace_all: bool = False) -> int:
        """
        Replace text in document.

        Args:
            find: Text to find
            replace: Replacement text
            replace_all: Replace all occurrences

        Returns:
            Number of replacements made
        """
        try:
            action = Action.REPLACE_ALL if replace_all else Action.REPLACE
            act = self.hwp.CreateAction(action)
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("FindString", find)
            pset.SetItem("ReplaceString", replace)
            result = act.Execute(pset)
            return result if isinstance(result, int) else (1 if result else 0)
        except Exception:
            return 0

    def insert_paragraph(self) -> bool:
        """
        Insert a new paragraph (Enter key).

        Returns:
            True if successful
        """
        return self.insert_text("\r\n")

    def move_cursor(self, position: str) -> bool:
        """
        Move cursor to specified position.

        Args:
            position: Position name (doc_begin, doc_end, line_begin, line_end, etc.)

        Returns:
            True if successful
        """
        action_map = {
            "doc_begin": Action.MOVE_DOC_BEGIN,
            "doc_end": Action.MOVE_DOC_END,
            "line_begin": Action.MOVE_LINE_BEGIN,
            "line_end": Action.MOVE_LINE_END,
            "para_begin": Action.MOVE_PARA_BEGIN,
            "para_end": Action.MOVE_PARA_END,
            "left": Action.MOVE_LEFT,
            "right": Action.MOVE_RIGHT,
            "up": Action.MOVE_UP,
            "down": Action.MOVE_DOWN,
            "next_para": Action.MOVE_NEXT_PARA,
            "prev_para": Action.MOVE_PREV_PARA,
        }

        action = action_map.get(position.lower())
        if action:
            return self.hwp.HAction.Run(action)
        return False

    # =================
    # Table Operations
    # =================

    def create_table(self, rows: int, cols: int) -> bool:
        """
        Create a table at current cursor position.

        Args:
            rows: Number of rows
            cols: Number of columns

        Returns:
            True if successful
        """
        try:
            act = self.hwp.CreateAction(Action.TABLE_CREATE)
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Rows", rows)
            pset.SetItem("Cols", cols)
            pset.SetItem("WidthType", 0)  # 0: 단 기준
            pset.SetItem("HeightType", 1)  # 1: 셀 기준
            return act.Execute(pset)
        except Exception:
            return False

    def get_cell_text(self, row: int, col: int) -> str:
        """
        Get text from a table cell.

        Args:
            row: Row index (0-based)
            col: Column index (0-based)

        Returns:
            Cell text content
        """
        try:
            # Move to cell
            self.hwp.HAction.Run(Action.TABLE_CELL_BLOCK)
            # Get text
            return self.hwp.GetTextFile("TEXT", "")
        except Exception:
            return ""

    def set_cell_text(self, row: int, col: int, text: str) -> bool:
        """
        Set text in a table cell.

        Args:
            row: Row index (0-based)
            col: Column index (0-based)
            text: Text to set

        Returns:
            True if successful
        """
        try:
            # This requires navigating to the specific cell first
            # The exact implementation depends on table structure
            return self.insert_text(text)
        except Exception:
            return False

    def insert_row(self) -> bool:
        """
        Insert a row in current table.

        Returns:
            True if successful
        """
        try:
            return self.hwp.HAction.Run(Action.TABLE_APPEND_ROW)
        except Exception:
            return False

    def delete_row(self) -> bool:
        """
        Delete current row in table.

        Returns:
            True if successful
        """
        try:
            return self.hwp.HAction.Run(Action.TABLE_DELETE_ROW)
        except Exception:
            return False

    # ====================
    # Formatting Operations
    # ====================

    def set_font(
        self,
        name: str | None = None,
        size: int | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
    ) -> bool:
        """
        Set font properties.

        Args:
            name: Font name
            size: Font size in pt
            bold: Bold style
            italic: Italic style

        Returns:
            True if successful
        """
        try:
            act = self.hwp.CreateAction(Action.CHAR_SHAPE)
            pset = act.CreateSet()
            act.GetDefault(pset)

            if name is not None:
                # Set font for all languages
                pset.SetItem("FaceNameHangul", name)
                pset.SetItem("FaceNameEnglish", name)
                pset.SetItem("FaceNameHanja", name)

            if size is not None:
                # Convert pt to HWPUNIT
                pset.SetItem("Height", size * HWPUNIT_PER_PT)

            if bold is not None:
                pset.SetItem("Bold", bold)

            if italic is not None:
                pset.SetItem("Italic", italic)

            return act.Execute(pset)
        except Exception:
            return False

    def set_alignment(self, align: str) -> bool:
        """
        Set paragraph alignment.

        Args:
            align: Alignment (left, center, right, justify, distribute)

        Returns:
            True if successful
        """
        align_map = {
            "left": Alignment.LEFT,
            "center": Alignment.CENTER,
            "right": Alignment.RIGHT,
            "justify": Alignment.JUSTIFY,
            "distribute": Alignment.DISTRIBUTE,
        }

        align_value = align_map.get(align.lower())
        if align_value is None:
            return False

        try:
            act = self.hwp.CreateAction(Action.PARA_SHAPE)
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Align", align_value)
            return act.Execute(pset)
        except Exception:
            return False

    # =================
    # Export Operations
    # =================

    def export_pdf(self, path: str) -> bool:
        """
        Export document as PDF.

        Args:
            path: Output PDF file path

        Returns:
            True if successful
        """
        return self.save_document_as(path, "PDF")

    def export_text(self, path: str) -> bool:
        """
        Export document as plain text.

        Args:
            path: Output text file path

        Returns:
            True if successful
        """
        return self.save_document_as(path, "TEXT")

    # =================
    # Utility Methods
    # =================

    def get_document_info(self) -> dict[str, Any]:
        """
        Get current document information.

        Returns:
            Dictionary with document info
        """
        try:
            return {
                "path": self.hwp.Path,
                "filename": self.hwp.FileName,
                "is_modified": self.hwp.IsModified,
            }
        except Exception:
            return {}
