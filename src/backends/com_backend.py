"""COM Backend - Windows HWP COM interface wrapper."""

from __future__ import annotations

import os
import sys
from typing import Any

from .base import BaseBackend

# Check if we're on Windows
IS_WINDOWS = sys.platform == "win32"

if IS_WINDOWS:
    try:
        import win32com.client
        import pythoncom

        COM_AVAILABLE = True
    except ImportError:
        COM_AVAILABLE = False
else:
    COM_AVAILABLE = False


# HWP Constants
HWP_PROGID = "HWPFrame.HwpObject"


class ComBackend(BaseBackend):
    """Backend using Windows COM interface for HWP."""

    def __init__(self) -> None:
        """Initialize COM backend."""
        self._hwp: Any = None
        self._connected: bool = False

    @property
    def name(self) -> str:
        return "com"

    @property
    def supported_formats(self) -> list[str]:
        return ["HWP", "HWPX", "HWT", "PDF", "TEXT", "HTML"]

    @property
    def is_connected(self) -> bool:
        return self._connected and self._hwp is not None

    @staticmethod
    def is_available() -> bool:
        """Check if COM backend is available."""
        if not COM_AVAILABLE:
            return False

        # Try to check if HWP is installed
        try:
            import winreg

            key = winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT, "HWPFrame.HwpObject", 0, winreg.KEY_READ
            )
            winreg.CloseKey(key)
            return True
        except Exception:
            return False

    # ===================
    # Connection
    # ===================

    def connect(self, visible: bool = True, **kwargs: Any) -> bool:
        """Connect to HWP application via COM."""
        if not COM_AVAILABLE:
            raise RuntimeError("COM is not available on this platform")

        try:
            pythoncom.CoInitialize()

            # Try to connect to existing instance
            try:
                self._hwp = win32com.client.GetActiveObject(HWP_PROGID)
            except Exception:
                self._hwp = win32com.client.Dispatch(HWP_PROGID)

            # Register security module
            self._hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

            # Set visibility
            self._hwp.XHwpWindows.Item(0).Visible = visible

            self._connected = True
            return True

        except Exception as e:
            self._connected = False
            raise RuntimeError(f"Failed to connect to HWP: {e}")

    def disconnect(self) -> bool:
        """Disconnect from HWP."""
        if self._hwp is not None:
            try:
                self._hwp = None
                self._connected = False
                if COM_AVAILABLE:
                    pythoncom.CoUninitialize()
                return True
            except Exception:
                pass

        self._connected = False
        return True

    def _check_connected(self) -> None:
        """Ensure connected before operations."""
        if not self.is_connected:
            raise RuntimeError("HWP is not connected")

    # ===================
    # Document Operations
    # ===================

    def create_document(self) -> bool:
        self._check_connected()
        return self._hwp.HAction.Run("FileNew")

    def open_document(self, path: str) -> bool:
        self._check_connected()
        path = os.path.abspath(path)

        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")

        return self._hwp.Open(path)

    def save_document(self) -> bool:
        self._check_connected()
        return self._hwp.HAction.Run("FileSave")

    def save_document_as(self, path: str, format: str = "HWP") -> bool:
        self._check_connected()
        path = os.path.abspath(path)
        return self._hwp.SaveAs(path, format)

    def close_document(self) -> bool:
        self._check_connected()
        return self._hwp.HAction.Run("FileClose")

    # ===================
    # Text Operations
    # ===================

    def insert_text(self, text: str) -> bool:
        self._check_connected()
        try:
            act = self._hwp.CreateAction("InsertText")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Text", text)
            return act.Execute(pset)
        except Exception:
            return self._hwp.HAction.Run("InsertText", text)

    def get_text(self) -> str:
        self._check_connected()
        try:
            self._hwp.HAction.Run("SelectAll")
            text = self._hwp.GetTextFile("TEXT", "")
            self._hwp.HAction.Run("MoveDocBegin")
            return text
        except Exception:
            return ""

    def find_text(self, text: str) -> bool:
        self._check_connected()
        try:
            act = self._hwp.CreateAction("Find")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("FindString", text)
            return act.Execute(pset)
        except Exception:
            return False

    def replace_text(self, find: str, replace: str, replace_all: bool = False) -> int:
        self._check_connected()
        try:
            action = "ReplaceAll" if replace_all else "Replace"
            act = self._hwp.CreateAction(action)
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("FindString", find)
            pset.SetItem("ReplaceString", replace)
            result = act.Execute(pset)
            return result if isinstance(result, int) else (1 if result else 0)
        except Exception:
            return 0

    def insert_paragraph(self) -> bool:
        return self.insert_text("\r\n")

    def move_cursor(self, position: str) -> bool:
        self._check_connected()
        action_map = {
            "doc_begin": "MoveDocBegin",
            "doc_end": "MoveDocEnd",
            "line_begin": "MoveLineBegin",
            "line_end": "MoveLineEnd",
            "para_begin": "MoveParaBegin",
            "para_end": "MoveParaEnd",
            "left": "MoveLeft",
            "right": "MoveRight",
            "up": "MoveUp",
            "down": "MoveDown",
            "next_para": "MoveNextPara",
            "prev_para": "MovePrevPara",
        }

        action = action_map.get(position.lower())
        if action:
            return self._hwp.HAction.Run(action)
        return False

    # ===================
    # Table Operations
    # ===================

    def create_table(self, rows: int, cols: int) -> bool:
        self._check_connected()
        try:
            act = self._hwp.CreateAction("TableCreate")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Rows", rows)
            pset.SetItem("Cols", cols)
            pset.SetItem("WidthType", 0)
            pset.SetItem("HeightType", 1)
            return act.Execute(pset)
        except Exception:
            return False

    def get_cell_text(self, row: int, col: int) -> str:
        self._check_connected()
        try:
            self._hwp.HAction.Run("TableCellBlock")
            return self._hwp.GetTextFile("TEXT", "")
        except Exception:
            return ""

    def set_cell_text(self, row: int, col: int, text: str) -> bool:
        self._check_connected()
        return self.insert_text(text)

    def insert_row(self) -> bool:
        self._check_connected()
        return self._hwp.HAction.Run("TableAppendRow")

    def delete_row(self) -> bool:
        self._check_connected()
        return self._hwp.HAction.Run("TableDeleteRow")

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
        self._check_connected()
        try:
            act = self._hwp.CreateAction("CharShape")
            pset = act.CreateSet()
            act.GetDefault(pset)

            if name is not None:
                pset.SetItem("FaceNameHangul", name)
                pset.SetItem("FaceNameEnglish", name)
                pset.SetItem("FaceNameHanja", name)

            if size is not None:
                pset.SetItem("Height", size * 100)  # HWPUNIT

            if bold is not None:
                pset.SetItem("Bold", bold)

            if italic is not None:
                pset.SetItem("Italic", italic)

            return act.Execute(pset)
        except Exception:
            return False

    def set_alignment(self, align: str) -> bool:
        self._check_connected()
        align_map = {
            "left": 0,
            "center": 1,
            "right": 2,
            "justify": 3,
            "distribute": 4,
        }

        align_value = align_map.get(align.lower())
        if align_value is None:
            return False

        try:
            act = self._hwp.CreateAction("ParaShape")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Align", align_value)
            return act.Execute(pset)
        except Exception:
            return False

    # ===================
    # Export
    # ===================

    def export_pdf(self, path: str) -> bool:
        return self.save_document_as(path, "PDF")

    # ===================
    # Utility
    # ===================

    def get_document_info(self) -> dict[str, Any]:
        self._check_connected()
        try:
            return {
                "path": self._hwp.Path,
                "filename": self._hwp.FileName,
                "is_modified": self._hwp.IsModified,
                "backend": self.name,
            }
        except Exception:
            return {"backend": self.name}
