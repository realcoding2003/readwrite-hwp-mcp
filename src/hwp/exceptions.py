"""Custom exceptions for HWP operations."""


class HwpError(Exception):
    """Base exception for HWP operations."""

    def __init__(self, message: str, code: str | None = None):
        self.message = message
        self.code = code
        super().__init__(self.message)


class HwpConnectionError(HwpError):
    """Exception raised when HWP connection fails."""

    def __init__(self, message: str = "Failed to connect to HWP"):
        super().__init__(message, code="CONNECTION_ERROR")


class HwpDocumentError(HwpError):
    """Exception raised for document operation errors."""

    def __init__(self, message: str, code: str = "DOCUMENT_ERROR"):
        super().__init__(message, code=code)


class HwpNotConnectedError(HwpError):
    """Exception raised when HWP is not connected."""

    def __init__(self, message: str = "HWP is not connected"):
        super().__init__(message, code="NOT_CONNECTED")


class HwpFileNotFoundError(HwpError):
    """Exception raised when file is not found."""

    def __init__(self, path: str):
        super().__init__(f"File not found: {path}", code="FILE_NOT_FOUND")


class HwpSaveError(HwpError):
    """Exception raised when save operation fails."""

    def __init__(self, message: str = "Failed to save document"):
        super().__init__(message, code="SAVE_ERROR")
