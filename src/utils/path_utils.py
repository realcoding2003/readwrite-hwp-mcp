"""Path utility functions."""

import os
from pathlib import Path


def normalize_path(path: str) -> str:
    """
    Normalize file path to absolute path.

    Args:
        path: File path (relative or absolute)

    Returns:
        Absolute normalized path
    """
    return os.path.abspath(os.path.expanduser(path))


def ensure_directory(path: str) -> None:
    """
    Ensure directory exists for the given file path.

    Args:
        path: File path
    """
    directory = os.path.dirname(path)
    if directory:
        Path(directory).mkdir(parents=True, exist_ok=True)


def get_file_extension(path: str) -> str:
    """
    Get file extension from path.

    Args:
        path: File path

    Returns:
        File extension (lowercase, without dot)
    """
    return os.path.splitext(path)[1].lower().lstrip(".")


def is_hwp_file(path: str) -> bool:
    """
    Check if file is an HWP document.

    Args:
        path: File path

    Returns:
        True if HWP or HWPX file
    """
    ext = get_file_extension(path)
    return ext in ("hwp", "hwpx", "hwt")
