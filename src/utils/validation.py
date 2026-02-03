"""Input validation utilities."""

import os
import re

# Valid file formats for HWP
VALID_FORMATS = {"HWP", "HWPX", "HWT", "HTML", "TEXT", "PDF"}

# Valid cursor positions
VALID_POSITIONS = {
    "doc_begin",
    "doc_end",
    "line_begin",
    "line_end",
    "para_begin",
    "para_end",
    "left",
    "right",
    "up",
    "down",
    "next_para",
    "prev_para",
}

# Valid alignment values
VALID_ALIGNMENTS = {"left", "center", "right", "justify", "distribute"}


def validate_file_path(path: str) -> tuple[bool, str]:
    """
    Validate file path for safety.

    Args:
        path: File path to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not path:
        return False, "Path cannot be empty"

    # Check for path traversal attempts
    normalized = os.path.normpath(path)
    if ".." in normalized:
        return False, "Path traversal not allowed"

    # Check for invalid characters (Windows)
    invalid_chars = r'[<>:"|?*]'
    if re.search(invalid_chars, os.path.basename(path)):
        return False, "Path contains invalid characters"

    return True, ""


def validate_format(format: str) -> tuple[bool, str]:
    """
    Validate file format.

    Args:
        format: File format to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if format.upper() not in VALID_FORMATS:
        return False, f"Invalid format. Valid formats: {', '.join(VALID_FORMATS)}"
    return True, ""


def validate_position(position: str) -> tuple[bool, str]:
    """
    Validate cursor position.

    Args:
        position: Position to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if position.lower() not in VALID_POSITIONS:
        return False, f"Invalid position. Valid positions: {', '.join(VALID_POSITIONS)}"
    return True, ""


def validate_alignment(align: str) -> tuple[bool, str]:
    """
    Validate alignment value.

    Args:
        align: Alignment to validate

    Returns:
        Tuple of (is_valid, error_message)
    """
    if align.lower() not in VALID_ALIGNMENTS:
        return False, f"Invalid alignment. Valid values: {', '.join(VALID_ALIGNMENTS)}"
    return True, ""


def validate_table_dimensions(rows: int, cols: int) -> tuple[bool, str]:
    """
    Validate table dimensions.

    Args:
        rows: Number of rows
        cols: Number of columns

    Returns:
        Tuple of (is_valid, error_message)
    """
    if rows < 1:
        return False, "Rows must be at least 1"
    if cols < 1:
        return False, "Columns must be at least 1"
    if rows > 1000:
        return False, "Rows cannot exceed 1000"
    if cols > 100:
        return False, "Columns cannot exceed 100"
    return True, ""
