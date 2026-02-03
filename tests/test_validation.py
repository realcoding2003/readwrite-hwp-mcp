"""Tests for validation utilities."""

import pytest
from src.utils.validation import (
    validate_file_path,
    validate_format,
    validate_position,
    validate_alignment,
    validate_table_dimensions,
)


class TestValidateFilePath:
    """Tests for validate_file_path function."""

    def test_valid_path(self):
        is_valid, error = validate_file_path("C:\\Documents\\test.hwp")
        assert is_valid is True
        assert error == ""

    def test_empty_path(self):
        is_valid, error = validate_file_path("")
        assert is_valid is False
        assert "empty" in error.lower()

    def test_path_traversal(self):
        is_valid, error = validate_file_path("..\\..\\secret.txt")
        assert is_valid is False
        assert "traversal" in error.lower()


class TestValidateFormat:
    """Tests for validate_format function."""

    def test_valid_formats(self):
        for fmt in ["HWP", "HWPX", "PDF", "TEXT", "HTML"]:
            is_valid, error = validate_format(fmt)
            assert is_valid is True

    def test_invalid_format(self):
        is_valid, error = validate_format("DOCX")
        assert is_valid is False


class TestValidatePosition:
    """Tests for validate_position function."""

    def test_valid_positions(self):
        for pos in ["doc_begin", "doc_end", "left", "right"]:
            is_valid, error = validate_position(pos)
            assert is_valid is True

    def test_invalid_position(self):
        is_valid, error = validate_position("invalid")
        assert is_valid is False


class TestValidateAlignment:
    """Tests for validate_alignment function."""

    def test_valid_alignments(self):
        for align in ["left", "center", "right", "justify"]:
            is_valid, error = validate_alignment(align)
            assert is_valid is True

    def test_invalid_alignment(self):
        is_valid, error = validate_alignment("invalid")
        assert is_valid is False


class TestValidateTableDimensions:
    """Tests for validate_table_dimensions function."""

    def test_valid_dimensions(self):
        is_valid, error = validate_table_dimensions(5, 3)
        assert is_valid is True

    def test_zero_rows(self):
        is_valid, error = validate_table_dimensions(0, 3)
        assert is_valid is False

    def test_too_many_rows(self):
        is_valid, error = validate_table_dimensions(1001, 3)
        assert is_valid is False
