"""HWPX Document - In-memory document representation."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class TextRun:
    """Text run with formatting."""

    text: str
    char_style_id: str | None = None
    bold: bool = False
    italic: bool = False
    font_name: str | None = None
    font_size: int | None = None


@dataclass
class Paragraph:
    """Paragraph containing text runs."""

    id: str
    runs: list[TextRun] = field(default_factory=list)
    para_style_id: str | None = None
    alignment: str = "left"

    @property
    def text(self) -> str:
        """Get plain text content."""
        return "".join(run.text for run in self.runs)

    def add_text(self, text: str, **style: Any) -> None:
        """Add text with optional styling."""
        self.runs.append(TextRun(text=text, **style))


@dataclass
class TableCell:
    """Table cell."""

    row: int
    col: int
    paragraphs: list[Paragraph] = field(default_factory=list)
    row_span: int = 1
    col_span: int = 1

    @property
    def text(self) -> str:
        """Get cell text content."""
        return "\n".join(p.text for p in self.paragraphs)


@dataclass
class Table:
    """Table structure."""

    id: str
    rows: int
    cols: int
    cells: list[list[TableCell]] = field(default_factory=list)

    def __post_init__(self) -> None:
        """Initialize cells grid."""
        if not self.cells:
            self.cells = [
                [TableCell(row=r, col=c) for c in range(self.cols)]
                for r in range(self.rows)
            ]

    def get_cell(self, row: int, col: int) -> TableCell | None:
        """Get cell at position."""
        if 0 <= row < self.rows and 0 <= col < self.cols:
            return self.cells[row][col]
        return None

    def set_cell_text(self, row: int, col: int, text: str) -> bool:
        """Set cell text."""
        cell = self.get_cell(row, col)
        if cell:
            if not cell.paragraphs:
                cell.paragraphs.append(Paragraph(id=f"cell_{row}_{col}_p0"))
            cell.paragraphs[0].runs = [TextRun(text=text)]
            return True
        return False


@dataclass
class Section:
    """Document section."""

    id: str
    paragraphs: list[Paragraph] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)


@dataclass
class DocumentMetadata:
    """Document metadata."""

    title: str = ""
    author: str = ""
    subject: str = ""
    keywords: str = ""
    created: str = ""
    modified: str = ""


@dataclass
class HwpxDocument:
    """In-memory HWPX document representation."""

    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    sections: list[Section] = field(default_factory=list)
    styles: dict[str, dict] = field(default_factory=dict)

    # Current cursor position
    _current_section: int = 0
    _current_para: int = 0
    _modified: bool = False
    _path: str | None = None

    def __post_init__(self) -> None:
        """Ensure at least one section exists."""
        if not self.sections:
            self.sections.append(Section(id="sec_0"))
            self.sections[0].paragraphs.append(Paragraph(id="para_0"))

    @property
    def is_modified(self) -> bool:
        """Check if document has been modified."""
        return self._modified

    @property
    def path(self) -> str | None:
        """Get document file path."""
        return self._path

    @property
    def current_section(self) -> Section:
        """Get current section."""
        return self.sections[self._current_section]

    @property
    def current_paragraph(self) -> Paragraph:
        """Get current paragraph."""
        section = self.current_section
        if self._current_para < len(section.paragraphs):
            return section.paragraphs[self._current_para]
        return section.paragraphs[-1]

    def get_all_text(self) -> str:
        """Get all text content from document."""
        texts = []
        for section in self.sections:
            for para in section.paragraphs:
                texts.append(para.text)
        return "\n".join(texts)

    def insert_text(self, text: str, **style: Any) -> bool:
        """Insert text at current position."""
        self.current_paragraph.add_text(text, **style)
        self._modified = True
        return True

    def insert_paragraph(self) -> bool:
        """Insert new paragraph."""
        section = self.current_section
        new_para_id = f"para_{len(section.paragraphs)}"
        new_para = Paragraph(id=new_para_id)
        section.paragraphs.append(new_para)
        self._current_para = len(section.paragraphs) - 1
        self._modified = True
        return True

    def create_table(self, rows: int, cols: int) -> Table | None:
        """Create table at current position."""
        section = self.current_section
        table_id = f"table_{len(section.tables)}"
        table = Table(id=table_id, rows=rows, cols=cols)
        section.tables.append(table)
        self._modified = True
        return table

    def get_tables(self) -> list[Table]:
        """Get all tables in document."""
        tables = []
        for section in self.sections:
            tables.extend(section.tables)
        return tables

    def find_text(self, query: str) -> list[tuple[int, int, int]]:
        """
        Find text in document.

        Returns:
            List of (section_idx, para_idx, char_offset) tuples
        """
        results = []
        for s_idx, section in enumerate(self.sections):
            for p_idx, para in enumerate(section.paragraphs):
                text = para.text
                offset = 0
                while True:
                    pos = text.find(query, offset)
                    if pos == -1:
                        break
                    results.append((s_idx, p_idx, pos))
                    offset = pos + 1
        return results

    def replace_text(self, find: str, replace: str, replace_all: bool = False) -> int:
        """Replace text in document."""
        count = 0
        for section in self.sections:
            for para in section.paragraphs:
                for run in para.runs:
                    if find in run.text:
                        if replace_all:
                            run.text = run.text.replace(find, replace)
                            count += run.text.count(replace)
                        else:
                            run.text = run.text.replace(find, replace, 1)
                            count += 1
                            if not replace_all:
                                self._modified = True
                                return count
        if count > 0:
            self._modified = True
        return count

    def move_cursor(self, position: str) -> bool:
        """Move cursor to position."""
        if position == "doc_begin":
            self._current_section = 0
            self._current_para = 0
        elif position == "doc_end":
            self._current_section = len(self.sections) - 1
            self._current_para = len(self.current_section.paragraphs) - 1
        elif position == "next_para":
            section = self.current_section
            if self._current_para < len(section.paragraphs) - 1:
                self._current_para += 1
            elif self._current_section < len(self.sections) - 1:
                self._current_section += 1
                self._current_para = 0
        elif position == "prev_para":
            if self._current_para > 0:
                self._current_para -= 1
            elif self._current_section > 0:
                self._current_section -= 1
                self._current_para = len(self.current_section.paragraphs) - 1
        else:
            return False
        return True

    def set_paragraph_alignment(self, align: str) -> bool:
        """Set current paragraph alignment."""
        valid = {"left", "center", "right", "justify", "distribute"}
        if align.lower() in valid:
            self.current_paragraph.alignment = align.lower()
            self._modified = True
            return True
        return False

    def clear(self) -> None:
        """Clear document content."""
        self.sections = [Section(id="sec_0")]
        self.sections[0].paragraphs = [Paragraph(id="para_0")]
        self._current_section = 0
        self._current_para = 0
        self._modified = True
