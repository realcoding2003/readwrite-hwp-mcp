"""HWPX Reader - Parse HWPX files."""

from __future__ import annotations

import os
import zipfile
import xml.etree.ElementTree as ET
from typing import Any

from .document import (
    HwpxDocument,
    DocumentMetadata,
    Section,
    Paragraph,
    TextRun,
    Table,
    TableCell,
)


# HWPX XML namespaces
NAMESPACES = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "opf": "http://www.idpf.org/2007/opf",
    "dc": "http://purl.org/dc/elements/1.1/",
}


class HwpxReader:
    """Reader for HWPX files."""

    @staticmethod
    def read(path: str) -> HwpxDocument:
        """
        Read HWPX file and return document object.

        Args:
            path: Path to HWPX file

        Returns:
            HwpxDocument object

        Raises:
            FileNotFoundError: If file doesn't exist
            ValueError: If file is not a valid HWPX
        """
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")

        if not zipfile.is_zipfile(path):
            raise ValueError(f"Not a valid HWPX file: {path}")

        doc = HwpxDocument()
        doc._path = path
        doc.sections = []

        with zipfile.ZipFile(path, "r") as zf:
            # Read content manifest
            content_files = HwpxReader._get_content_files(zf)

            # Read metadata if available
            if "Contents/content.hpf" in zf.namelist():
                doc.metadata = HwpxReader._read_metadata(zf)

            # Read sections
            for section_file in content_files:
                if section_file in zf.namelist():
                    section = HwpxReader._read_section(zf, section_file)
                    doc.sections.append(section)

        # Ensure at least one section
        if not doc.sections:
            doc.sections.append(Section(id="sec_0"))
            doc.sections[0].paragraphs.append(Paragraph(id="para_0"))

        doc._modified = False
        return doc

    @staticmethod
    def _get_content_files(zf: zipfile.ZipFile) -> list[str]:
        """Get list of section files from manifest."""
        files = []

        # Try to read from content.hpf
        if "Contents/content.hpf" in zf.namelist():
            try:
                with zf.open("Contents/content.hpf") as f:
                    tree = ET.parse(f)
                    root = tree.getroot()

                    # Find all section references
                    for item in root.iter():
                        href = item.get("href", "")
                        if "section" in href.lower() and href.endswith(".xml"):
                            # Normalize path
                            if not href.startswith("Contents/"):
                                href = f"Contents/{href}"
                            files.append(href)
            except Exception:
                pass

        # Fallback: find section files directly
        if not files:
            for name in sorted(zf.namelist()):
                if "section" in name.lower() and name.endswith(".xml"):
                    files.append(name)

        return files

    @staticmethod
    def _read_metadata(zf: zipfile.ZipFile) -> DocumentMetadata:
        """Read document metadata from content.hpf."""
        metadata = DocumentMetadata()

        try:
            with zf.open("Contents/content.hpf") as f:
                tree = ET.parse(f)
                root = tree.getroot()

                # Find metadata section
                for meta in root.iter():
                    if "metadata" in meta.tag.lower():
                        for child in meta:
                            tag = child.tag.split("}")[-1].lower()
                            text = child.text or ""

                            if "title" in tag:
                                metadata.title = text
                            elif "creator" in tag or "author" in tag:
                                metadata.author = text
                            elif "subject" in tag:
                                metadata.subject = text
                            elif "date" in tag:
                                if "created" in child.attrib.get("event", ""):
                                    metadata.created = text
                                else:
                                    metadata.modified = text
        except Exception:
            pass

        return metadata

    @staticmethod
    def _read_section(zf: zipfile.ZipFile, section_file: str) -> Section:
        """Read a section file."""
        section_id = os.path.splitext(os.path.basename(section_file))[0]
        section = Section(id=section_id)

        try:
            with zf.open(section_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()

                para_count = 0
                table_count = 0

                # Process all elements
                for elem in root.iter():
                    tag = elem.tag.split("}")[-1]

                    # Paragraph (hp:p)
                    if tag == "p":
                        para = HwpxReader._parse_paragraph(elem, para_count)
                        section.paragraphs.append(para)
                        para_count += 1

                    # Table (hp:tbl)
                    elif tag == "tbl":
                        table = HwpxReader._parse_table(elem, table_count)
                        if table:
                            section.tables.append(table)
                            table_count += 1

        except Exception as e:
            # Return empty section on error
            pass

        # Ensure at least one paragraph
        if not section.paragraphs:
            section.paragraphs.append(Paragraph(id=f"{section_id}_para_0"))

        return section

    @staticmethod
    def _parse_paragraph(elem: ET.Element, index: int) -> Paragraph:
        """Parse paragraph element."""
        para_id = elem.get("id", f"para_{index}")
        para = Paragraph(id=para_id)

        # Get paragraph style reference
        para_style = elem.get("paraPrIDRef", "")
        para.para_style_id = para_style

        # Find text runs
        for child in elem.iter():
            tag = child.tag.split("}")[-1]

            # Text run (hp:run)
            if tag == "run":
                run = HwpxReader._parse_run(child)
                if run:
                    para.runs.append(run)

            # Direct text (hp:t)
            elif tag == "t" and child.text:
                para.runs.append(TextRun(text=child.text))

        return para

    @staticmethod
    def _parse_run(elem: ET.Element) -> TextRun | None:
        """Parse text run element."""
        text_parts = []
        char_style = elem.get("charPrIDRef", "")

        for child in elem.iter():
            tag = child.tag.split("}")[-1]
            if tag == "t" and child.text:
                text_parts.append(child.text)

        if text_parts:
            return TextRun(
                text="".join(text_parts),
                char_style_id=char_style if char_style else None,
            )
        return None

    @staticmethod
    def _parse_table(elem: ET.Element, index: int) -> Table | None:
        """Parse table element."""
        try:
            # Get table dimensions
            rows = 0
            cols = 0

            # Count rows and columns
            for child in elem.iter():
                tag = child.tag.split("}")[-1]
                if tag == "tr":
                    rows += 1
                    col_count = sum(1 for c in child.iter() if c.tag.split("}")[-1] == "tc")
                    cols = max(cols, col_count)

            if rows == 0 or cols == 0:
                return None

            table = Table(id=f"table_{index}", rows=rows, cols=cols)

            # Parse cells
            row_idx = 0
            for child in elem.iter():
                tag = child.tag.split("}")[-1]
                if tag == "tr":
                    col_idx = 0
                    for cell_elem in child.iter():
                        cell_tag = cell_elem.tag.split("}")[-1]
                        if cell_tag == "tc" and col_idx < cols:
                            cell = table.get_cell(row_idx, col_idx)
                            if cell:
                                # Get cell text
                                for p_elem in cell_elem.iter():
                                    if p_elem.tag.split("}")[-1] == "p":
                                        para = HwpxReader._parse_paragraph(p_elem, 0)
                                        cell.paragraphs.append(para)
                            col_idx += 1
                    row_idx += 1

            return table

        except Exception:
            return None
