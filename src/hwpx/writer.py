"""HWPX Writer - Write HWPX files."""

from __future__ import annotations

import os
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import BinaryIO

from .document import HwpxDocument, Section, Paragraph, Table


# HWPX XML namespaces
NS_HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS_HS = "http://www.hancom.co.kr/hwpml/2011/section"
NS_HC = "http://www.hancom.co.kr/hwpml/2011/core"
NS_HH = "http://www.hancom.co.kr/hwpml/2011/head"
NS_OPF = "http://www.idpf.org/2007/opf"
NS_DC = "http://purl.org/dc/elements/1.1/"


def _register_namespaces() -> None:
    """Register XML namespaces."""
    ET.register_namespace("hp", NS_HP)
    ET.register_namespace("hs", NS_HS)
    ET.register_namespace("hc", NS_HC)
    ET.register_namespace("hh", NS_HH)
    ET.register_namespace("opf", NS_OPF)
    ET.register_namespace("dc", NS_DC)


class HwpxWriter:
    """Writer for HWPX files."""

    @staticmethod
    def write(doc: HwpxDocument, path: str) -> bool:
        """
        Write document to HWPX file.

        Args:
            doc: HwpxDocument to write
            path: Output file path

        Returns:
            True if successful
        """
        _register_namespaces()

        # Use atomic write with temp file
        temp_path = f"{path}.tmp.{os.getpid()}"
        backup_path = f"{path}.backup"

        try:
            # Write to temp file
            HwpxWriter._write_hwpx(doc, temp_path)

            # Validate
            if not zipfile.is_zipfile(temp_path):
                raise ValueError("Failed to create valid HWPX file")

            # Backup existing file
            if os.path.exists(path):
                shutil.copy2(path, backup_path)

            # Move temp to final
            shutil.move(temp_path, path)

            # Remove backup
            if os.path.exists(backup_path):
                os.remove(backup_path)

            doc._path = path
            doc._modified = False
            return True

        except Exception as e:
            # Cleanup temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)

            # Restore backup
            if os.path.exists(backup_path):
                shutil.move(backup_path, path)

            raise

    @staticmethod
    def _write_hwpx(doc: HwpxDocument, path: str) -> None:
        """Write HWPX file content."""
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            # Write mimetype (uncompressed, first file)
            zf.writestr("mimetype", "application/hwp+zip", compress_type=zipfile.ZIP_STORED)

            # Write version.xml
            zf.writestr("version.xml", HwpxWriter._create_version_xml())

            # Write META-INF/container.xml
            zf.writestr("META-INF/container.xml", HwpxWriter._create_container_xml())

            # Write Contents/content.hpf
            zf.writestr("Contents/content.hpf", HwpxWriter._create_content_hpf(doc))

            # Write Contents/header.xml
            zf.writestr("Contents/header.xml", HwpxWriter._create_header_xml(doc))

            # Write section files
            for i, section in enumerate(doc.sections):
                section_xml = HwpxWriter._create_section_xml(section, i)
                zf.writestr(f"Contents/section{i}.xml", section_xml)

            # Write settings.xml
            zf.writestr("settings.xml", HwpxWriter._create_settings_xml())

    @staticmethod
    def _create_version_xml() -> str:
        """Create version.xml content."""
        root = ET.Element("hh:head", {
            "xmlns:hh": NS_HH,
            "version": "1.1",
        })
        return ET.tostring(root, encoding="unicode", xml_declaration=True)

    @staticmethod
    def _create_container_xml() -> str:
        """Create META-INF/container.xml content."""
        root = ET.Element("container", {
            "version": "1.0",
        })
        rootfiles = ET.SubElement(root, "rootfiles")
        ET.SubElement(rootfiles, "rootfile", {
            "full-path": "Contents/content.hpf",
            "media-type": "application/oebps-package+xml",
        })
        return ET.tostring(root, encoding="unicode", xml_declaration=True)

    @staticmethod
    def _create_content_hpf(doc: HwpxDocument) -> str:
        """Create Contents/content.hpf content."""
        root = ET.Element("opf:package", {
            "xmlns:opf": NS_OPF,
            "xmlns:dc": NS_DC,
            "version": "1.0",
        })

        # Metadata
        metadata = ET.SubElement(root, "opf:metadata")
        title = ET.SubElement(metadata, f"{{{NS_DC}}}title")
        title.text = doc.metadata.title or "Untitled"

        creator = ET.SubElement(metadata, f"{{{NS_DC}}}creator")
        creator.text = doc.metadata.author or "HWP MCP Server"

        date = ET.SubElement(metadata, f"{{{NS_DC}}}date")
        date.text = datetime.now().isoformat()

        # Manifest
        manifest = ET.SubElement(root, "opf:manifest")
        ET.SubElement(manifest, "opf:item", {
            "id": "header",
            "href": "header.xml",
            "media-type": "application/xml",
        })

        for i in range(len(doc.sections)):
            ET.SubElement(manifest, "opf:item", {
                "id": f"section{i}",
                "href": f"section{i}.xml",
                "media-type": "application/xml",
            })

        # Spine
        spine = ET.SubElement(root, "opf:spine")
        for i in range(len(doc.sections)):
            ET.SubElement(spine, "opf:itemref", {"idref": f"section{i}"})

        return ET.tostring(root, encoding="unicode", xml_declaration=True)

    @staticmethod
    def _create_header_xml(doc: HwpxDocument) -> str:
        """Create Contents/header.xml content."""
        root = ET.Element("hh:head", {"xmlns:hh": NS_HH})

        # Basic document settings
        doc_setting = ET.SubElement(root, "hh:docSetting")
        ET.SubElement(doc_setting, "hh:defaultLang").text = "ko"

        return ET.tostring(root, encoding="unicode", xml_declaration=True)

    @staticmethod
    def _create_section_xml(section: Section, index: int) -> str:
        """Create section XML content."""
        root = ET.Element("hs:sec", {
            "xmlns:hs": NS_HS,
            "xmlns:hp": NS_HP,
            "id": section.id,
        })

        # Write paragraphs
        for para in section.paragraphs:
            p_elem = ET.SubElement(root, f"{{{NS_HP}}}p", {"id": para.id})

            if para.para_style_id:
                p_elem.set("paraPrIDRef", para.para_style_id)

            # Write text runs
            for run in para.runs:
                if run.text:
                    run_elem = ET.SubElement(p_elem, f"{{{NS_HP}}}run")
                    if run.char_style_id:
                        run_elem.set("charPrIDRef", run.char_style_id)

                    t_elem = ET.SubElement(run_elem, f"{{{NS_HP}}}t")
                    t_elem.text = run.text

        # Write tables
        for table in section.tables:
            tbl_elem = ET.SubElement(root, f"{{{NS_HP}}}tbl", {"id": table.id})

            for row_idx, row in enumerate(table.cells):
                tr_elem = ET.SubElement(tbl_elem, f"{{{NS_HP}}}tr")

                for cell in row:
                    tc_elem = ET.SubElement(tr_elem, f"{{{NS_HP}}}tc")

                    # Write cell paragraphs
                    for para in cell.paragraphs:
                        p_elem = ET.SubElement(tc_elem, f"{{{NS_HP}}}p", {"id": para.id})
                        for run in para.runs:
                            if run.text:
                                run_elem = ET.SubElement(p_elem, f"{{{NS_HP}}}run")
                                t_elem = ET.SubElement(run_elem, f"{{{NS_HP}}}t")
                                t_elem.text = run.text

        return ET.tostring(root, encoding="unicode", xml_declaration=True)

    @staticmethod
    def _create_settings_xml() -> str:
        """Create settings.xml content."""
        root = ET.Element("settings")
        cursor = ET.SubElement(root, "cursor")
        cursor.set("section", "0")
        cursor.set("paragraph", "0")
        return ET.tostring(root, encoding="unicode", xml_declaration=True)
