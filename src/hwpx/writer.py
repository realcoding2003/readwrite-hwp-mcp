"""HWPX Writer - Write HWPX files."""

from __future__ import annotations

import os
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import BinaryIO

from .document import HwpxDocument, Section, Paragraph, Table, Image


# HWPX XML namespaces (complete set matching real HWPX files)
NS_HA = "http://www.hancom.co.kr/hwpml/2011/app"
NS_HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS_HP10 = "http://www.hancom.co.kr/hwpml/2016/paragraph"
NS_HS = "http://www.hancom.co.kr/hwpml/2011/section"
NS_HC = "http://www.hancom.co.kr/hwpml/2011/core"
NS_HH = "http://www.hancom.co.kr/hwpml/2011/head"
NS_HHS = "http://www.hancom.co.kr/hwpml/2011/history"
NS_HM = "http://www.hancom.co.kr/hwpml/2011/master-page"
NS_HPF = "http://www.hancom.co.kr/schema/2011/hpf"
NS_HV = "http://www.hancom.co.kr/hwpml/2011/version"
NS_OPF = "http://www.idpf.org/2007/opf/"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_CONFIG = "urn:oasis:names:tc:opendocument:xmlns:config:1.0"
NS_OCF = "urn:oasis:names:tc:opendocument:xmlns:container"
NS_ODF = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
NS_OOXMLCHART = "http://www.hancom.co.kr/hwpml/2016/ooxmlchart"
NS_HWPUNITCHAR = "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"
NS_EPUB = "http://www.idpf.org/2007/ops"


def _register_namespaces() -> None:
    """Register XML namespaces to avoid ns0, ns1, etc. prefixes."""
    ET.register_namespace("ha", NS_HA)
    ET.register_namespace("hp", NS_HP)
    ET.register_namespace("hp10", NS_HP10)
    ET.register_namespace("hs", NS_HS)
    ET.register_namespace("hc", NS_HC)
    ET.register_namespace("hh", NS_HH)
    ET.register_namespace("hhs", NS_HHS)
    ET.register_namespace("hm", NS_HM)
    ET.register_namespace("hpf", NS_HPF)
    ET.register_namespace("hv", NS_HV)
    ET.register_namespace("opf", NS_OPF)
    ET.register_namespace("dc", NS_DC)
    ET.register_namespace("config", NS_CONFIG)
    ET.register_namespace("ocf", NS_OCF)
    ET.register_namespace("odf", NS_ODF)
    ET.register_namespace("ooxmlchart", NS_OOXMLCHART)
    ET.register_namespace("hwpunitchar", NS_HWPUNITCHAR)
    ET.register_namespace("epub", NS_EPUB)


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
        # Collect all images from all sections
        all_images: list[Image] = []
        for section in doc.sections:
            all_images.extend(section.images)

        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            # Write mimetype (uncompressed, first file)
            zf.writestr("mimetype", "application/hwp+zip", compress_type=zipfile.ZIP_STORED)

            # Write version.xml
            zf.writestr("version.xml", HwpxWriter._create_version_xml())

            # Write META-INF files
            zf.writestr("META-INF/container.xml", HwpxWriter._create_container_xml())
            zf.writestr("META-INF/manifest.xml", HwpxWriter._create_manifest_xml())

            # Write Contents/content.hpf (includes image references)
            zf.writestr("Contents/content.hpf", HwpxWriter._create_content_hpf(doc, all_images))

            # Write Contents/header.xml
            zf.writestr("Contents/header.xml", HwpxWriter._create_header_xml(doc))

            # Write section files
            for i, section in enumerate(doc.sections):
                section_xml = HwpxWriter._create_section_xml(section, i)
                zf.writestr(f"Contents/section{i}.xml", section_xml)

            # Write settings.xml
            zf.writestr("settings.xml", HwpxWriter._create_settings_xml())

            # Write image files to BinData/
            for image in all_images:
                if os.path.exists(image.file_path):
                    ext = os.path.splitext(image.file_path)[1].lower()
                    zf.write(image.file_path, f"BinData/{image.binary_id}{ext}")

    @staticmethod
    def _create_version_xml() -> str:
        """Create version.xml content."""
        xml_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="1" buildNumber="0" os="10" xmlVersion="1.5" application="Hancom Office Hangul" appVersion="12.0.0.1"/>'''
        return xml_content

    @staticmethod
    def _create_container_xml() -> str:
        """Create META-INF/container.xml content."""
        xml_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/></ocf:rootfiles></ocf:container>'''
        return xml_content

    @staticmethod
    def _create_manifest_xml() -> str:
        """Create META-INF/manifest.xml content."""
        xml_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'''
        return xml_content

    @staticmethod
    def _create_content_hpf(doc: HwpxDocument, images: list[Image] | None = None) -> str:
        """Create Contents/content.hpf content."""
        title = doc.metadata.title or "Untitled"
        creator = doc.metadata.author or "HWP MCP Server"
        date = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")

        # Build manifest items for images
        image_items = ""
        if images:
            for image in images:
                ext = os.path.splitext(image.file_path)[1].lower()
                image_items += f'<opf:item id="{image.binary_id}" href="BinData/{image.binary_id}{ext}" media-type="{image.media_type}" isEmbeded="1"/>'

        # Build manifest items for sections
        section_items = ""
        section_refs = ""
        for i in range(len(doc.sections)):
            section_items += f'<opf:item id="section{i}" href="Contents/section{i}.xml" media-type="application/xml"/>'
            section_refs += f'<opf:itemref idref="section{i}" linear="yes"/>'

        # Match real HWPX format with all namespaces
        xml_content = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><opf:package xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" version="" unique-identifier="" id=""><opf:metadata><opf:title xml:space="preserve">{title}</opf:title><opf:language>ko</opf:language><opf:meta name="creator" content="text">{creator}</opf:meta><opf:meta name="CreatedDate" content="text">{date}</opf:meta><opf:meta name="ModifiedDate" content="text">{date}</opf:meta></opf:metadata><opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>{image_items}{section_items}<opf:item id="settings" href="settings.xml" media-type="application/xml"/></opf:manifest><opf:spine><opf:itemref idref="header" linear="yes"/>{section_refs}</opf:spine></opf:package>'''
        return xml_content

    @staticmethod
    def _create_header_xml(doc: HwpxDocument) -> str:
        """Create Contents/header.xml content with required style references."""
        # All namespaces matching real HWPX format
        ns_decl = 'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'

        xml_content = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hh:head {ns_decl} version="1.5" secCnt="1"><hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/><hh:refList><hh:fontfaces itemCnt="7"><hh:fontface lang="HANGUL" fontCnt="1"><hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="0"/></hh:fontface><hh:fontface lang="LATIN" fontCnt="1"><hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="0"/></hh:fontface><hh:fontface lang="HANJA" fontCnt="1"><hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="0"/></hh:fontface><hh:fontface lang="JAPANESE" fontCnt="1"><hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="0"/></hh:fontface><hh:fontface lang="OTHER" fontCnt="1"><hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="0"/></hh:fontface><hh:fontface lang="SYMBOL" fontCnt="1"><hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="0"/></hh:fontface><hh:fontface lang="USER" fontCnt="1"><hh:font id="0" face="함초롬돋움" type="TTF" isEmbedded="0"/></hh:fontface></hh:fontfaces><hh:borderFills itemCnt="2"><hh:borderFill id="1" threeD="0" shadow="0" centerLine="0" breakCellSeparateLine="0"><hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/><hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/><hh:topBorder type="NONE" width="0.1 mm" color="#000000"/><hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/></hh:borderFill><hh:borderFill id="2" threeD="0" shadow="0" centerLine="0" breakCellSeparateLine="0"><hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/><hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/><hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/><hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/></hh:borderFill></hh:borderFills><hh:charProperties itemCnt="1"><hh:charPr id="0" height="1000" textColor="#000000" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1"><hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/><hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/><hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/></hh:charPr></hh:charProperties><hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties><hh:numberings itemCnt="1"><hh:numbering id="1" start="1"/></hh:numberings><hh:paraProperties itemCnt="1"><hh:paraPr id="0" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" checked="0"><hh:align horizontal="JUSTIFY" vertical="BASELINE"/><hh:heading type="NONE" idRef="0" level="0"/><hh:margin indent="0" left="0" right="0" prev="0" next="0"/><hh:lineSpacing type="PERCENT" value="160"/><hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0"/><hh:autoSpacing eAsianEng="0" eAsianNum="0"/></hh:paraPr></hh:paraProperties><hh:styles itemCnt="1"><hh:style id="0" type="PARA" name="바탕글" engName="Normal" paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langIDRef="0" lockForm="0"/></hh:styles><hh:memoProperties itemCnt="0"/></hh:refList><hh:compatibleDocument targetProgram="HWP201X"/><hh:docOption><hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/></hh:docOption><hh:trackchageConfig flags="0"/></hh:head>'''
        return xml_content

    @staticmethod
    def _create_section_xml(section: Section, index: int) -> str:
        """Create section XML content."""
        import random

        # All namespaces for section XML (matching real HWPX format)
        ns_decl = 'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'

        paragraphs_xml = []
        para_id = random.randint(1000000000, 9999999999)

        # Section properties for first paragraph
        sec_pr = '''<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="1" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="WIDELY" width="59528" height="84188" gutterType="LEFT_ONLY"><hp:margin header="2834" footer="2834" gutter="0" left="5669" right="5669" top="4251" bottom="4251"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr><hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill></hp:secPr>'''

        is_first = True

        # Write paragraphs
        for para in section.paragraphs:
            text_content = ""
            for run in para.runs:
                if run.text:
                    # Escape XML special characters
                    escaped_text = (run.text
                        .replace("&", "&amp;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;")
                        .replace('"', "&quot;"))
                    text_content += f'<hp:run charPrIDRef="0"><hp:t>{escaped_text}</hp:t></hp:run>'

            if is_first and index == 0:
                # First paragraph includes section properties
                p_xml = f'<hp:p id="{para_id}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">{sec_pr}</hp:run>{text_content}</hp:p>'
                is_first = False
            else:
                p_xml = f'<hp:p id="{para_id}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">{text_content}</hp:p>'

            paragraphs_xml.append(p_xml)
            para_id += 1

        # Write tables
        for table in section.tables:
            table_id = random.randint(2000000000, 2999999999)
            # Calculate table dimensions
            col_width = 15000  # Each column width
            row_height = 1000  # Each row height
            table_width = col_width * table.cols
            table_height = row_height * table.rows

            table_rows = []
            for row_idx, row in enumerate(table.cells):
                cells_xml = []
                for col_idx, cell in enumerate(row):
                    cell_text = ""
                    for cell_para in cell.paragraphs:
                        for run in cell_para.runs:
                            if run.text:
                                escaped = (run.text
                                    .replace("&", "&amp;")
                                    .replace("<", "&lt;")
                                    .replace(">", "&gt;"))
                                cell_text += escaped

                    # Match real HWPX tc structure with subList
                    cell_xml = f'<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="2"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0"><hp:p id="{para_id}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t>{cell_text}</hp:t></hp:run></hp:p></hp:subList><hp:cellAddr colAddr="{col_idx}" rowAddr="{row_idx}"/><hp:cellSpan colSpan="1" rowSpan="1"/><hp:cellSz width="{col_width}" height="{row_height}"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/></hp:tc>'
                    cells_xml.append(cell_xml)
                    para_id += 1

                table_rows.append(f'<hp:tr>{"".join(cells_xml)}</hp:tr>')

            # Table XML - match real HWPX tbl structure with sz, pos, margins
            tbl_xml = f'<hp:tbl id="{table_id}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="0" rowCnt="{table.rows}" colCnt="{table.cols}" cellSpacing="0" borderFillIDRef="2" noAdjust="0"><hp:sz width="{table_width}" widthRelTo="ABSOLUTE" height="{table_height}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="283" right="283" top="283" bottom="283"/><hp:inMargin left="510" right="510" top="141" bottom="141"/>{"".join(table_rows)}</hp:tbl>'
            tbl_p_xml = f'<hp:p id="{para_id}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">{tbl_xml}<hp:t/></hp:run></hp:p>'
            paragraphs_xml.append(tbl_p_xml)
            para_id += 1

        # Write images
        z_order = 1
        for image in section.images:
            pic_id = random.randint(2000000000, 2999999999)
            inst_id = random.randint(900000000, 999999999)

            # Use image dimensions
            width = image.width if image.width else 20000
            height = image.height if image.height else 20000
            org_width = image.original_width if image.original_width else width
            org_height = image.original_height if image.original_height else height

            # Calculate scale matrix values
            scale_x = width / org_width if org_width else 1.0
            scale_y = height / org_height if org_height else 1.0

            # Calculate center for rotation
            center_x = width // 2
            center_y = height // 2

            # Build pic XML matching real HWPX structure
            pic_xml = f'''<hp:pic id="{pic_id}" zOrder="{z_order}" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="{inst_id}" reverse="0"><hp:offset x="0" y="0"/><hp:orgSz width="{org_width}" height="{org_height}"/><hp:curSz width="{width}" height="{height}"/><hp:flip horizontal="0" vertical="0"/><hp:rotationInfo angle="0" centerX="{center_x}" centerY="{center_y}" rotateimage="1"/><hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:scaMatrix e1="{scale_x:.6f}" e2="0" e3="0" e4="0" e5="{scale_y:.6f}" e6="0"/><hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo><hc:img binaryItemIDRef="{image.binary_id}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/><hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="{org_width}" y="0"/><hc:pt2 x="{org_width}" y="{org_height}"/><hc:pt3 x="0" y="{org_height}"/></hp:imgRect><hp:imgClip left="0" right="0" top="0" bottom="0"/><hp:inMargin left="0" right="0" top="0" bottom="0"/><hp:imgDim dimwidth="{org_width}" dimheight="{org_height}"/><hp:effects/><hp:sz width="{width}" widthRelTo="ABSOLUTE" height="{height}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/></hp:pic>'''

            pic_p_xml = f'<hp:p id="{para_id}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">{pic_xml}<hp:t/></hp:run></hp:p>'
            paragraphs_xml.append(pic_p_xml)
            para_id += 1
            z_order += 1

        content = "".join(paragraphs_xml)
        return f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec {ns_decl}>{content}</hs:sec>'

    @staticmethod
    def _create_settings_xml() -> str:
        """Create settings.xml content."""
        xml_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/></ha:HWPApplicationSetting>'''
        return xml_content
