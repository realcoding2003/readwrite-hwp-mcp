"""HWP MCP Server - Main entry point with hybrid backend support."""

import asyncio
import logging
import os
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server

from .backends import create_backend, get_available_backends, BaseBackend

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("hwp-mcp")

# Global backend instance
_backend: BaseBackend | None = None


def get_backend() -> BaseBackend:
    """Get or create the backend instance."""
    global _backend
    if _backend is None:
        # Check for explicit backend preference
        preferred = os.environ.get("HWP_BACKEND", "").lower()
        _backend = create_backend(preferred if preferred else None)
        logger.info(f"Using backend: {_backend.name}")
    return _backend


def create_server() -> Server:
    """Create and configure the MCP server."""
    server = Server("hwp-mcp")
    backend = get_backend()

    # Register tools
    _register_info_tools(server, backend)
    _register_document_tools(server, backend)
    _register_text_tools(server, backend)
    _register_table_tools(server, backend)
    _register_style_tools(server, backend)

    logger.info(f"Registered MCP tools (backend: {backend.name})")
    return server


def _register_info_tools(server: Server, backend: BaseBackend) -> None:
    """Register information tools."""

    @server.tool()
    async def hwp_status() -> dict:
        """
        현재 HWP MCP 서버 상태를 확인합니다.

        Returns:
            백엔드 정보, 사용 가능한 백엔드 목록, 연결 상태
        """
        return {
            "backend": backend.name,
            "is_connected": backend.is_connected,
            "supported_formats": backend.supported_formats,
            "available_backends": get_available_backends(),
        }


def _register_document_tools(server: Server, backend: BaseBackend) -> None:
    """Register document management tools."""

    @server.tool()
    async def hwp_connect(visible: bool = True) -> dict:
        """
        한글 프로그램/백엔드에 연결합니다.

        Args:
            visible: 한글 창을 표시할지 여부 (COM 백엔드만 해당)

        Returns:
            연결 결과
        """
        try:
            result = backend.connect(visible=visible)
            return {
                "success": result,
                "message": f"Connected to {backend.name} backend",
                "backend": backend.name,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_disconnect() -> dict:
        """백엔드 연결을 해제합니다."""
        try:
            result = backend.disconnect()
            return {"success": result, "message": "Disconnected"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_create() -> dict:
        """새 문서를 생성합니다."""
        try:
            result = backend.create_document()
            return {"success": result, "message": "New document created"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_open(path: str) -> dict:
        """
        기존 문서를 엽니다.

        Args:
            path: 문서 파일 경로 (.hwp 또는 .hwpx)
        """
        try:
            result = backend.open_document(path)
            return {"success": result, "message": f"Document opened: {path}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_save() -> dict:
        """현재 문서를 저장합니다."""
        try:
            result = backend.save_document()
            return {"success": result, "message": "Document saved"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_save_as(path: str, format: str = "HWPX") -> dict:
        """
        문서를 다른 이름/형식으로 저장합니다.

        Args:
            path: 저장할 파일 경로
            format: 파일 형식 (HWPX, HWP, PDF 등 - 백엔드에 따라 다름)
        """
        try:
            result = backend.save_document_as(path, format)
            return {"success": result, "message": f"Document saved as: {path}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_close() -> dict:
        """현재 문서를 닫습니다."""
        try:
            result = backend.close_document()
            return {"success": result, "message": "Document closed"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_get_info() -> dict:
        """현재 문서 정보를 가져옵니다."""
        try:
            info = backend.get_document_info()
            return {"success": True, "info": info}
        except Exception as e:
            return {"success": False, "error": str(e)}


def _register_text_tools(server: Server, backend: BaseBackend) -> None:
    """Register text editing tools."""

    @server.tool()
    async def hwp_insert_text(text: str) -> dict:
        """
        현재 위치에 텍스트를 삽입합니다.

        Args:
            text: 삽입할 텍스트
        """
        try:
            result = backend.insert_text(text)
            return {"success": result, "message": f"Text inserted: {text[:50]}..."}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_get_text() -> dict:
        """문서 전체 텍스트를 가져옵니다."""
        try:
            text = backend.get_text()
            return {"success": True, "text": text}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_find_text(text: str) -> dict:
        """
        문서에서 텍스트를 검색합니다.

        Args:
            text: 검색할 텍스트
        """
        try:
            found = backend.find_text(text)
            return {"success": True, "found": found}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_replace_text(find: str, replace: str, replace_all: bool = False) -> dict:
        """
        텍스트를 치환합니다.

        Args:
            find: 찾을 텍스트
            replace: 바꿀 텍스트
            replace_all: 모두 바꾸기 여부
        """
        try:
            count = backend.replace_text(find, replace, replace_all)
            return {"success": True, "count": count}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_insert_paragraph() -> dict:
        """새 단락을 삽입합니다."""
        try:
            result = backend.insert_paragraph()
            return {"success": result, "message": "Paragraph inserted"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_move_cursor(position: str) -> dict:
        """
        커서를 이동합니다.

        Args:
            position: 위치 (doc_begin, doc_end, next_para, prev_para 등)
        """
        try:
            result = backend.move_cursor(position)
            return {"success": result, "message": f"Cursor moved to: {position}"}
        except Exception as e:
            return {"success": False, "error": str(e)}


def _register_table_tools(server: Server, backend: BaseBackend) -> None:
    """Register table manipulation tools."""

    @server.tool()
    async def hwp_create_table(rows: int, cols: int) -> dict:
        """
        표를 생성합니다.

        Args:
            rows: 행 수
            cols: 열 수
        """
        try:
            result = backend.create_table(rows, cols)
            return {"success": result, "message": f"Table created: {rows}x{cols}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_get_cell(row: int, col: int) -> dict:
        """
        표 셀 내용을 가져옵니다.

        Args:
            row: 행 인덱스 (0부터)
            col: 열 인덱스 (0부터)
        """
        try:
            text = backend.get_cell_text(row, col)
            return {"success": True, "row": row, "col": col, "text": text}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_set_cell(row: int, col: int, text: str) -> dict:
        """
        표 셀에 텍스트를 설정합니다.

        Args:
            row: 행 인덱스 (0부터)
            col: 열 인덱스 (0부터)
            text: 설정할 텍스트
        """
        try:
            result = backend.set_cell_text(row, col, text)
            return {"success": result, "message": f"Cell ({row},{col}) set"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_insert_row() -> dict:
        """표에 행을 추가합니다."""
        try:
            result = backend.insert_row()
            return {"success": result, "message": "Row inserted"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_delete_row() -> dict:
        """표에서 행을 삭제합니다."""
        try:
            result = backend.delete_row()
            return {"success": result, "message": "Row deleted"}
        except Exception as e:
            return {"success": False, "error": str(e)}


def _register_style_tools(server: Server, backend: BaseBackend) -> None:
    """Register style and formatting tools."""

    @server.tool()
    async def hwp_set_font(
        name: str | None = None,
        size: int | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
    ) -> dict:
        """
        글꼴을 설정합니다.

        Args:
            name: 글꼴 이름
            size: 글꼴 크기 (pt)
            bold: 굵게
            italic: 기울임
        """
        try:
            result = backend.set_font(name=name, size=size, bold=bold, italic=italic)
            return {"success": result, "message": "Font settings applied"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_set_align(align: str) -> dict:
        """
        단락 정렬을 설정합니다.

        Args:
            align: 정렬 방식 (left, center, right, justify, distribute)
        """
        try:
            result = backend.set_alignment(align)
            return {"success": result, "message": f"Alignment set to: {align}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_export_pdf(path: str) -> dict:
        """
        문서를 PDF로 내보냅니다.

        Args:
            path: PDF 파일 저장 경로

        Note:
            HWPX 백엔드에서는 지원되지 않습니다.
        """
        try:
            result = backend.export_pdf(path)
            return {"success": result, "message": f"Exported to PDF: {path}"}
        except NotImplementedError as e:
            return {"success": False, "error": str(e), "hint": "Use COM backend for PDF export"}
        except Exception as e:
            return {"success": False, "error": str(e)}


async def main() -> None:
    """Main entry point for the MCP server."""
    server = create_server()

    async with stdio_server() as (read_stream, write_stream):
        logger.info("HWP MCP Server running on stdio")
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


if __name__ == "__main__":
    asyncio.run(main())
