"""Text editing MCP tools."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from mcp.server import Server
    from ..hwp.controller import HwpController


def register_text_tools(server: Server, controller: HwpController) -> None:
    """Register text editing tools with MCP server."""

    @server.tool()
    async def hwp_insert_text(text: str) -> dict:
        """
        현재 커서 위치에 텍스트를 삽입합니다.

        Args:
            text: 삽입할 텍스트

        Returns:
            삽입 결과
        """
        try:
            result = controller.insert_text(text)
            return {"success": result, "message": f"Text inserted: {text[:50]}..."}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_get_text() -> dict:
        """
        문서 전체 텍스트를 가져옵니다.

        Returns:
            문서 텍스트 내용
        """
        try:
            text = controller.get_text()
            return {"success": True, "text": text}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_find_text(text: str) -> dict:
        """
        문서에서 텍스트를 검색합니다.

        Args:
            text: 검색할 텍스트

        Returns:
            검색 결과
        """
        try:
            found = controller.find_text(text)
            return {
                "success": True,
                "found": found,
                "message": f"Text {'found' if found else 'not found'}: {text}",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_replace_text(
        find: str, replace: str, replace_all: bool = False
    ) -> dict:
        """
        문서에서 텍스트를 치환합니다.

        Args:
            find: 찾을 텍스트
            replace: 바꿀 텍스트
            replace_all: 모두 바꾸기 여부 (기본값: False)

        Returns:
            치환 결과 및 치환 횟수
        """
        try:
            count = controller.replace_text(find, replace, replace_all)
            return {
                "success": True,
                "count": count,
                "message": f"Replaced {count} occurrence(s)",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_insert_paragraph() -> dict:
        """
        새 단락을 삽입합니다 (Enter 키).

        Returns:
            삽입 결과
        """
        try:
            result = controller.insert_paragraph()
            return {"success": result, "message": "Paragraph inserted"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_move_cursor(position: str) -> dict:
        """
        커서를 지정된 위치로 이동합니다.

        Args:
            position: 위치 (doc_begin, doc_end, line_begin, line_end,
                          para_begin, para_end, left, right, up, down,
                          next_para, prev_para)

        Returns:
            이동 결과
        """
        try:
            result = controller.move_cursor(position)
            return {"success": result, "message": f"Cursor moved to: {position}"}
        except Exception as e:
            return {"success": False, "error": str(e)}
