"""Table manipulation MCP tools."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from mcp.server import Server
    from ..hwp.controller import HwpController


def register_table_tools(server: Server, controller: HwpController) -> None:
    """Register table manipulation tools with MCP server."""

    @server.tool()
    async def hwp_create_table(rows: int, cols: int) -> dict:
        """
        현재 커서 위치에 표를 생성합니다.

        Args:
            rows: 행 수
            cols: 열 수

        Returns:
            생성 결과
        """
        try:
            result = controller.create_table(rows, cols)
            return {
                "success": result,
                "message": f"Table created: {rows}x{cols}",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_get_cell(row: int, col: int) -> dict:
        """
        표 셀의 내용을 가져옵니다.

        Args:
            row: 행 인덱스 (0부터 시작)
            col: 열 인덱스 (0부터 시작)

        Returns:
            셀 내용
        """
        try:
            text = controller.get_cell_text(row, col)
            return {
                "success": True,
                "row": row,
                "col": col,
                "text": text,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_set_cell(row: int, col: int, text: str) -> dict:
        """
        표 셀에 텍스트를 설정합니다.

        Args:
            row: 행 인덱스 (0부터 시작)
            col: 열 인덱스 (0부터 시작)
            text: 설정할 텍스트

        Returns:
            설정 결과
        """
        try:
            result = controller.set_cell_text(row, col, text)
            return {
                "success": result,
                "message": f"Cell ({row},{col}) set to: {text[:50]}",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_insert_row() -> dict:
        """
        현재 표에 행을 추가합니다.

        Returns:
            추가 결과
        """
        try:
            result = controller.insert_row()
            return {"success": result, "message": "Row inserted"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_delete_row() -> dict:
        """
        현재 행을 삭제합니다.

        Returns:
            삭제 결과
        """
        try:
            result = controller.delete_row()
            return {"success": result, "message": "Row deleted"}
        except Exception as e:
            return {"success": False, "error": str(e)}
