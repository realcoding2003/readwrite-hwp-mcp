"""Style and formatting MCP tools."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from mcp.server import Server
    from ..hwp.controller import HwpController


def register_style_tools(server: Server, controller: HwpController) -> None:
    """Register style and formatting tools with MCP server."""

    @server.tool()
    async def hwp_set_font(
        name: str | None = None,
        size: int | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
    ) -> dict:
        """
        글꼴 속성을 설정합니다.

        Args:
            name: 글꼴 이름 (예: "맑은 고딕", "나눔고딕")
            size: 글꼴 크기 (pt)
            bold: 굵게 여부
            italic: 기울임 여부

        Returns:
            설정 결과
        """
        try:
            result = controller.set_font(
                name=name, size=size, bold=bold, italic=italic
            )
            return {"success": result, "message": "Font settings applied"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_set_align(align: str) -> dict:
        """
        단락 정렬을 설정합니다.

        Args:
            align: 정렬 방식 (left, center, right, justify, distribute)

        Returns:
            설정 결과
        """
        try:
            result = controller.set_alignment(align)
            return {"success": result, "message": f"Alignment set to: {align}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_export_pdf(path: str) -> dict:
        """
        문서를 PDF로 내보냅니다.

        Args:
            path: PDF 파일 저장 경로

        Returns:
            내보내기 결과
        """
        try:
            result = controller.export_pdf(path)
            return {"success": result, "message": f"Exported to PDF: {path}"}
        except Exception as e:
            return {"success": False, "error": str(e)}
