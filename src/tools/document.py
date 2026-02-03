"""Document management MCP tools."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from mcp.server import Server
    from ..hwp.controller import HwpController


def register_document_tools(server: Server, controller: HwpController) -> None:
    """Register document management tools with MCP server."""

    @server.tool()
    async def hwp_connect(visible: bool = True) -> dict:
        """
        한글 프로그램에 연결합니다.

        Args:
            visible: 한글 창을 표시할지 여부 (기본값: True)

        Returns:
            연결 결과
        """
        try:
            result = controller.connect(visible=visible)
            return {"success": result, "message": "HWP connected successfully"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_disconnect() -> dict:
        """
        한글 프로그램 연결을 해제합니다.

        Returns:
            연결 해제 결과
        """
        try:
            result = controller.disconnect()
            return {"success": result, "message": "HWP disconnected"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_create() -> dict:
        """
        새 한글 문서를 생성합니다.

        Returns:
            생성 결과
        """
        try:
            result = controller.create_document()
            return {"success": result, "message": "New document created"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_open(path: str) -> dict:
        """
        기존 한글 문서를 엽니다.

        Args:
            path: 문서 파일 경로

        Returns:
            열기 결과
        """
        try:
            result = controller.open_document(path)
            return {"success": result, "message": f"Document opened: {path}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_save() -> dict:
        """
        현재 문서를 저장합니다.

        Returns:
            저장 결과
        """
        try:
            result = controller.save_document()
            return {"success": result, "message": "Document saved"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_save_as(path: str, format: str = "HWP") -> dict:
        """
        현재 문서를 다른 이름/형식으로 저장합니다.

        Args:
            path: 저장할 파일 경로
            format: 파일 형식 (HWP, HWPX, PDF, TEXT 등)

        Returns:
            저장 결과
        """
        try:
            result = controller.save_document_as(path, format)
            return {"success": result, "message": f"Document saved as: {path}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_close() -> dict:
        """
        현재 문서를 닫습니다.

        Returns:
            닫기 결과
        """
        try:
            result = controller.close_document()
            return {"success": result, "message": "Document closed"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @server.tool()
    async def hwp_get_info() -> dict:
        """
        현재 문서 정보를 가져옵니다.

        Returns:
            문서 정보 (경로, 파일명, 수정 여부)
        """
        try:
            info = controller.get_document_info()
            return {"success": True, "info": info}
        except Exception as e:
            return {"success": False, "error": str(e)}
