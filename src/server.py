"""HWP MCP Server - Main entry point."""

import asyncio
import logging
from contextlib import asynccontextmanager
from typing import AsyncIterator

from mcp.server import Server
from mcp.server.stdio import stdio_server

from .hwp import HwpController
from .tools import (
    register_document_tools,
    register_text_tools,
    register_table_tools,
    register_style_tools,
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("hwp-mcp")

# Global controller instance
_controller: HwpController | None = None


def get_controller() -> HwpController:
    """Get or create the HWP controller instance."""
    global _controller
    if _controller is None:
        _controller = HwpController()
    return _controller


@asynccontextmanager
async def server_lifespan(server: Server) -> AsyncIterator[dict]:
    """
    Server lifespan context manager.

    Handles startup and shutdown of the HWP controller.
    """
    controller = get_controller()
    logger.info("HWP MCP Server starting...")

    try:
        yield {"controller": controller}
    finally:
        # Cleanup on shutdown
        logger.info("HWP MCP Server shutting down...")
        if controller.is_connected:
            controller.disconnect()


def create_server() -> Server:
    """Create and configure the MCP server."""
    server = Server("hwp-mcp")
    controller = get_controller()

    # Register all tools
    register_document_tools(server, controller)
    register_text_tools(server, controller)
    register_table_tools(server, controller)
    register_style_tools(server, controller)

    logger.info("All MCP tools registered")
    return server


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
