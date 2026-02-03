"""Backend factory - Auto-detect and create appropriate backend."""

from __future__ import annotations

import sys
from typing import TYPE_CHECKING

from .base import BaseBackend
from .hwpx_backend import HwpxBackend

if TYPE_CHECKING:
    from .com_backend import ComBackend


def get_available_backends() -> dict[str, bool]:
    """
    Get list of available backends and their status.

    Returns:
        Dictionary of backend name -> availability
    """
    backends = {
        "hwpx": HwpxBackend.is_available(),
    }

    # Check COM backend (Windows only)
    if sys.platform == "win32":
        try:
            from .com_backend import ComBackend

            backends["com"] = ComBackend.is_available()
        except ImportError:
            backends["com"] = False
    else:
        backends["com"] = False

    return backends


def create_backend(backend_type: str | None = None) -> BaseBackend:
    """
    Create a backend instance.

    Args:
        backend_type: Specific backend to use ('com', 'hwpx', or None for auto)

    Returns:
        Backend instance

    Raises:
        ValueError: If requested backend is not available
    """
    available = get_available_backends()

    # Specific backend requested
    if backend_type is not None:
        backend_type = backend_type.lower()

        if backend_type == "com":
            if not available.get("com"):
                raise ValueError(
                    "COM backend is not available. "
                    "Requires Windows with HWP installed."
                )
            from .com_backend import ComBackend

            return ComBackend()

        elif backend_type == "hwpx":
            return HwpxBackend()

        else:
            raise ValueError(f"Unknown backend type: {backend_type}")

    # Auto-detect best backend
    # Prefer COM on Windows (full feature support)
    if available.get("com"):
        from .com_backend import ComBackend

        return ComBackend()

    # Fallback to HWPX (always available)
    return HwpxBackend()


def get_recommended_backend() -> str:
    """
    Get recommended backend for current platform.

    Returns:
        Backend name ('com' or 'hwpx')
    """
    available = get_available_backends()

    if available.get("com"):
        return "com"
    return "hwpx"
