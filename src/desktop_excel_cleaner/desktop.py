"""Helpers for locating Windows desktop files."""

from __future__ import annotations

import os
from pathlib import Path

from .config import SUPPORTED_EXCEL_EXTENSIONS


def get_desktop_path() -> Path:
    """Return the most likely desktop path for the current Windows user.

    Windows 11 systems may redirect Desktop to OneDrive. This function checks
    common OneDrive locations before falling back to the normal user Desktop.
    It also works on non-Windows systems, which makes automated tests portable.
    """

    home = Path.home()
    candidates = [
        Path(os.environ["USERPROFILE"]) / "OneDrive" / "Desktop"
        if os.environ.get("USERPROFILE")
        else None,
        Path(os.environ["USERPROFILE"]) / "Desktop"
        if os.environ.get("USERPROFILE")
        else None,
        home / "OneDrive" / "Desktop",
        home / "Desktop",
    ]

    for candidate in candidates:
        if candidate and candidate.exists():
            return candidate

    return home / "Desktop"


def list_excel_files(desktop_path: Path | None = None) -> list[Path]:
    """Return Excel files found on the desktop, sorted by file name."""

    desktop = desktop_path or get_desktop_path()
    if not desktop.exists():
        return []

    return sorted(
        path
        for path in desktop.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXCEL_EXTENSIONS
    )


def resolve_desktop_file(path_text: str, desktop_path: Path | None = None) -> Path:
    """Resolve a user-provided file name or path.

    Bare file names are interpreted as files on the desktop. Absolute paths and
    relative paths containing folders are resolved as provided.
    """

    path = Path(path_text).expanduser()
    if path.is_absolute() or path.parent != Path("."):
        return path
    return (desktop_path or get_desktop_path()) / path
