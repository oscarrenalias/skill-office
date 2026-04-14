"""
thumbnails.py — dependency checking for the thumbnail generation pipeline.

The thumbnail workflow requires three external dependencies:
  - soffice (LibreOffice) for PPTX → PDF conversion
  - pdftoppm (poppler-utils) for PDF → image rasterisation
  - Pillow for image post-processing

Call check_dependencies() early in any entry point that uses this module.
"""

from __future__ import annotations

import shutil
import sys


def check_dependencies() -> None:
    """Verify that soffice, pdftoppm, and Pillow are available.

    Prints a clear, actionable install hint to stderr for each missing
    dependency, then raises SystemExit(1) if any are absent.
    """
    missing = False

    if shutil.which("soffice") is None:
        print(
            "Error: 'soffice' (LibreOffice) not found in PATH.\n"
            "  Install on macOS:  brew install --cask libreoffice\n"
            "  Install on Debian: sudo apt-get install libreoffice",
            file=sys.stderr,
        )
        missing = True

    if shutil.which("pdftoppm") is None:
        print(
            "Error: 'pdftoppm' (poppler-utils) not found in PATH.\n"
            "  Install on macOS:  brew install poppler\n"
            "  Install on Debian: sudo apt-get install poppler-utils",
            file=sys.stderr,
        )
        missing = True

    try:
        import PIL  # noqa: F401
    except ImportError:
        print(
            "Error: Pillow is not installed.\n"
            "  pip install 'pypptx[thumbnails]'",
            file=sys.stderr,
        )
        missing = True

    if missing:
        raise SystemExit(1)
