"""
thumbnails.py — dependency checking and conversion pipeline for thumbnail generation.

The thumbnail workflow requires three external dependencies:
  - soffice (LibreOffice) for PPTX → PDF conversion
  - pdftoppm (poppler-utils) for PDF → image rasterisation
  - Pillow for image post-processing

Call check_dependencies() early in any entry point that uses this module.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


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


def pptx_to_jpegs(pptx_path: Path | str, temp_dir: Path | str) -> list[Path]:
    """Convert a .pptx file to a list of per-page JPEG images.

    Runs a two-step subprocess pipeline:
      1. ``soffice --headless --convert-to pdf``  →  intermediate PDF in *temp_dir*
      2. ``pdftoppm -jpeg -r 150``                →  per-page JPEGs in *temp_dir*

    Args:
        pptx_path: Path to the source .pptx file.
        temp_dir:  Directory for intermediate and output files (managed by caller).

    Returns:
        An ordered list of :class:`~pathlib.Path` objects, one JPEG per slide,
        sorted by page number.

    Raises:
        RuntimeError: If either subprocess exits with a non-zero return code.
                      The message includes the captured stderr output.
    """
    pptx_path = Path(pptx_path)
    temp_dir = Path(temp_dir)

    # Step 1: .pptx → PDF via LibreOffice
    soffice_result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(temp_dir),
            str(pptx_path),
        ],
        capture_output=True,
        text=True,
    )
    if soffice_result.returncode != 0:
        raise RuntimeError(
            f"soffice failed (exit {soffice_result.returncode}):\n{soffice_result.stderr}"
        )

    pdf_path = temp_dir / (pptx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError(
            f"soffice did not produce expected PDF at {pdf_path}"
        )

    # Step 2: PDF → per-page JPEGs via pdftoppm
    jpeg_prefix = str(temp_dir / pptx_path.stem)
    pdftoppm_result = subprocess.run(
        [
            "pdftoppm",
            "-jpeg",
            "-r", "150",
            str(pdf_path),
            jpeg_prefix,
        ],
        capture_output=True,
        text=True,
    )
    if pdftoppm_result.returncode != 0:
        raise RuntimeError(
            f"pdftoppm failed (exit {pdftoppm_result.returncode}):\n{pdftoppm_result.stderr}"
        )

    # Collect output JPEGs; pdftoppm zero-pads page numbers so lexicographic
    # sort matches page order.
    jpegs = sorted(temp_dir.glob(f"{pptx_path.stem}-*.jpg"))
    return jpegs
