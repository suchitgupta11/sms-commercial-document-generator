from __future__ import annotations

from pathlib import Path

from app_core.utils.platform import convert_docx_to_pdf


def generate_pdf_from_docx(docx_path: str | Path, output_dir: str | Path | None = None) -> Path:
    """Production PDF service wrapper used by UI/services.

    The implementation remains LibreOffice/Pages-backed today, but this layer
    allows replacing the PDF engine later without touching document generators.
    """
    return Path(convert_docx_to_pdf(Path(docx_path), output_dir=output_dir))
