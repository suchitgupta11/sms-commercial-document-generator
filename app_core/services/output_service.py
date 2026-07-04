from __future__ import annotations

from pathlib import Path

from app_core.config.settings import OUTPUT_DIR

DOCUMENT_OUTPUT_FOLDERS = {
    "quotation": "quotations",
    "invoice": "invoices",
    "proforma": "proforma_invoices",
    "challan": "challans",
    "salary_slip": "salary_slips",
}


def ensure_output_folder(document_key: str) -> Path:
    folder_name = DOCUMENT_OUTPUT_FOLDERS.get(document_key, document_key)
    path = OUTPUT_DIR / folder_name
    path.mkdir(parents=True, exist_ok=True)
    return path
