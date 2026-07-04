from __future__ import annotations
from pathlib import Path

from app_core.engines.modules.invoice.generate_invoice import generate_invoice, generate_invoice_docx


def generate_purchase_order(xlsx_path, output_dir=None, mode: str = "both"):
    """Generate Purchase Order DOCX/PDF using the premium invoice-style layout."""
    outputs = []
    if mode in ("docx", "both"):
        outputs.append(Path(generate_invoice_docx(xlsx_path, output_dir, document_kind="purchase_order")))
    if mode in ("pdf", "both"):
        outputs.append(Path(generate_invoice(xlsx_path, output_dir, document_kind="purchase_order")))
    return outputs


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python generate_purchase_order.py <template.xlsx> [output_folder] [--both|--docx|--pdf]")
        raise SystemExit(1)
    xlsx = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 and not sys.argv[2].startswith("--") else None
    mode = "both"
    for arg in sys.argv[2:]:
        if arg in ("--both", "--docx", "--pdf"):
            mode = arg.replace("--", "")
    for p in generate_purchase_order(xlsx, out, mode):
        print(p)
