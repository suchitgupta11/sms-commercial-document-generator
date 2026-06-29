"""Proforma Invoice generator for Commercial Document Generator Pro V5.4.1 FINAL.

This module intentionally reuses the Tax Invoice engine so the PI layout remains
consistent with Invoice. Only document labels, footer text, and output filename
suffix change through the invoice engine's document_kind='proforma' mode.
"""
from pathlib import Path

try:
    from app_core.engines.modules.invoice.generate_invoice import generate_invoice, generate_invoice_docx
except Exception:  # direct script execution fallback
    import sys
    CURRENT_DIR = Path(__file__).resolve().parent
    INVOICE_DIR = CURRENT_DIR.parent / "invoice"
    if str(INVOICE_DIR) not in sys.path:
        sys.path.insert(0, str(INVOICE_DIR))
    from generate_invoice import generate_invoice, generate_invoice_docx  # noqa: E402


def generate_proforma_invoice(xlsx_path, output_dir=None, mode="both"):
    outputs = []
    if mode in ("pdf", "both"):
        outputs.append(generate_invoice(xlsx_path, output_dir, document_kind="proforma"))
    if mode in ("docx", "both"):
        outputs.append(generate_invoice_docx(xlsx_path, output_dir, document_kind="proforma"))
    return outputs


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_proforma_invoice.py <template.xlsx> [output_folder] [--both|--docx|--pdf]")
        sys.exit(1)

    xlsx = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 and not sys.argv[2].startswith("--") else None
    mode = "both"
    for arg in sys.argv[2:]:
        if arg in ("--both", "--docx", "--pdf"):
            mode = arg.replace("--", "")

    for p in generate_proforma_invoice(xlsx, out, mode=mode):
        print(p)
