from __future__ import annotations

from pathlib import Path
from typing import List

from app_core.utils.platform import convert_docx_to_pdf
from app_core.services.registry_service import DocumentDefinition


class DocumentGenerationService:
    """Thin service wrapper around finalized document-generation engines."""

    def generate(self, document: DocumentDefinition, template_path: Path, output_dir: Path) -> List[Path]:
        output_dir.mkdir(parents=True, exist_ok=True)
        key = document.engine_key

        if key == "quotation":
            from app_core.engines.modules.quotation.generate_proposal import read_excel, clean, create_docx
            import re
            customer, quotation, *_ = read_excel(template_path)
            quote_no = clean(quotation.get("Quotation No.", "Quotation")).replace("/", "-").replace("\\", "-")
            cust = re.sub(r"[^A-Za-z0-9]+", "_", clean(customer.get("Customer Name", "Customer")).replace("M/s.", "")).strip("_")[:35]
            doc_type = re.sub(r"[^A-Za-z0-9]+", "_", clean(quotation.get("Document Type", "Commercial Proposal"))).strip("_")
            docx = output_dir / f"{quote_no}_{cust}_{doc_type}.docx"
            create_docx(template_path, docx)
            pdf = convert_docx_to_pdf(docx, output_dir=output_dir)
            return [docx, pdf]

        if key == "invoice":
            from app_core.engines.modules.invoice.generate_invoice import generate_invoice, generate_invoice_docx
            pdf = Path(generate_invoice(template_path, output_dir, document_kind="invoice"))
            docx = Path(generate_invoice_docx(template_path, output_dir, document_kind="invoice"))
            return [docx, pdf]

        if key == "proforma":
            from app_core.engines.modules.proforma.generate_proforma_invoice import generate_proforma_invoice
            return [Path(p) for p in generate_proforma_invoice(template_path, output_dir, mode="both")]

        if key == "challan":
            from app_core.engines.modules.challan.generate_challan import generate_challan
            return [Path(p) for p in generate_challan(template_path, output_dir)]

        if key == "purchase_order":
            from app_core.engines.modules.purchase_order.generate_purchase_order import generate_purchase_order
            return [Path(p) for p in generate_purchase_order(template_path, output_dir, mode="both")]

        if key == "salary_slip":
            from app_core.engines.modules.salary_slip.generate_salary_slip import generate
            return [Path(p) for p in generate(template_path, output_dir)]

        raise RuntimeError(f"Unsupported document engine: {key}")
