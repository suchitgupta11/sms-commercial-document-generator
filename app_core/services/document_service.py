from __future__ import annotations

from pathlib import Path

from app_core.config.settings import TEMPLATES_DIR, OUTPUT_DIR

DOCUMENT_REGISTRY = {
    "quotation": {
        "display_name": "Quotation",
        "template": TEMPLATES_DIR / "quotation" / "Proposal_Data_Template.xlsx",
        "output_dir": OUTPUT_DIR / "quotations",
    },
    "invoice": {
        "display_name": "Tax Invoice",
        "template": TEMPLATES_DIR / "invoice" / "SMS_Invoice_Data_Template.xlsx",
        "output_dir": OUTPUT_DIR / "invoices",
    },
    "proforma": {
        "display_name": "Proforma Invoice",
        "template": TEMPLATES_DIR / "proforma" / "SMS_Proforma_Invoice_Template.xlsx",
        "output_dir": OUTPUT_DIR / "proforma_invoices",
    },
    "challan": {
        "display_name": "Delivery Challan",
        "template": TEMPLATES_DIR / "challan" / "SMS_Delivery_Challan_Template.xlsx",
        "output_dir": OUTPUT_DIR / "challans",
    },
    "salary_slip": {
        "display_name": "Salary Slip",
        "template": TEMPLATES_DIR / "salary_slip" / "SMS_Salary_Slip_Template.xlsx",
        "output_dir": OUTPUT_DIR / "salary_slips",
    },
}


def get_template(document_key: str) -> Path:
    return Path(DOCUMENT_REGISTRY[document_key]["template"])
