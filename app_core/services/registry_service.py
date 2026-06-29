from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List

from app_core.config.settings import MODULES_DIR, OUTPUT_DIR
from app_core.services.resource_service import ResourceManager


@dataclass(frozen=True)
class DocumentDefinition:
    display_name: str
    engine_key: str
    icon: str
    description: str
    default_template: Path
    output_subdir: Path
    module_dir: Path
    formats: str
    required_sheets: List[str]


DOCUMENTS: Dict[str, DocumentDefinition] = {
    "Quotation": DocumentDefinition(
        display_name="Quotation",
        engine_key="quotation",
        icon="📄",
        description="Generate professional commercial quotations with GST, freight, terms and bank details.",
        default_template=ResourceManager.template("quotation", "Proposal_Data_Template.xlsx"),
        output_subdir=OUTPUT_DIR / "quotations",
        module_dir=MODULES_DIR / "quotation",
        formats="DOCX + PDF",
        required_sheets=["Customer_Details", "Quotation_Details", "Items"],
    ),
    "Tax Invoice": DocumentDefinition(
        display_name="Tax Invoice",
        engine_key="invoice",
        icon="🧾",
        description="Create premium GST invoices with bill-to, ship-to, HSN summary and declaration.",
        default_template=ResourceManager.template("invoice", "SMS_Invoice_Data_Template.xlsx"),
        output_subdir=OUTPUT_DIR / "invoices",
        module_dir=MODULES_DIR / "invoice",
        formats="DOCX + PDF",
        required_sheets=["Company_Details", "Bank_Details", "Customer_Details", "Invoice_Details", "Items", "Terms", "Settings"],
    ),
    "Proforma Invoice": DocumentDefinition(
        display_name="Proforma Invoice",
        engine_key="proforma",
        icon="💼",
        description="Prepare payment-request proforma invoices using the finalized invoice-style layout.",
        default_template=ResourceManager.template("proforma", "SMS_Proforma_Invoice_Template.xlsx"),
        output_subdir=OUTPUT_DIR / "proforma_invoices",
        module_dir=MODULES_DIR / "proforma",
        formats="DOCX + PDF",
        required_sheets=["Company_Details", "Bank_Details", "Customer_Details", "Invoice_Details", "Items", "Terms", "Settings"],
    ),
    "Delivery Challan": DocumentDefinition(
        display_name="Delivery Challan",
        engine_key="challan",
        icon="🚚",
        description="Generate challans with dispatch details, customer details, item list and signature area.",
        default_template=ResourceManager.template("challan", "SMS_Delivery_Challan_Template.xlsx"),
        output_subdir=OUTPUT_DIR / "challans",
        module_dir=MODULES_DIR / "challan",
        formats="DOCX + PDF",
        required_sheets=["Company_Details", "Customer_Details", "Challan_Details", "Items", "Settings"],
    ),

    "Purchase Order": DocumentDefinition(
        display_name="Purchase Order",
        engine_key="purchase_order",
        icon="🛒",
        description="Create supplier purchase orders using the finalized premium invoice-style layout.",
        default_template=ResourceManager.template("purchase_order", "SMS_Purchase_Order_Template.xlsx"),
        output_subdir=OUTPUT_DIR / "purchase_orders",
        module_dir=MODULES_DIR / "purchase_order",
        formats="DOCX + PDF",
        required_sheets=["Company_Details", "Bank_Details", "Customer_Details", "Invoice_Details", "Items", "Terms", "Settings"],
    ),
    "Salary Slip": DocumentDefinition(
        display_name="Salary Slip",
        engine_key="salary_slip",
        icon="👤",
        description="Create premium salary slips with employee, payroll, earnings, deductions and net payable.",
        default_template=ResourceManager.template("salary_slip", "SMS_Salary_Slip_Template.xlsx"),
        output_subdir=OUTPUT_DIR / "salary_slips",
        module_dir=MODULES_DIR / "salary_slip",
        formats="DOCX + PDF",
        required_sheets=["Company_Details", "Employee_Details", "Salary_Details", "Earnings", "Deductions", "Settings"],
    ),
}


def document_names() -> list[str]:
    return list(DOCUMENTS.keys())


def get_document(name: str) -> DocumentDefinition:
    return DOCUMENTS[name]
