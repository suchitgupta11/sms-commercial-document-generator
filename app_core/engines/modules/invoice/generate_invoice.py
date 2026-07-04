
from pathlib import Path
try:
    from app_core.config.settings import BASE_DIR, ASSETS_DIR
except Exception:
    BASE_DIR = Path(__file__).resolve().parents[4]
    ASSETS_DIR = BASE_DIR / "resources" / "assets"
import sys, textwrap, math, os
from datetime import datetime

from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

# -----------------------------
# Utility functions
# -----------------------------

def clean(v):
    return "" if v is None else str(v).strip()

def parse_rate(v):
    if v is None or clean(v) == "":
        return 0.0
    if isinstance(v, str):
        s = v.strip().replace("%", "").replace(",", "")
        if not s:
            return 0.0
        n = float(s)
    else:
        n = float(v)
    return n / 100.0 if n > 1 else n

def parse_float(v):
    if v is None or clean(v) == "":
        return 0.0
    if isinstance(v, str):
        return float(v.replace(",", "").strip())
    return float(v)

def map_sheet(wb, name):
    ws = wb[name]
    data = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and clean(row[0]):
            data[clean(row[0])] = row[1]
    return data

ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

def words_under_1000(n):
    n = int(n)
    out = []
    if n >= 100:
        out.append(ONES[n // 100])
        out.append("Hundred")
        n %= 100
    if n >= 20:
        out.append(TENS[n // 10])
        n %= 10
    if n > 0:
        out.append(ONES[n])
    return " ".join(out)

def amount_words(amount):
    n = int(round(amount))
    if n == 0:
        return "Indian Rupees Zero Only"
    parts = []
    crore = n // 10000000
    n %= 10000000
    lakh = n // 100000
    n %= 100000
    thousand = n // 1000
    n %= 1000
    if crore:
        parts.append(words_under_1000(crore) + " Crore")
    if lakh:
        parts.append(words_under_1000(lakh) + " Lakh")
    if thousand:
        parts.append(words_under_1000(thousand) + " Thousand")
    if n:
        parts.append(words_under_1000(n))
    return "Indian Rupees " + " ".join(parts) + " Only"

def money(v):
    return f"{float(v):,.2f}"


DOCUMENT_PROFILES = {
    "invoice": {
        "top_title": "TAX INVOICE",
        "top_suffix": True,
        "details_title": "INVOICE DETAILS",
        "no_label": "Invoice No.",
        "date_label": "Date",
        "no_keys": ["Invoice No.", "PI No.", "Proforma No.", "Proforma Invoice No."],
        "date_keys": ["Invoice Date", "PI Date", "Proforma Date", "Proforma Invoice Date"],
        "dispatch_title": "Invoice / Dispatch",
        "continued_title": "TAX INVOICE CONTINUED",
        "summary_title": "Invoice Summary",
        "footer_default": "This is a Computer Generated Tax Invoice",
        "file_suffix": "Tax_Invoice",
        "docx_title": "TAX INVOICE",
        "docx_no_label": "Invoice No.",
        "docx_date_label": "Date",
    },
    "proforma": {
        "top_title": "PROFORMA INVOICE",
        "top_suffix": False,
        "details_title": "PROFORMA DETAILS",
        "no_label": "PI No.",
        "date_label": "PI Date",
        "no_keys": ["PI No.", "Proforma No.", "Proforma Invoice No.", "Invoice No."],
        "date_keys": ["PI Date", "Proforma Date", "Proforma Invoice Date", "Invoice Date"],
        "dispatch_title": "Proforma / Dispatch",
        "continued_title": "PROFORMA INVOICE CONTINUED",
        "summary_title": "Proforma Summary",
        "footer_default": "This is a Computer Generated Proforma Invoice",
        "file_suffix": "Proforma_Invoice",
        "docx_title": "PROFORMA INVOICE",
        "docx_no_label": "PI No.",
        "docx_date_label": "PI Date",
    },
    "purchase_order": {
        "top_title": "PURCHASE ORDER",
        "top_suffix": False,
        "details_title": "PO DETAILS",
        "no_label": "PO No.",
        "date_label": "PO Date",
        "no_keys": ["PO No.", "Purchase Order No.", "Voucher No.", "Order No.", "Invoice No."],
        "date_keys": ["PO Date", "Purchase Order Date", "Dated", "Order Date", "Invoice Date"],
        "dispatch_title": "PO / Dispatch",
        "continued_title": "PURCHASE ORDER CONTINUED",
        "summary_title": "Purchase Order Summary",
        "footer_default": "This is a Computer Generated Purchase Order",
        "file_suffix": "Purchase_Order",
        "docx_title": "PURCHASE ORDER",
        "docx_no_label": "PO No.",
        "docx_date_label": "PO Date",
    },
}

def get_document_profile(data):
    kind = clean(data.get("document_kind")) or "invoice"
    return DOCUMENT_PROFILES.get(kind, DOCUMENT_PROFILES["invoice"])

def first_value(source, keys):
    for key in keys:
        value = clean(source.get(key))
        if value:
            return value
    return ""

def document_no(data):
    return first_value(data["invoice"], get_document_profile(data)["no_keys"])

def document_date(data):
    return first_value(data["invoice"], get_document_profile(data)["date_keys"])

def _first_from_maps(data, candidates):
    """Return first non-empty value from invoice/terms/settings maps."""
    for section in ("invoice", "terms", "settings"):
        src = data.get(section, {})
        for key in candidates:
            value = clean(src.get(key))
            if value:
                return value
    return ""

def proforma_valid_until(data):
    return _first_from_maps(data, [
        "Valid Until", "PI Valid Until", "Proforma Valid Until",
        "Validity", "Price Validity", "Offer Validity"
    ]) or "30 Days"

def proforma_payment_terms(data):
    return _first_from_maps(data, [
        "Payment Terms", "PI Payment Terms", "Proforma Payment Terms", "Payment"
    ]) or "100% Advance"

def proforma_delivery(data):
    return _first_from_maps(data, [
        "Delivery", "Delivery Period", "PI Delivery", "Proforma Delivery", "Lead Time"
    ]) or "As per discussion"

def document_detail_rows(data):
    """Rows shown in the premium details card for PDF and DOCX header image."""
    profile = get_document_profile(data)
    invoice = data.get("invoice", {})
    if clean(data.get("document_kind")) == "purchase_order":
        return [
            (profile["no_label"], document_no(data), True),
            (profile["date_label"], document_date(data), False),
            ("Supplier Ref", _first_from_maps(data, ["Supplier Ref./Order No.", "Supplier Ref", "Supplier Order No.", "Other Reference(s)"]) or "-", False),
            ("Payment Terms", _first_from_maps(data, ["Mode/Terms of Payment", "Payment Terms", "Payment"]) or "As per agreement", False),
            ("Delivery", _first_from_maps(data, ["Terms of Delivery", "Delivery", "Delivery Period"]) or "As per discussion", False),
        ]
    if clean(data.get("document_kind")) == "proforma":
        return [
            (profile["no_label"], document_no(data), True),
            (profile["date_label"], document_date(data), False),
            ("Valid Until", proforma_valid_until(data), False),
            ("Payment Terms", proforma_payment_terms(data), False),
            ("Delivery", proforma_delivery(data), False),
        ]
    return [
        (profile["no_label"], document_no(data), True),
        (profile["date_label"], document_date(data), False),
        ("e-Way Bill", clean(invoice.get("E-Way Bill No.")) or "-", False),
    ]

def footer_note(data):
    profile = get_document_profile(data)
    terms = data.get("terms", {})
    if clean(data.get("document_kind")) == "purchase_order":
        return clean(terms.get("PO Footer Note")) or clean(terms.get("Footer Note")) or profile["footer_default"]
    if clean(data.get("document_kind")) == "proforma":
        return clean(terms.get("Proforma Footer Note")) or profile["footer_default"]
    return clean(terms.get("Footer Note")) or profile["footer_default"]

def declaration_text(data):
    terms = data.get("terms", {})
    if clean(data.get("document_kind")) == "purchase_order":
        return (clean(terms.get("PO Declaration")) or clean(terms.get("Declaration")) or
                "Please supply the above goods/services as per agreed terms. This Purchase Order is valid subject to acceptance of commercial terms and delivery schedule.")
    if clean(data.get("document_kind")) == "proforma":
        return (clean(terms.get("Proforma Declaration")) or
                "This is a Proforma Invoice issued for reference and advance payment purposes only. "
                "It is not a Tax Invoice. A final Tax Invoice will be issued upon dispatch of goods.")
    return clean(terms.get("Declaration"))


def freight_item_from_data(data):
    """Return an optional freight charge as a taxable line item.

    Freight is configured in Invoice_Details / Settings / Terms using these keys:
    Freight Description, Freight HSN/SAC, Freight Amount, Freight GST %, Freight Unit.
    The charge participates in totals and HSN/GST summary exactly like the PDF reference.
    """
    amount = _first_from_maps(data, [
        "Freight Amount", "Freight Charges", "Freight", "Freight Value",
        "Freight (INC)", "Freight INC", "Freight Inclusive"
    ])
    freight_amount = parse_float(amount) if clean(amount) else 0.0
    if freight_amount <= 0:
        return None
    gst_raw = _first_from_maps(data, ["Freight GST %", "Freight GST", "Freight Tax %", "Default GST %"])
    gst_pct = parse_rate(gst_raw) if clean(gst_raw) else 0.0
    desc = _first_from_maps(data, ["Freight Description", "Freight Label", "Freight Name"]) or "Freight Charges"
    hsn = _first_from_maps(data, ["Freight HSN/SAC", "Freight HSN", "Freight SAC"]) or "996812"
    unit = _first_from_maps(data, ["Freight Unit", "Freight UOM"]) or "Lot"
    gst_amount = freight_amount * gst_pct
    return {
        "sr": "F",
        "description": desc,
        "hsn": hsn,
        "qty": 1.0,
        "unit": unit,
        "rate": freight_amount,
        "discount_pct": 0.0,
        "gst_pct": gst_pct,
        "gross": freight_amount,
        "discount_amount": 0.0,
        "taxable": freight_amount,
        "gst_amount": gst_amount,
        "is_freight": True,
    }

# -----------------------------
# Excel reading
# -----------------------------

def read_invoice_data(xlsx_path):
    wb = load_workbook(xlsx_path, data_only=True)
    company = map_sheet(wb, "Company_Details")
    bank = map_sheet(wb, "Bank_Details")
    customer = map_sheet(wb, "Customer_Details")
    invoice = map_sheet(wb, "Invoice_Details")
    terms = map_sheet(wb, "Terms")
    settings = map_sheet(wb, "Settings")

    ws = wb["Items"]
    headers = [clean(c.value) for c in ws[1]]
    header_map = {h.lower().replace(" ", "").replace("/", "").replace("%", ""): i for i, h in enumerate(headers)}

    def col(row, *names):
        for name in names:
            key = name.lower().replace(" ", "").replace("/", "").replace("%", "")
            if key in header_map:
                idx = header_map[key]
                return row[idx] if idx < len(row) else None
        return None

    product_discount_cells = []
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        desc = col(row, "Description")
        if not clean(desc):
            continue
        qty = parse_float(col(row, "Qty"))
        rate = parse_float(col(row, "Rate"))
        discount_raw = col(row, "Discount %", "Disc", "Discount")
        discount_present = clean(discount_raw) != ""
        product_discount_cells.append(discount_present)
        discount_pct = parse_rate(discount_raw)
        gst_pct = parse_rate(col(row, "GST %"))
        gross = qty * rate
        discount_amount = gross * discount_pct
        taxable = gross - discount_amount
        gst_amount = taxable * gst_pct
        items.append({
            "sr": clean(col(row, "Sr")) or str(len(items)+1),
            "description": clean(desc),
            "hsn": clean(col(row, "HSN/SAC")),
            "qty": qty,
            "unit": clean(col(row, "Unit")),
            "rate": rate,
            "discount_pct": discount_pct,
            "gst_pct": gst_pct,
            "gross": gross,
            "discount_amount": discount_amount,
            "taxable": taxable,
            "gst_amount": gst_amount,
            "discount_present": discount_present,
        })

    show_discount_column = bool(product_discount_cells) and all(product_discount_cells)

    data_preview = {
        "company": company,
        "bank": bank,
        "customer": customer,
        "invoice": invoice,
        "terms": terms,
        "settings": settings,
        "items": items,
        "document_kind": "invoice",
        "show_discount_column": show_discount_column,
    }
    freight_item = freight_item_from_data(data_preview)
    if freight_item:
        items.append(freight_item)

    taxable_total = sum(i["taxable"] for i in items)
    gst_total = sum(i["gst_amount"] for i in items)
    grand_total = taxable_total + gst_total

    # HSN summary
    hsn_summary = {}
    for it in items:
        key = (it["hsn"], it["gst_pct"])
        row = hsn_summary.setdefault(key, {"hsn": it["hsn"], "gst_pct": it["gst_pct"], "taxable": 0.0, "gst": 0.0})
        row["taxable"] += it["taxable"]
        row["gst"] += it["gst_amount"]

    return {
        "company": company,
        "bank": bank,
        "customer": customer,
        "invoice": invoice,
        "terms": terms,
        "settings": settings,
        "items": items,
        "taxable_total": taxable_total,
        "gst_total": gst_total,
        "grand_total": grand_total,
        "hsn_summary": list(hsn_summary.values()),
        "document_kind": "invoice",
        "show_discount_column": show_discount_column,
    }

# -----------------------------
# PDF drawing
# -----------------------------

W, H = A4
M = 11 * mm
CW = W - 2 * M

NAVY = colors.HexColor("#0B2F4F")
BLUE = colors.HexColor("#155C94")
ACCENT = colors.HexColor("#1F78B4")
LIGHT_BLUE = colors.HexColor("#EAF4FB")
SOFT_BLUE = colors.HexColor("#F4F9FD")
SOFT_GREY = colors.HexColor("#F8FAFC")
GREY = colors.HexColor("#5E6C84")
BORDER = colors.HexColor("#C9D6E2")
TEXT = colors.HexColor("#172B4D")
WHITE = colors.white

class InvoicePDF:
    def __init__(self, output_path, data, project_dir):
        self.output_path = output_path
        self.data = data
        self.project_dir = Path(project_dir)
        self.c = canvas.Canvas(str(output_path), pagesize=A4)
        self.page_no = 0
        self.total_pages = 1

    def line(self, x1, y1, x2, y2, color=BORDER, lw=0.6):
        self.c.setStrokeColor(color); self.c.setLineWidth(lw); self.c.line(x1, y1, x2, y2)

    def rect(self, x, y, w, h, stroke=BORDER, fill=None, lw=0.6, radius=0):
        if fill:
            self.c.setFillColor(fill)
            if radius: self.c.roundRect(x, y, w, h, radius, stroke=0, fill=1)
            else: self.c.rect(x, y, w, h, stroke=0, fill=1)
        if stroke:
            self.c.setStrokeColor(stroke); self.c.setLineWidth(lw)
            if radius: self.c.roundRect(x, y, w, h, radius, stroke=1, fill=0)
            else: self.c.rect(x, y, w, h, stroke=1, fill=0)

    def text(self, x, y, s, size=8, font="Helvetica", color=TEXT):
        self.c.setFont(font, size); self.c.setFillColor(color); self.c.drawString(x, y, str(s))

    def right(self, x, y, s, size=8, font="Helvetica", color=TEXT):
        self.c.setFont(font, size); self.c.setFillColor(color); self.c.drawRightString(x, y, str(s))

    def center(self, x, y, s, size=8, font="Helvetica", color=TEXT):
        self.c.setFont(font, size); self.c.setFillColor(color); self.c.drawCentredString(x, y, str(s))

    def wrap_lines(self, s, width_chars):
        lines = []
        for raw in str(s).split("\n"):
            if raw:
                lines.extend(textwrap.wrap(raw, width_chars))
            else:
                lines.append("")
        return lines

    def wrap_text(self, x, y, s, width_chars=45, size=7.2, leading=8.5, font="Helvetica", color=TEXT, max_lines=None):
        lines = self.wrap_lines(s, width_chars)
        if max_lines: lines = lines[:max_lines]
        for i, ln in enumerate(lines):
            self.text(x, y - i*leading, ln, size, font, color)
        return y - len(lines)*leading

    def label_value(self, x, y, label, value, lw=26*mm, size=7.1):
        self.text(x, y, label, size, "Helvetica-Bold", GREY)
        self.text(x+lw, y, value, size, "Helvetica", TEXT)

    def section_title(self, x, y, w, title):
        self.rect(x, y-6.5*mm, w, 6.5*mm, stroke=None, fill=LIGHT_BLUE)
        self.text(x+3*mm, y-4.4*mm, title, 8.3, "Helvetica-Bold", NAVY)

    def draw_logo_centered(self, box_x, box_y, box_w, box_h):
        logo_rel = clean(self.data["company"].get("Logo File")) or "assets/logo_from_doc.png"
        candidates = [
            self.project_dir / logo_rel,
            ASSETS_DIR / "invoice" / Path(logo_rel).name,
            ASSETS_DIR / Path(logo_rel).name,
            Path(__file__).resolve().parent / "assets" / Path(logo_rel).name,
        ]
        logo_path = next((p for p in candidates if p.exists()), candidates[0])
        if logo_path.exists():
            try:
                img = ImageReader(str(logo_path))
                iw, ih = img.getSize()
                scale = min(box_w/iw, box_h/ih)
                dw, dh = iw*scale, ih*scale
                self.c.drawImage(img, box_x+(box_w-dw)/2, box_y+(box_h-dh)/2, dw, dh, preserveAspectRatio=True, mask="auto")
                return
            except Exception:
                pass
        self.center(box_x+box_w/2, box_y+box_h/2, "SMS", 17, "Helvetica-Bold", BLUE)

    def draw_footer(self):
        terms = self.data["terms"]
        self.line(M, 14*mm, M+CW, 14*mm, BORDER, 0.5)
        footer = footer_note(self.data)
        email = clean(self.data["company"].get("Email"))
        self.center(W/2, 9*mm, f"{footer} | SMS Controls & Automation | {email}", 7.1, "Helvetica", GREY)
        self.right(W-M, 9*mm, f"Page {self.page_no} of {self.total_pages}", 7.1, "Helvetica", GREY)

    def start_page(self):
        self.page_no += 1
        self.rect(0, 0, W, H, stroke=None, fill=WHITE)

    def draw_header(self):
        company = self.data["company"]
        invoice = self.data["invoice"]
        y = H - M
        header_h = 44*mm
        self.rect(M, y-header_h, CW, header_h, stroke=NAVY, fill=WHITE, lw=0.8, radius=3)
        self.rect(M, y-9*mm, CW, 9*mm, stroke=None, fill=NAVY)
        self.text(M+4*mm, y-6.2*mm, clean(company.get("Company Name")) or "SMS Controls & Automation", 9.5, "Helvetica-Bold", WHITE)
        copy_type = clean(invoice.get("Copy Type")) or "ORIGINAL COPY"
        profile = get_document_profile(self.data)
        title_text = profile["top_title"] + (f" | {copy_type}" if profile.get("top_suffix") else "")
        self.right(M+CW-4*mm, y-6.2*mm, title_text, 9.5, "Helvetica-Bold", WHITE)

        self.draw_logo_centered(M + 3*mm, y - header_h + 8*mm, 34*mm, 24*mm)

        cx = M + 41*mm
        self.text(cx, y-17*mm, clean(company.get("Company Name")), 13.5, "Helvetica-Bold", NAVY)
        self.text(cx, y-22*mm, clean(company.get("Tagline")), 7.8, "Helvetica", GREY)
        self.text(cx, y-27*mm, clean(company.get("Address Line 1")), 7.1, "Helvetica", TEXT)
        self.text(cx, y-32*mm, clean(company.get("Address Line 2")), 7.1, "Helvetica", TEXT)
        gst_line = f"GSTIN: {clean(company.get('GSTIN'))} | State: {clean(company.get('State'))}, Code: {clean(company.get('State Code'))}"
        self.text(cx, y-36*mm, gst_line, 6.9, "Helvetica", TEXT)
        contact_parts = []
        if clean(company.get("Email")):
            contact_parts.append(f"Email: {clean(company.get('Email'))}")
        phone_value = clean(company.get("Phone")) or clean(company.get("Mobile")) or clean(company.get("Contact")) or clean(company.get("Contact Number")) or "+91-8826059159 / 8595231536"
        if phone_value:
            contact_parts.append(f"Phone: {phone_value}")
        if contact_parts:
            self.text(cx, y-41*mm, " | ".join(contact_parts), 6.9, "Helvetica", TEXT)

        detail_rows = document_detail_rows(self.data)
        card_w = 58*mm
        # V5.3.1: PI has 5 detail rows, so keep the card clearly below the top navy header.
        # Earlier top edge touched the navy strip and looked overlapped in the rendered PI PDF.
        is_pi_details = len(detail_rows) > 3
        card_h = 27*mm if not is_pi_details else 29*mm
        card_x = M + CW - card_w - 4*mm
        card_y = y - header_h + (5.5*mm if not is_pi_details else 2.0*mm)
        self.rect(card_x, card_y, card_w, card_h, stroke=ACCENT, fill=SOFT_BLUE, lw=0.7, radius=3)
        title_h = 7.5*mm if not is_pi_details else 7.0*mm
        self.rect(card_x, card_y+card_h-title_h, card_w, title_h, stroke=None, fill=NAVY)
        self.center(card_x+card_w/2, card_y+card_h-(5.1*mm if not is_pi_details else 4.8*mm), profile["details_title"], 7.8 if not is_pi_details else 7.0, "Helvetica-Bold", WHITE)

        label_x = card_x + 4.2*mm
        value_x = card_x + (27*mm if not is_pi_details else 24.5*mm)
        row_y = card_y + card_h - (12.5*mm if not is_pi_details else 10.8*mm)
        row_gap = (5.6*mm if not is_pi_details else 4.25*mm)
        for label, value, is_bold in detail_rows:
            self.text(label_x, row_y, label, 7.0 if not is_pi_details else 5.35, "Helvetica-Bold", GREY)
            self.text(value_x, row_y, str(value), 7.6 if not is_pi_details else 5.45, "Helvetica-Bold" if is_bold else "Helvetica", NAVY if is_bold else TEXT)
            row_y -= row_gap
        return y - header_h - 5*mm

    def draw_compact_continuation_header(self):
        invoice = self.data["invoice"]
        customer = self.data["customer"]
        y = H - M
        self.rect(M, y-18*mm, CW, 18*mm, stroke=NAVY, fill=WHITE, lw=0.7, radius=2)
        self.rect(M, y-7*mm, CW, 7*mm, stroke=None, fill=NAVY)
        self.text(M+4*mm, y-4.8*mm, "SMS CONTROLS & AUTOMATION", 8.8, "Helvetica-Bold", WHITE)
        profile = get_document_profile(self.data)
        self.right(M+CW-4*mm, y-4.8*mm, f"{profile['continued_title']} | {document_no(self.data)}", 8.8, "Helvetica-Bold", WHITE)
        self.text(M+4*mm, y-13*mm, "Bill To: " + clean(customer.get("Bill To Name")), 7.5, "Helvetica-Bold", NAVY)
        self.right(M+CW-4*mm, y-13*mm, profile["date_label"] + ": " + document_date(self.data), 7.5, "Helvetica", TEXT)
        return y - 23*mm

    def draw_party(self, y):
        customer = self.data["customer"]
        invoice = self.data["invoice"]
        box_h = 43*mm
        col_gap = 2*mm
        bill_w = 62*mm
        ship_w = 62*mm
        detail_w = CW - bill_w - ship_w - 2*col_gap
        x_bill = M
        x_ship = x_bill + bill_w + col_gap
        x_det = x_ship + ship_w + col_gap
        profile = get_document_profile(self.data)
        party_titles = ("Supplier", "Deliver To", profile["dispatch_title"]) if clean(self.data.get("document_kind")) == "purchase_order" else ("Bill To", "Ship To", profile["dispatch_title"])
        for x, w, title in [(x_bill,bill_w,party_titles[0]), (x_ship,ship_w,party_titles[1]), (x_det,detail_w,party_titles[2])]:
            self.rect(x, y-box_h, w, box_h, stroke=BORDER, fill=WHITE, lw=0.5, radius=2)
            self.section_title(x, y, w, title)

        self.text(x_bill+3*mm, y-12*mm, clean(customer.get("Bill To Name")), 8.0, "Helvetica-Bold", NAVY)
        bill_addr = f"{clean(customer.get('Bill To Address'))}\nGSTIN: {clean(customer.get('Bill To GSTIN'))}\nState: {clean(customer.get('Bill To State'))}, Code: {clean(customer.get('Bill To State Code'))}"
        self.wrap_text(x_bill+3*mm, y-17*mm, bill_addr, 34, 7.1, 8.1, max_lines=5)
        self.text(x_ship+3*mm, y-12*mm, clean(customer.get("Ship To Name")), 8.0, "Helvetica-Bold", NAVY)
        ship_addr = f"{clean(customer.get('Ship To Address'))}\nGSTIN: {clean(customer.get('Ship To GSTIN'))}\nState: {clean(customer.get('Ship To State'))}, Code: {clean(customer.get('Ship To State Code'))}"
        self.wrap_text(x_ship+3*mm, y-17*mm, ship_addr, 34, 7.1, 8.1, max_lines=5)

        if clean(self.data.get("document_kind")) == "proforma":
            rows = [
                ("Customer Ref", _first_from_maps(self.data, ["Customer Ref", "Customer Reference", "Reference No.", "Enquiry No.", "PO No."]) or "-"),
                ("Payment", proforma_payment_terms(self.data)),
                ("Delivery", proforma_delivery(self.data)),
                ("Valid Until", proforma_valid_until(self.data)),
                ("Destination", clean(invoice.get("Destination")) or "-"),
                ("Supply", clean(invoice.get("Place of Supply")) or "-"),
            ]
            for i, (lbl, val) in enumerate(rows):
                self.label_value(x_det+3*mm, y-(12+i*5)*mm, lbl, val, 24*mm)
        else:
            for i,(lbl,key) in enumerate([
                ("PO No.", "PO No."),
                ("PO Date", "PO Date"),
                ("Payment", "Payment Terms"),
                ("Dispatch", "Dispatch Through"),
                ("Destination", "Destination"),
                ("Supply", "Place of Supply")
            ]):
                self.label_value(x_det+3*mm, y-(12+i*5)*mm, lbl, clean(invoice.get(key)), 24*mm)
        return y - box_h - 5*mm

    def show_discount_column(self):
        return bool(self.data.get("show_discount_column"))

    @property
    def item_cols(self):
        # V5.4.1: Disc column is optional. If discount is not entered for every
        # product line in the template, hide the Disc column and redistribute
        # its width into Description/Taxable columns for a cleaner layout.
        if self.show_discount_column():
            cols = [8*mm, 70*mm, 18*mm, 12*mm, 13*mm, 21*mm, 13*mm, 33*mm]
        else:
            cols = [8*mm, 78*mm, 18*mm, 13*mm, 13*mm, 22*mm, 36*mm]
        xs = [M]
        for w in cols[:-1]: xs.append(xs[-1]+w)
        return xs, cols

    @property
    def item_headers(self):
        if self.show_discount_column():
            return ["Sr", "Description of Goods / Services", "HSN", "Qty", "Unit", "Rate", "Disc", "Taxable Value"]
        return ["Sr", "Description of Goods / Services", "HSN", "Qty", "Unit", "Rate", "Taxable Value"]

    def draw_item_header(self, y):
        xs, cols = self.item_cols
        row_h_header = 9*mm
        self.rect(M, y-row_h_header, CW, row_h_header, stroke=None, fill=NAVY)
        for i,h in enumerate(self.item_headers):
            self.center(xs[i]+cols[i]/2, y-5.9*mm, h, 6.9, "Helvetica-Bold", WHITE)
            if i > 0:
                self.line(xs[i], y, xs[i], y-row_h_header, colors.HexColor("#6B8FB0"), 0.35)
        return y - row_h_header

    def item_row_height(self, item):
        desc_width = 54 if not self.show_discount_column() else 48
        desc_lines = len(self.wrap_lines(item["description"], desc_width))
        return max(10*mm, (7 + min(desc_lines, 3)*4.2) * mm)

    def draw_item_row(self, y, item, row_h, stripe=False):
        xs, cols = self.item_cols
        fill = SOFT_GREY if stripe else WHITE
        self.rect(M, y-row_h, CW, row_h, stroke=BORDER, fill=fill, lw=0.35)
        self.center(xs[0]+cols[0]/2, y-6*mm, item["sr"], 7.1)
        desc_width = 54 if not self.show_discount_column() else 48
        lines = self.wrap_lines(item["description"], desc_width)
        if lines:
            first_font = "Helvetica-Bold" if item.get("is_freight") or lines[0] else "Helvetica"
            self.text(xs[1]+2*mm, y-5.5*mm, lines[0], 7.0, first_font, TEXT)
            for j, ln in enumerate(lines[1:3]):
                self.text(xs[1]+2*mm, y-(10+j*4.2)*mm, ln, 6.8, "Helvetica", TEXT)
        self.center(xs[2]+cols[2]/2, y-6*mm, item["hsn"], 7.1)
        self.center(xs[3]+cols[3]/2, y-6*mm, f"{item['qty']:g}", 7.1)
        self.center(xs[4]+cols[4]/2, y-6*mm, item["unit"], 7.1)
        self.right(xs[5]+cols[5]-2*mm, y-6*mm, money(item["rate"]), 7.1)
        if self.show_discount_column():
            disc = "-" if abs(item["discount_pct"]) < 0.000001 else f"{item['discount_pct']:.0%}"
            self.center(xs[6]+cols[6]/2, y-6*mm, disc, 7.1)
            self.right(xs[7]+cols[7]-2*mm, y-6*mm, money(item["taxable"]), 7.1)
        else:
            self.right(xs[6]+cols[6]-2*mm, y-6*mm, money(item["taxable"]), 7.1)
        for i in range(1, len(xs)):
            self.line(xs[i], y, xs[i], y-row_h, BORDER, 0.3)
        return y - row_h

    def draw_items_page(self, y, page_items):
        top_y = y
        y = self.draw_item_header(y)
        for idx, item in enumerate(page_items):
            y = self.draw_item_row(y, item, self.item_row_height(item), stripe=(idx % 2 == 1))
        self.rect(M, y, CW, top_y - y, stroke=BORDER, fill=None, lw=0.5)
        return y - 5*mm

    def normalize_hsn_mode(self):
        mode = clean(self.data["settings"].get("HSN Summary Mode")) or "Group by HSN"
        mode_key = mode.lower().replace("-", " ").replace("_", " ").strip()
        if mode_key in ("group by hsn", "group", "hsn", "grouped"):
            return "group"
        if mode_key in ("consolidated", "mixed", "summary"):
            return "consolidated"
        if mode_key in ("detailed", "line wise", "linewise", "line"):
            return "detailed"
        if mode_key in ("hidden", "hide", "none", "no"):
            return "hidden"
        return "group"

    def hsn_rows_for_mode(self):
        mode = self.normalize_hsn_mode()
        if mode == "hidden":
            return []

        if mode == "consolidated":
            if len(self.data["hsn_summary"]) > 1:
                rate = "Mixed"
                hsn = "Mixed"
            elif self.data["hsn_summary"]:
                rate = f"{self.data['hsn_summary'][0]['gst_pct']:.0%}"
                hsn = self.data["hsn_summary"][0]["hsn"]
            else:
                rate = "-"
                hsn = "-"
            return [{
                "hsn": hsn,
                "taxable": self.data["taxable_total"],
                "rate": rate,
                "gst": self.data["gst_total"],
                "total_tax": self.data["gst_total"],
            }]

        if mode == "detailed":
            rows = []
            for it in self.data["items"]:
                rows.append({
                    "hsn": f"{it['sr']} / {it['hsn']}",
                    "taxable": it["taxable"],
                    "rate": f"{it['gst_pct']:.0%}",
                    "gst": it["gst_amount"],
                    "total_tax": it["gst_amount"],
                })
            return rows

        # Default: group by HSN and GST rate.
        rows = []
        for row in sorted(self.data["hsn_summary"], key=lambda r: (r["hsn"], r["gst_pct"])):
            rows.append({
                "hsn": row["hsn"],
                "taxable": row["taxable"],
                "rate": f"{row['gst_pct']:.0%}",
                "gst": row["gst"],
                "total_tax": row["gst"],
            })
        return rows

    def draw_hsn_summary(self, y):
        mode = self.normalize_hsn_mode()
        if mode == "hidden":
            return y

        rows = self.hsn_rows_for_mode()
        # Header + row count + total row. Cap row height to keep this section compact.
        row_h = 6.2*mm
        max_rows_on_page = 7
        visible_rows = rows[:max_rows_on_page]
        extra_count = max(0, len(rows) - len(visible_rows))
        hsn_h = 8*mm + (len(visible_rows) * row_h) + 8*mm
        if extra_count:
            hsn_h += 5*mm

        self.rect(M, y-hsn_h, CW, hsn_h, stroke=BORDER, fill=WHITE, lw=0.5, radius=2)
        self.rect(M, y-8*mm, CW, 8*mm, stroke=None, fill=NAVY)

        title = {
            "group": "HSN/SAC Tax Summary - Group by HSN",
            "consolidated": "HSN/SAC Tax Summary - Consolidated",
            "detailed": "HSN/SAC Tax Summary - Detailed",
        }.get(mode, "HSN/SAC Tax Summary")
        self.text(M+3*mm, y-5.5*mm, title, 7.3, "Helvetica-Bold", WHITE)

        hsn_cols = [38*mm, 43*mm, 34*mm, 39*mm, 34*mm]
        hsn_xs = [M]
        for w in hsn_cols[:-1]:
            hsn_xs.append(hsn_xs[-1] + w)

        header_y = y - 8*mm
        for i, h in enumerate(["HSN/SAC", "Taxable Value", "GST Rate", "GST Amount", "Total Tax"]):
            self.center(hsn_xs[i] + hsn_cols[i]/2, header_y - 4.4*mm, h, 6.8, "Helvetica-Bold", NAVY)
            if i > 0:
                self.line(hsn_xs[i], header_y, hsn_xs[i], y-hsn_h, BORDER, 0.35)

        cur_y = header_y - 7*mm
        for idx, row in enumerate(visible_rows):
            if idx % 2 == 1:
                self.rect(M, cur_y-row_h+1.2*mm, CW, row_h, stroke=None, fill=SOFT_GREY)
            self.center(hsn_xs[0] + hsn_cols[0]/2, cur_y-3.2*mm, row["hsn"], 6.8)
            self.right(hsn_xs[1] + hsn_cols[1]-3*mm, cur_y-3.2*mm, money(row["taxable"]), 6.8)
            self.center(hsn_xs[2] + hsn_cols[2]/2, cur_y-3.2*mm, row["rate"], 6.8)
            self.right(hsn_xs[3] + hsn_cols[3]-3*mm, cur_y-3.2*mm, money(row["gst"]), 6.8)
            self.right(hsn_xs[4] + hsn_cols[4]-3*mm, cur_y-3.2*mm, money(row["total_tax"]), 6.8)
            self.line(M, cur_y-row_h+1.2*mm, M+CW, cur_y-row_h+1.2*mm, BORDER, 0.25)
            cur_y -= row_h

        if extra_count:
            self.text(M+3*mm, cur_y-2.8*mm, f"+ {extra_count} more detailed HSN rows included in invoice items", 6.6, "Helvetica-Oblique", GREY)
            cur_y -= 5*mm

        # Total row
        self.rect(M, cur_y-row_h+1.2*mm, CW, row_h, stroke=None, fill=LIGHT_BLUE)
        self.center(hsn_xs[0] + hsn_cols[0]/2, cur_y-3.2*mm, "Total", 6.9, "Helvetica-Bold", NAVY)
        self.right(hsn_xs[1] + hsn_cols[1]-3*mm, cur_y-3.2*mm, money(self.data["taxable_total"]), 6.9, "Helvetica-Bold", NAVY)
        self.right(hsn_xs[3] + hsn_cols[3]-3*mm, cur_y-3.2*mm, money(self.data["gst_total"]), 6.9, "Helvetica-Bold", NAVY)
        self.right(hsn_xs[4] + hsn_cols[4]-3*mm, cur_y-3.2*mm, money(self.data["gst_total"]), 6.9, "Helvetica-Bold", NAVY)

        return y - hsn_h - 5*mm

    def draw_final_blocks(self, y):
        bank = self.data["bank"]
        terms = self.data["terms"]
        left_w = 112*mm
        right_w = CW - left_w - 4*mm
        left_x = M
        right_x = M + left_w + 4*mm
        panel_h = 52*mm

        self.rect(left_x, y-panel_h, left_w, panel_h, stroke=BORDER, fill=WHITE, lw=0.5, radius=2)
        self.section_title(left_x, y, left_w, "Amount, Tax Words & Bank Details")
        self.text(left_x+4*mm, y-12*mm, "Amount Chargeable (in words)", 7.3, "Helvetica-Bold", GREY)
        self.wrap_text(left_x+4*mm, y-17*mm, amount_words(self.data["grand_total"]), 65, 7.5, 8.8)
        self.text(left_x+4*mm, y-27*mm, "Tax Amount (in words)", 7.3, "Helvetica-Bold", GREY)
        self.wrap_text(left_x+4*mm, y-32*mm, amount_words(self.data["gst_total"]), 65, 7.5, 8.8)

        if clean(bank.get("Show Bank Details")).lower() != "no":
            self.text(left_x+4*mm, y-42*mm, "Bank Details", 7.5, "Helvetica-Bold", NAVY)
            self.text(left_x+4*mm, y-47*mm, f"Bank: {clean(bank.get('Bank Name'))} | A/c Name: {clean(bank.get('Account Name'))}", 6.9, "Helvetica", TEXT)
            self.text(left_x+4*mm, y-51*mm, f"A/c No.: {clean(bank.get('Account Number'))} | IFSC: {clean(bank.get('IFSC Code'))} | Branch: {clean(bank.get('Branch'))}", 6.9, "Helvetica", TEXT)

        self.rect(right_x, y-panel_h, right_w, panel_h, stroke=BLUE, fill=SOFT_BLUE, lw=0.8, radius=2)
        self.rect(right_x, y-9*mm, right_w, 9*mm, stroke=None, fill=NAVY)
        self.center(right_x+right_w/2, y-6.1*mm, "TAX SUMMARY", 8.5, "Helvetica-Bold", WHITE)
        self.text(right_x+4*mm, y-17*mm, "Taxable Value", 8)
        self.right(right_x+right_w-4*mm, y-17*mm, money(self.data["taxable_total"]), 8, "Helvetica-Bold")
        tax_type = clean(self.data["settings"].get("Tax Type")) or "IGST"
        self.text(right_x+4*mm, y-25*mm, f"{tax_type} Total", 8)
        self.right(right_x+right_w-4*mm, y-25*mm, money(self.data["gst_total"]), 8, "Helvetica-Bold")
        self.line(right_x+4*mm, y-31*mm, right_x+right_w-4*mm, y-31*mm, BLUE, 0.7)
        self.text(right_x+4*mm, y-40*mm, "GRAND TOTAL", 10, "Helvetica-Bold", NAVY)
        self.right(right_x+right_w-4*mm, y-40*mm, "Rs. " + money(self.data["grand_total"]), 11.5, "Helvetica-Bold", NAVY)
        y -= panel_h + 5*mm

        y = self.draw_hsn_summary(y)

        bottom_h = 30*mm
        cols_bottom = [58*mm, 76*mm, CW-58*mm-76*mm]
        bx = [M, M+cols_bottom[0], M+cols_bottom[0]+cols_bottom[1]]
        for i, title in enumerate(["Remarks", "Declaration", "For SMS Controls & Automation"]):
            self.rect(bx[i], y-bottom_h, cols_bottom[i], bottom_h, stroke=BORDER, fill=WHITE, lw=0.5, radius=2)
            self.section_title(bx[i], y, cols_bottom[i], title)
        remarks = f"1. Warranty: {clean(terms.get('Warranty'))}\n2. Delivery: {clean(terms.get('Delivery'))}\n3. {clean(terms.get('Jurisdiction'))}"
        self.wrap_text(bx[0]+3*mm, y-12*mm, remarks, 33, 6.8, 7.9)
        self.wrap_text(bx[1]+3*mm, y-12*mm, declaration_text(self.data), 45, 6.8, 7.9)
        self.center(bx[2]+cols_bottom[2]/2, y-22*mm, "Authorized Signatory", 8.2, "Helvetica-Bold", NAVY)
        return y - bottom_h

    def final_blocks_required_height(self):
        # Summary + configurable HSN summary + Remarks/Declaration/Signature + gaps + footer safety.
        mode = self.normalize_hsn_mode()
        if mode == "hidden":
            hsn_height = 0
        else:
            row_count = len(self.hsn_rows_for_mode())
            visible = min(row_count, 7)
            hsn_height = 8*mm + visible * 6.2*mm + 8*mm
            if row_count > visible:
                hsn_height += 5*mm
        return (52*mm + 5*mm + hsn_height + (5*mm if hsn_height else 0) + 30*mm + 16*mm)

    def item_table_height(self, page_items):
        h = 9*mm  # item header
        for item in page_items:
            h += self.item_row_height(item)
        return h + 5*mm

    def paginate_items(self, items):
        """Dynamic pagination.

        Earlier V1 always split after 5 items, which caused large blank space on page 1.
        V1.1 fills each page based on available height. The final page reserves space for
        totals, HSN summary, bank details, declaration, signature and footer.
        """
        pages = []
        remaining = list(items)

        # Available item area on the first page after header + Bill/Ship/Dispatch section.
        first_y = H - M
        first_y -= (44*mm + 5*mm)   # header
        first_y -= (43*mm + 5*mm)   # party section
        first_available = first_y - 24*mm  # footer/safety bottom limit

        # Available item area on continuation pages after compact header.
        continuation_y = H - M - 23*mm
        continuation_available = continuation_y - 24*mm

        while remaining:
            is_first = len(pages) == 0
            available = first_available if is_first else continuation_available

            # Reserve final blocks if all remaining items can fit with the final summary.
            final_reserved = self.final_blocks_required_height()
            chosen = []
            used = 9*mm  # item table header

            for item in remaining:
                rh = self.item_row_height(item)
                if used + rh <= available:
                    chosen.append(item)
                    used += rh
                else:
                    break

            if not chosen:
                chosen = [remaining[0]]

            # If this would be the last page but final blocks do not fit, move items to next page.
            while len(chosen) > 1 and len(chosen) == len(remaining) and (used + final_reserved > available):
                removed = chosen.pop()
                used -= self.item_row_height(removed)

            # If only one item remains and final blocks still do not fit, keep item here and create summary page later.
            pages.append(chosen)
            remaining = remaining[len(chosen):]

        # If final blocks cannot fit on the last item page, add an empty final page for totals.
        # This prevents overlap and keeps footer independent.
        last_is_first = len(pages) == 1
        last_available = first_available if last_is_first else continuation_available
        if pages:
            last_used = self.item_table_height(pages[-1])
            if last_used + self.final_blocks_required_height() > last_available:
                pages.append([])

        return pages

    def render(self):
        item_pages = self.paginate_items(self.data["items"])
        self.total_pages = len(item_pages)

        for idx, page_items in enumerate(item_pages):
            self.start_page()
            if idx == 0:
                y = self.draw_header()
                y = self.draw_party(y)
            else:
                y = self.draw_compact_continuation_header()

            if page_items:
                y = self.draw_items_page(y, page_items)
            else:
                # final summary only page
                self.text(M, y, get_document_profile(self.data)["summary_title"], 9, "Helvetica-Bold", NAVY)
                y -= 8*mm

            if idx == len(item_pages) - 1:
                self.draw_final_blocks(y)

            self.draw_footer()
            self.c.showPage()

        self.c.save()



def _hex(color_obj):
    """Convert ReportLab HexColor/Color to RRGGBB for DOCX shading."""
    try:
        return color_obj.hexval().replace("0x", "").replace("#", "")[-6:].upper()
    except Exception:
        return str(color_obj).replace("#", "").upper()

DOCX_NAVY = _hex(NAVY)
DOCX_BLUE = _hex(BLUE)
DOCX_LIGHT_BLUE = _hex(LIGHT_BLUE)
DOCX_SOFT_BLUE = _hex(SOFT_BLUE)
DOCX_SOFT_GREY = _hex(SOFT_GREY)
DOCX_BORDER = _hex(BORDER)
DOCX_TEXT = _hex(TEXT)
DOCX_GREY = _hex(GREY)
DOCX_WHITE = "FFFFFF"


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn('w:shd'))
    if shd is None:
        shd = OxmlElement('w:shd')
        tc_pr.append(shd)
    shd.set(qn('w:fill'), fill)


def set_cell_border(cell, color=DOCX_BORDER, size='6', val='single'):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tc_pr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right'):
        tag = 'w:{}'.format(edge)
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn('w:val'), val)
        element.set(qn('w:sz'), size)
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), color)


def set_cell_margins(cell, top=80, start=90, bottom=80, end=90):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    mar = tc_pr.first_child_found_in('w:tcMar')
    if mar is None:
        mar = OxmlElement('w:tcMar')
        tc_pr.append(mar)
    for m, v in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        node = mar.find(qn(f'w:{m}'))
        if node is None:
            node = OxmlElement(f'w:{m}')
            mar.append(node)
        node.set(qn('w:w'), str(v))
        node.set(qn('w:type'), 'dxa')


def set_cell_width(cell, width_twips):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.first_child_found_in('w:tcW')
    if tc_w is None:
        tc_w = OxmlElement('w:tcW')
        tc_pr.append(tc_w)
    tc_w.set(qn('w:w'), str(width_twips))
    tc_w.set(qn('w:type'), 'dxa')
    # python-docx width API helps Word/LibreOffice honour the XML width.
    try:
        cell.width = Pt(width_twips / 20)
    except Exception:
        pass


def set_table_widths(table, widths):
    for row in table.rows:
        for idx, width in enumerate(widths):
            if idx < len(row.cells):
                set_cell_width(row.cells[idx], width)


def set_table_fixed_grid(table, widths):
    """Force exact DOCX table columns. Without tblGrid LibreOffice may redistribute columns equally."""
    table.autofit = False
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    layout = tbl_pr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tbl_pr.append(layout)
    layout.set(qn('w:type'), 'fixed')

    tbl_w = tbl_pr.find(qn('w:tblW'))
    if tbl_w is None:
        tbl_w = OxmlElement('w:tblW')
        tbl_pr.append(tbl_w)
    tbl_w.set(qn('w:w'), str(sum(widths)))
    tbl_w.set(qn('w:type'), 'dxa')

    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement('w:tblGrid')
    for w in widths:
        col = OxmlElement('w:gridCol')
        col.set(qn('w:w'), str(w))
        grid.append(col)
    tbl.insert(0, grid)
    set_table_widths(table, widths)


def clean_cell(cell, shade=None, border=DOCX_BORDER):
    if shade:
        set_cell_shading(cell, shade)
    set_cell_border(cell, border)
    set_cell_margins(cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def cell_text(cell, text_value='', bold=False, size=8, color=DOCX_TEXT, align=None, shade=None, border=DOCX_BORDER):
    cell.text = ''
    clean_cell(cell, shade, border)
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    run = p.add_run(str(text_value))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Arial'
    run.font.color.rgb = RGBColor.from_string(color)
    return p


def add_cell_line(cell, text_value='', bold=False, size=7.5, color=DOCX_TEXT, align=None):
    p = cell.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run(str(text_value))
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = 'Arial'
    r.font.color.rgb = RGBColor.from_string(color)
    return p


def add_docx_section_title(doc, title):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    cell_text(t.rows[0].cells[0], title, True, 8.5, DOCX_WHITE, WD_ALIGN_PARAGRAPH.CENTER, DOCX_NAVY, DOCX_NAVY)
    return t




def set_row_height(row, height_twips, exact=False):
    row.height = Pt(height_twips / 20)
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY if exact else WD_ROW_HEIGHT_RULE.AT_LEAST


def set_table_borders(table, color=DOCX_BORDER, size='6', val='single'):
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, color=color, size=size, val=val)


def remove_cell_border(cell):
    set_cell_border(cell, color='FFFFFF', size='0', val='nil')


def paragraph_run(paragraph, text, bold=False, size=8, color=DOCX_TEXT, font='Arial'):
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = font
    run.font.color.rgb = RGBColor.from_string(color)
    return run

def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement('w:tblHeader')
    tbl_header.set(qn('w:val'), 'true')
    tr_pr.append(tbl_header)


def _style_doc(doc):
    styles = doc.styles
    styles['Normal'].font.name = 'Arial'
    styles['Normal'].font.size = Pt(8)
    for section in doc.sections:
        section.top_margin = Inches(0.34)
        section.bottom_margin = Inches(0.38)
        section.left_margin = Inches(0.34)
        section.right_margin = Inches(0.34)
        section.header_distance = Inches(0.12)
        section.footer_distance = Inches(0.12)



def _docx_font(size, bold=False):
    """Return a scalable invoice-style font.

    Important fix:
    On macOS the previous font lookup often missed Arial/Helvetica, so Pillow
    fell back to a tiny bitmap default font. That is why the quotation header
    text became unreadably small on GUI-generated reports.
    """
    if ImageFont is None:
        return None

    candidates = [
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",

        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/Supplemental/Helvetica.ttf",

        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/calibrib.ttf" if bold else "C:/Windows/Fonts/calibri.ttf",
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
    ]

    for fp in candidates:
        try:
            if Path(fp).exists():
                return ImageFont.truetype(fp, int(size))
        except Exception:
            pass

    search_roots = [
        Path("/System/Library/Fonts"),
        Path("/System/Library/Fonts/Supplemental"),
        Path("/Library/Fonts"),
        Path.home() / "Library" / "Fonts",
        Path("C:/Windows/Fonts"),
        Path("/usr/share/fonts"),
    ]
    preferred = [
        "Arial Bold.ttf" if bold else "Arial.ttf",
        "LiberationSans-Bold.ttf" if bold else "LiberationSans-Regular.ttf",
        "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
        "Helvetica.ttc",
        "Calibri Bold.ttf" if bold else "Calibri.ttf",
        "Segoe UI Bold.ttf" if bold else "Segoe UI.ttf",
    ]
    for root in search_roots:
        try:
            if not root.exists():
                continue
            for name in preferred:
                matches = list(root.rglob(name))
                if matches:
                    return ImageFont.truetype(str(matches[0]), size)
        except Exception:
            continue

    try:
        return ImageFont.load_default(size=int(size))
    except TypeError:
        return ImageFont.load_default()

def _crop_logo_padding(logo):
    """Remove transparent/white padding around the logo to match PDF sizing."""
    try:
        logo = logo.convert('RGBA')
        alpha = logo.getchannel('A')
        bbox = alpha.getbbox()
        if bbox:
            logo = logo.crop(bbox)
        return logo
    except Exception:
        return logo


def _draw_text_fit(draw, xy, text, max_width, font_size, fill, bold=False, min_size=16):
    text = str(text)
    size = font_size
    while size >= min_size:
        f = _docx_font(size, bold)
        bbox = draw.textbbox((0, 0), text, font=f)
        if (bbox[2] - bbox[0]) <= max_width:
            draw.text(xy, text, font=f, fill=fill)
            return f
        size -= 1
    f = _docx_font(min_size, bold)
    draw.text(xy, text, font=f, fill=fill)
    return f


def _render_invoice_docx_header_image(data, profile, project_dir, out_path):
    """Render Invoice/PI header as a single high-resolution image.

    V5.2.0 rationale:
    Word/LibreOffice kept redistributing nested DOCX tables, causing the header
    to look like equal-width blocks. The PDF header is a fixed visual design, so
    DOCX now embeds a generated, high-resolution header image. This gives Invoice
    and PI the same premium header geometry as the PDF while keeping the rest of
    the DOCX editable.
    """
    company = data['company']; invoice = data['invoice']

    SCALE = 3
    BW, BH = 1500, 350
    Wpx, Hpx = BW * SCALE, BH * SCALE
    img = Image.new('RGB', (Wpx, Hpx), 'white')
    d = ImageDraw.Draw(img)

    navy = '#0B2F4F'
    accent = '#1F78B4'
    soft_blue = '#F4F9FD'
    text_col = '#172B4D'
    muted = '#5E6C84'
    white = '#FFFFFF'

    def sc(v):
        return int(round(v * SCALE))

    def rect(t):
        return tuple(sc(v) for v in t)

    # Outer box and top strip - same proportions as PDF.
    d.rounded_rectangle(rect((2, 2, BW-3, BH-3)), radius=sc(7), outline=navy, width=sc(2), fill=white)
    d.rectangle(rect((2, 2, BW-3, 72)), fill=navy)

    company_name = clean(company.get('Company Name')) or 'SMS Controls & Automation'
    copy_type = clean(invoice.get('Copy Type')) or 'ORIGINAL COPY'
    # Invoice keeps copy type; Proforma is a request/payment document and should not show ORIGINAL COPY.
    top_title = profile['top_title'] + (f' | {copy_type}' if profile.get('top_suffix') else '')
    d.text((sc(32), sc(24)), company_name, font=_docx_font(sc(30), True), fill=white)
    title_font = _docx_font(sc(30), True)
    tb = d.textbbox((0, 0), top_title, font=title_font)
    d.text((Wpx - sc(32) - (tb[2]-tb[0]), sc(24)), top_title, font=title_font, fill=white)

    # Logo position and size match the invoice PDF visual balance.
    logo_candidates = [
        ASSETS_DIR / 'invoice' / 'logo_from_doc.png',
        ASSETS_DIR / 'logo_from_doc.png',
        project_dir / 'resources' / 'assets' / 'invoice' / 'logo_from_doc.png',
        project_dir / 'assets' / 'logo_from_doc.png',
        Path(__file__).resolve().parent / 'assets' / 'logo_from_doc.png',
    ]
    logo_path = next((p for p in logo_candidates if p.exists()), logo_candidates[0])
    logo_box = rect((65, 105, 285, 310))
    if logo_path.exists():
        try:
            logo = _crop_logo_padding(Image.open(logo_path))
            target_w = logo_box[2] - logo_box[0]
            target_h = logo_box[3] - logo_box[1]
            scale = min(target_w / logo.width, target_h / logo.height)
            logo = logo.resize((max(1, int(logo.width * scale)), max(1, int(logo.height * scale))), Image.LANCZOS)
            lx = logo_box[0] + (target_w - logo.width) // 2
            ly = logo_box[1] + (target_h - logo.height) // 2
            if logo.mode != 'RGBA':
                logo = logo.convert('RGBA')
            img.paste(logo, (lx, ly), logo)
        except Exception:
            d.text((sc(105), sc(184)), 'SMS', font=_docx_font(sc(62), True), fill=navy)
    else:
        d.text((sc(105), sc(184)), 'SMS', font=_docx_font(sc(62), True), fill=navy)

    # Company block: starts right of logo, not in equal columns.
    cx = sc(330)
    _draw_text_fit(d, (cx, sc(104)), company_name, sc(610), sc(43), navy, True, sc(30))
    tagline = clean(company.get('Tagline')) or 'Industrial Automation | AC/DC Drives | Repairs | Retrofitting'
    _draw_text_fit(d, (cx, sc(154)), tagline, sc(610), sc(23), muted, False, sc(18))
    address1 = clean(company.get('Address Line 1')) or 'Office No.2, Plot No.241, Loha Mandi'
    address2 = clean(company.get('Address Line 2')) or 'Ghaziabad, Uttar Pradesh - 201009'
    gst = clean(company.get('GSTIN')) or '09JSMPS2386Q1ZB'
    state = clean(company.get('State')) or clean(company.get('State Name')) or 'Uttar Pradesh'
    state_code = clean(company.get('State Code')) or '09'
    email = clean(company.get('Email')) or 'info@smscontrols.com'
    phone = clean(company.get('Phone')) or clean(company.get('Contact No.')) or clean(company.get('Mobile')) or '+91-8826059159 / 8595231536'
    _draw_text_fit(d, (cx, sc(198)), address1, sc(610), sc(21), text_col, False, sc(17))
    _draw_text_fit(d, (cx, sc(233)), address2, sc(610), sc(21), text_col, False, sc(17))
    _draw_text_fit(d, (cx, sc(268)), f'GSTIN: {gst} | State: {state}, Code: {state_code}', sc(610), sc(20), text_col, False, sc(16))
    _draw_text_fit(d, (cx, sc(303)), f'Email: {email} | Phone: {phone}', sc(610), sc(20), text_col, False, sc(15))

    # Details card - right side, compact, with navy title strip and light body.
    rows = document_detail_rows(data)
    card_w, card_h = sc(440), sc(205 if len(rows) <= 3 else 232)
    card_x, card_y = Wpx - sc(72) - card_w, sc(95 if len(rows) <= 3 else 78)
    d.rounded_rectangle((card_x, card_y, card_x + card_w, card_y + card_h), radius=sc(6), outline=accent, width=sc(2), fill=soft_blue)
    d.rectangle((card_x, card_y, card_x + card_w, card_y + sc(58)), fill=navy)
    head = profile['details_title']
    hf = _docx_font(sc(27), True)
    hb = d.textbbox((0, 0), head, font=hf)
    d.text((card_x + (card_w - (hb[2]-hb[0]))//2, card_y + sc(17)), head, font=hf, fill=white)

    y = card_y + sc(80 if len(rows) > 3 else 86)
    row_gap = sc(31 if len(rows) > 3 else 43)
    label_font_size = sc(16 if len(rows) > 3 else 20)
    value_font_size = sc(16 if len(rows) > 3 else 21)
    for label, value, is_bold in rows:
        d.text((card_x + sc(38), y), label, font=_docx_font(label_font_size, True), fill=muted)
        d.text((card_x + sc(210), y), str(value), font=_docx_font(value_font_size, is_bold), fill=navy if is_bold else text_col)
        y += row_gap

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path, quality=96)
    return out_path


def _add_header_block(doc, data, profile, project_dir):
    # V5.2.0: use a single rasterized PDF-matched header image. This avoids
    # Word/LibreOffice table reflow and fixes both Invoice and PI in one place.
    tmp_dir = Path(project_dir) / '_generated_headers'
    header_path = tmp_dir / f"{clean(data.get('document_kind')) or 'invoice'}_{safe_file_part(document_no(data)) or 'sample'}_header_v521.png"
    _render_invoice_docx_header_image(data, profile, project_dir, header_path)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run()
    # Available width: A4 page - 0.34in left - 0.34in right ≈ 7.59in.
    run.add_picture(str(header_path), width=Inches(7.25))


def _add_party_block(doc, data, profile):
    customer=data['customer']; invoice=data['invoice']
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    party = doc.add_table(rows=2, cols=3)
    party.style = 'Table Grid'
    party.alignment = WD_TABLE_ALIGNMENT.CENTER
    party.autofit = False
    set_table_fixed_grid(party, [3800, 3800, 3680])  # V5.3.2: full visual width aligned with summary blocks
    headers = ['Supplier', 'Deliver To', profile['dispatch_title']] if clean(data.get('document_kind')) == 'purchase_order' else ['Bill To', 'Ship To', profile['dispatch_title']]
    for i, header in enumerate(headers):
        cell_text(party.rows[0].cells[i], header, True, 8.2, DOCX_NAVY, WD_ALIGN_PARAGRAPH.LEFT, DOCX_LIGHT_BLUE, DOCX_BORDER)
    bill=party.rows[1].cells[0]; ship=party.rows[1].cells[1]; dispatch=party.rows[1].cells[2]
    cell_text(bill, clean(customer.get('Bill To Name')), True, 8, DOCX_NAVY)
    for line in [clean(customer.get('Bill To Address')), f"GSTIN: {clean(customer.get('Bill To GSTIN'))}", f"State: {clean(customer.get('Bill To State'))}, Code: {clean(customer.get('Bill To State Code'))}"]:
        add_cell_line(bill, line, False, 7.4, DOCX_TEXT)
    cell_text(ship, clean(customer.get('Ship To Name')), True, 8, DOCX_NAVY)
    for line in [clean(customer.get('Ship To Address')), f"GSTIN: {clean(customer.get('Ship To GSTIN'))}", f"State: {clean(customer.get('Ship To State'))}, Code: {clean(customer.get('Ship To State Code'))}"]:
        add_cell_line(ship, line, False, 7.4, DOCX_TEXT)
    if clean(data.get('document_kind')) == 'proforma':
        ref = _first_from_maps(data, ['Customer Ref', 'Customer Reference', 'Reference No.', 'Enquiry No.', 'PO No.']) or '-'
        cell_text(dispatch, f"Customer Ref: {ref}", False, 7.6, DOCX_TEXT)
        for line in [f"Payment Terms: {proforma_payment_terms(data)}", f"Delivery: {proforma_delivery(data)}", f"Valid Until: {proforma_valid_until(data)}", f"Destination: {clean(invoice.get('Destination')) or '-'}"]:
            add_cell_line(dispatch, line, False, 7.4, DOCX_TEXT)
    else:
        cell_text(dispatch, f"PO No.: {clean(invoice.get('PO No.'))}", False, 7.6, DOCX_TEXT)
        for line in [f"PO Date: {clean(invoice.get('PO Date'))}", f"Payment: {clean(invoice.get('Payment Terms'))}", f"Dispatch: {clean(invoice.get('Dispatch Through'))}", f"Destination: {clean(invoice.get('Destination'))}"]:
            add_cell_line(dispatch, line, False, 7.4, DOCX_TEXT)


def _add_items_table(doc, data):
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    # V5.4.1: Disc column is visible only when every product line has a
    # discount cell populated in the Excel template. When hidden, its width is
    # redistributed so the table remains full-width and balanced in DOCX/PDF.
    show_disc = bool(data.get('show_discount_column'))
    col_count = 8 if show_disc else 7
    t = doc.add_table(rows=1, cols=col_count)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    if show_disc:
        widths = [483, 4201, 1080, 721, 778, 1261, 778, 1978]
        headers = ['Sr', 'Description of Goods / Services', 'HSN', 'Qty', 'Unit', 'Rate', 'Disc', 'Taxable Value']
    else:
        widths = [483, 4680, 1080, 780, 778, 1320, 2170]
        headers = ['Sr', 'Description of Goods / Services', 'HSN', 'Qty', 'Unit', 'Rate', 'Taxable Value']
    set_table_fixed_grid(t, widths)
    for i, header in enumerate(headers):
        cell_text(t.rows[0].cells[i], header, True, 6.7, DOCX_WHITE, WD_ALIGN_PARAGRAPH.CENTER, DOCX_NAVY, DOCX_NAVY)
    set_repeat_table_header(t.rows[0])
    for idx, it in enumerate(data['items']):
        row = t.add_row().cells
        if show_disc:
            vals = [it['sr'], it['description'], it['hsn'], f"{it['qty']:g}", it['unit'], money(it['rate']), '-' if abs(it['discount_pct']) < 1e-6 else f"{it['discount_pct']:.0%}", money(it['taxable'])]
        else:
            vals = [it['sr'], it['description'], it['hsn'], f"{it['qty']:g}", it['unit'], money(it['rate']), money(it['taxable'])]
        for i, val in enumerate(vals):
            shade = DOCX_SOFT_GREY if idx % 2 else DOCX_WHITE
            right_cols = (3,5,7) if show_disc else (3,5,6)
            center_cols = (0,2,4,6) if show_disc else (0,2,4)
            align = WD_ALIGN_PARAGRAPH.RIGHT if i in right_cols else (WD_ALIGN_PARAGRAPH.CENTER if i in center_cols else WD_ALIGN_PARAGRAPH.LEFT)
            taxable_idx = 7 if show_disc else 6
            cell_text(row[i], val, bool(it.get('is_freight') and i in (0,1,taxable_idx)), 6.6, DOCX_TEXT, align, shade)
    return t


def _add_summary_and_bank(doc, data):
    bank=data['bank']
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    wrap = doc.add_table(rows=2, cols=2)
    wrap.alignment = WD_TABLE_ALIGNMENT.CENTER
    wrap.autofit = False
    set_table_widths(wrap, [6700, 4100])
    # Header row mirrors the PDF's navy section bars.
    cell_text(wrap.rows[0].cells[0], 'Amount, Tax Words & Bank Details', True, 8.2, DOCX_WHITE, WD_ALIGN_PARAGRAPH.LEFT, DOCX_NAVY, DOCX_NAVY)
    cell_text(wrap.rows[0].cells[1], 'TAX SUMMARY', True, 8.6, DOCX_WHITE, WD_ALIGN_PARAGRAPH.CENTER, DOCX_NAVY, DOCX_NAVY)
    left, right = wrap.rows[1].cells
    clean_cell(left, DOCX_WHITE)
    clean_cell(right, DOCX_SOFT_BLUE, DOCX_BLUE)
    cell_text(left, 'Amount Chargeable (in words)', True, 7.1, DOCX_GREY, WD_ALIGN_PARAGRAPH.LEFT, DOCX_WHITE)
    add_cell_line(left, amount_words(data['grand_total']), False, 7.2, DOCX_TEXT)
    add_cell_line(left, 'Tax Amount (in words)', True, 7.1, DOCX_GREY)
    add_cell_line(left, amount_words(data['gst_total']), False, 7.2, DOCX_TEXT)
    if clean(bank.get('Show Bank Details')).lower() != 'no':
        add_cell_line(left, 'Bank Details', True, 7.5, DOCX_NAVY)
        add_cell_line(left, f"Bank: {clean(bank.get('Bank Name'))} | A/c Name: {clean(bank.get('Account Name'))}", False, 7.0, DOCX_TEXT)
        add_cell_line(left, f"A/c No.: {clean(bank.get('Account Number'))} | IFSC: {clean(bank.get('IFSC Code'))} | Branch: {clean(bank.get('Branch'))}", False, 7.0, DOCX_TEXT)

    tax_type = clean(data['settings'].get('Tax Type')) or 'IGST'
    rows = [('Taxable Value', money(data['taxable_total'])), (f'{tax_type} Total', money(data['gst_total'])), ('GRAND TOTAL', 'Rs. ' + money(data['grand_total']))]
    # Replace initial paragraph content in the right cell while keeping soft-blue background.
    cell_text(right, f"{rows[0][0]}: {rows[0][1]}", False, 7.7, DOCX_TEXT, WD_ALIGN_PARAGRAPH.RIGHT, DOCX_SOFT_BLUE, DOCX_BLUE)
    for idx, (lab, val) in enumerate(rows[1:], start=1):
        p = add_cell_line(right, f"{lab}: {val}", idx == 2, 8.5 if idx == 2 else 7.7, DOCX_NAVY if idx == 2 else DOCX_TEXT)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def _hsn_rows_for_docx(data):
    # Keep aligned with PDF default/group mode.
    rows=[]
    for row in sorted(data['hsn_summary'], key=lambda r: (r['hsn'], r['gst_pct'])):
        rows.append([row['hsn'], money(row['taxable']), f"{row['gst_pct']:.0%}", money(row['gst']), money(row['gst'])])
    return rows


def _add_hsn_summary(doc, data):
    mode = clean(data['settings'].get('HSN Summary Mode')).lower()
    if mode == 'hidden':
        return
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    add_docx_section_title(doc, 'HSN/SAC Tax Summary')
    t = doc.add_table(rows=1, cols=5)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    set_table_widths(t, [2100, 2400, 1900, 2200, 2200])
    for i, h in enumerate(['HSN/SAC', 'Taxable Value', 'GST Rate', 'GST Amount', 'Total Tax']):
        cell_text(t.rows[0].cells[i], h, True, 7.1, DOCX_NAVY, WD_ALIGN_PARAGRAPH.CENTER, DOCX_LIGHT_BLUE)
    for idx, vals in enumerate(_hsn_rows_for_docx(data)):
        row = t.add_row().cells
        for i, val in enumerate(vals):
            shade = DOCX_SOFT_GREY if idx % 2 else DOCX_WHITE
            align = WD_ALIGN_PARAGRAPH.RIGHT if i in (1,3,4) else WD_ALIGN_PARAGRAPH.CENTER
            cell_text(row[i], val, False, 6.9, DOCX_TEXT, align, shade)
    row = t.add_row().cells
    totals = ['Total', money(data['taxable_total']), '', money(data['gst_total']), money(data['gst_total'])]
    for i, val in enumerate(totals):
        align = WD_ALIGN_PARAGRAPH.RIGHT if i in (1,3,4) else WD_ALIGN_PARAGRAPH.CENTER
        cell_text(row[i], val, True, 7.1, DOCX_NAVY, align, DOCX_LIGHT_BLUE)


def _add_bottom_blocks(doc, data):
    terms=data['terms']
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    bottom = doc.add_table(rows=2, cols=3)
    bottom.alignment = WD_TABLE_ALIGNMENT.CENTER
    bottom.autofit = False
    set_table_widths(bottom, [3100, 4700, 3000])
    for i, title in enumerate(['Remarks', 'Declaration', 'For SMS Controls & Automation']):
        cell_text(bottom.rows[0].cells[i], title, True, 7.8, DOCX_WHITE, WD_ALIGN_PARAGRAPH.LEFT if i < 2 else WD_ALIGN_PARAGRAPH.CENTER, DOCX_NAVY, DOCX_NAVY)
    remarks = f"1. Warranty: {clean(terms.get('Warranty'))}\n2. Delivery: {clean(terms.get('Delivery'))}\n3. {clean(terms.get('Jurisdiction'))}"
    cell_text(bottom.rows[1].cells[0], remarks, False, 6.8, DOCX_TEXT)
    cell_text(bottom.rows[1].cells[1], declaration_text(data), False, 6.8, DOCX_TEXT)
    cell_text(bottom.rows[1].cells[2], '\n\n\nAuthorized Signatory', True, 7.8, DOCX_NAVY, WD_ALIGN_PARAGRAPH.CENTER)


def _resolve_project_dir(xlsx_path: Path) -> Path:
    """Return application root independent of whether the template is centralized or legacy."""
    return BASE_DIR


def generate_invoice_docx(xlsx_path, output_dir=None, document_kind="invoice"):
    """Generate premium editable DOCX matching the Invoice/PI PDF layout.

    V5.0.1 fix: the earlier DOCX path used a plain Word table layout. This
    version applies the same color palette, section headers, table borders,
    HSN/GST summary, bank details and footer treatment used by the PDF engine.
    """
    xlsx_path = Path(xlsx_path)
    project_dir = _resolve_project_dir(xlsx_path)
    data = read_invoice_data(xlsx_path)
    data["document_kind"] = document_kind
    profile = get_document_profile(data)
    output_dir = Path(output_dir) if output_dir is not None else project_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    inv_no = safe_file_part(document_no(data)) or ("PI" if document_kind == "proforma" else "Invoice")
    cust = safe_file_part(data["customer"].get("Bill To Name")) or "Customer"
    docx_path = output_dir / f"{inv_no}_{cust}_{profile['file_suffix']}.docx"

    doc = Document()
    _style_doc(doc)
    _add_header_block(doc, data, profile, project_dir)
    _add_party_block(doc, data, profile)
    _add_items_table(doc, data)
    _add_summary_and_bank(doc, data)
    _add_hsn_summary(doc, data)
    _add_bottom_blocks(doc, data)

    footer = doc.sections[0].footer.paragraphs[0]
    footer.text = footer_note(data)
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in footer.runs:
        r.font.size = Pt(7)
        r.font.name = 'Arial'
        r.font.color.rgb = RGBColor.from_string(DOCX_GREY)

    doc.save(docx_path)
    return docx_path


def safe_file_part(s):
    return "".join(ch if ch.isalnum() or ch in "-_ " else "_" for ch in clean(s)).strip().replace(" ", "_")

def generate_invoice(xlsx_path, output_dir=None, document_kind="invoice"):
    xlsx_path = Path(xlsx_path)
    project_dir = _resolve_project_dir(xlsx_path)
    data = read_invoice_data(xlsx_path)
    data["document_kind"] = document_kind
    profile = get_document_profile(data)
    if output_dir is None:
        output_dir = project_dir / "output"
    else:
        output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    inv_no = safe_file_part(document_no(data)) or ("PI" if document_kind == "proforma" else "Invoice")
    cust = safe_file_part(data["customer"].get("Bill To Name")) or "Customer"
    pdf_path = output_dir / f"{inv_no}_{cust}_{profile['file_suffix']}.pdf"

    pdf = InvoicePDF(pdf_path, data, project_dir)
    pdf.render()
    return pdf_path

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_invoice.py <template.xlsx> [output_folder] [--both|--docx|--pdf]")
        sys.exit(1)
    xlsx = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) >= 3 and not sys.argv[2].startswith("--") else None
    mode = "pdf"
    for arg in sys.argv[2:]:
        if arg in ("--both", "--docx", "--pdf"):
            mode = arg.replace("--", "")

    outputs = []
    if mode in ("pdf", "both"):
        outputs.append(generate_invoice(xlsx, out, document_kind="invoice"))
    if mode in ("docx", "both"):
        outputs.append(generate_invoice_docx(xlsx, out, document_kind="invoice"))

    for p in outputs:
        print(p)
