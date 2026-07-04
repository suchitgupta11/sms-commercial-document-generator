
from pathlib import Path
import sys, os, re
from datetime import datetime

try:
    from openpyxl import load_workbook
except ImportError:
    raise SystemExit("Please install dependencies: python3 -m pip install openpyxl reportlab")

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.platypus import Table, TableStyle, Paragraph
    from reportlab.lib.styles import ParagraphStyle
except ImportError:
    raise SystemExit("Please install dependencies: python3 -m pip install reportlab")

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:
    raise SystemExit("Please install dependencies: python3 -m pip install python-docx")

try:
    from PIL import Image, ImageDraw, ImageFont, ImageChops
except ImportError:
    Image = ImageDraw = ImageFont = ImageChops = None

try:
    from platform_utils import setup_logger
except Exception:
    try:
        sys.path.append(str(Path(__file__).resolve().parents[2]))
        from platform_utils import setup_logger
    except Exception:
        import logging
        def setup_logger(name):
            logger = logging.getLogger(name)
            if not logger.handlers:
                logger.setLevel(logging.INFO)
                handler = logging.StreamHandler()
                handler.setFormatter(logging.Formatter("%(levelname)s:%(name)s:%(message)s"))
                logger.addHandler(handler)
            return logger

def normalize_input_file(input_path):
    input_path = Path(input_path)
    suffix = input_path.suffix.lower()
    if suffix in [".xlsx", ".xls"]:
        return input_path
    if suffix == ".numbers":
        import platform
        import subprocess
        import tempfile
        if platform.system() != "Darwin":
            raise RuntimeError(".numbers files can only be converted on macOS with Apple Numbers installed.")
        out_dir = Path(tempfile.mkdtemp(prefix="sms_challan_numbers_"))
        script = (
            'tell application "Numbers"\n'
            f'open POSIX file "{input_path}"\n'
            'delay 1\n'
            'set theDoc to front document\n'
            f'export theDoc to POSIX file "{out_dir}" as Microsoft Excel\n'
            'close theDoc saving no\n'
            'end tell\n'
        )
        subprocess.run(["osascript", "-e", script], check=True)
        files = list(out_dir.glob("*.xlsx")) or list(out_dir.glob("*.xls"))
        if not files:
            raise RuntimeError("Unable to convert .numbers file to Excel. Please check Apple Numbers installation.")
        return files[0]
    raise RuntimeError(f"Unsupported input file type: {suffix}. Please use .xlsx, .xls, or .numbers.")

NAVY = colors.HexColor("#0D344F")
BLUE = colors.HexColor("#1E96D4")
SOFT = colors.HexColor("#EAF2FB")
TEXT = colors.HexColor("#0B2545")
MUTED = colors.HexColor("#5B6B82")
BORDER = colors.HexColor("#D5E0EC")
LIGHT_BG = colors.HexColor("#F7FAFD")
ACCENT = colors.HexColor("#F0A62A")
WHITE = colors.white
CW = A4[0] - 20*mm
GREY = MUTED
SOFT_BLUE = SOFT

def clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def money(v):
    try:
        return f"{float(v):,.2f}"
    except Exception:
        return clean(v)

def safe_filename(s):
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", clean(s))
    return s.strip("_") or "Delivery_Challan"

def get_map(ws):
    data = {}
    if ws is None:
        return data
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        data[clean(row[0])] = row[1] if len(row) > 1 else ""
    return data

def read_template(xlsx_path):
    wb = load_workbook(xlsx_path, data_only=True)
    required = ["Company_Details", "Customer_Details", "Challan_Details", "Items", "Settings"]
    missing = [s for s in required if s not in wb.sheetnames]
    if missing:
        raise ValueError(f"Missing required worksheet(s): {', '.join(missing)}. Available worksheet(s): {', '.join(wb.sheetnames)}")

    company = get_map(wb["Company_Details"])
    customer = get_map(wb["Customer_Details"])
    challan = get_map(wb["Challan_Details"])
    settings = get_map(wb["Settings"])

    ws = wb["Items"]
    headers = [clean(c.value) for c in ws[1]]
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None or clean(v) == "" for v in row):
            continue
        item = {}
        for i, h in enumerate(headers):
            item[h] = row[i] if i < len(row) else ""
        items.append(item)
    return company, customer, challan, items, settings

def first_existing_logo():
    candidates = [
        Path(__file__).parent / "assets" / "logo_from_doc.png",
        Path(__file__).parent / "assets" / "logo.png",
        Path(__file__).parent.parent / "invoice" / "assets" / "logo_from_doc.png",
        Path(__file__).parent.parent / "invoice" / "assets" / "logo.png",
    ]
    for c in candidates:
        if c.exists():
            return c
    return None

class ChallanPDF:
    def __init__(self, output_path, company, customer, challan, items, settings):
        self.output_path = Path(output_path)
        self.company = company
        self.customer = customer
        self.challan = challan
        self.items = items
        self.settings = settings
        self.logo = first_existing_logo()
        self.c = canvas.Canvas(str(self.output_path), pagesize=A4)
        self.W, self.H = A4
        self.M = 10 * mm

    def text_fit(self, text, x, y, max_width, size, font="Helvetica-Bold", min_size=7.5, color=TEXT):
        text = clean(text)
        s = size
        while s > min_size and stringWidth(text, font, s) > max_width:
            s -= 0.35
        self.c.setFont(font, s)
        self.c.setFillColor(color)
        self.c.drawString(x, y, text)

    def line(self, x1, y1, x2, y2, color=BORDER, width=0.5):
        self.c.setStrokeColor(color)
        self.c.setLineWidth(width)
        self.c.line(x1, y1, x2, y2)

    def draw_logo_centered(self, box_x, box_y, box_w, box_h):
        if self.logo and self.logo.exists():
            try:
                from reportlab.lib.utils import ImageReader
                img = ImageReader(str(self.logo))
                iw, ih = img.getSize()
                scale = min(box_w/iw, box_h/ih)
                dw, dh = iw*scale, ih*scale
                self.c.drawImage(img, box_x+(box_w-dw)/2, box_y+(box_h-dh)/2, dw, dh, preserveAspectRatio=True, mask="auto")
                return
            except Exception:
                pass
        self.c.setFillColor(BLUE)
        self.c.setFont("Helvetica-Bold", 17)
        self.c.drawCentredString(box_x+box_w/2, box_y+box_h/2, "SMS")

    def draw_header(self):
        """Invoice-matched header for Delivery Challan.

        Same dimensions/proportions as Tax Invoice header.
        Only text labels are changed:
        - TAX INVOICE -> DELIVERY CHALLAN
        - INVOICE DETAILS -> CHALLAN DETAILS
        """
        c, W, H, M = self.c, self.W, self.H, self.M
        company = self.company
        challan = self.challan

        y = H - M
        header_h = 44*mm

        # Same outer header as invoice
        c.setStrokeColor(NAVY)
        c.setLineWidth(0.8)
        c.roundRect(M, y-header_h, W-2*M, header_h, 3, stroke=1, fill=0)

        c.setFillColor(NAVY)
        c.rect(M, y-9*mm, W-2*M, 9*mm, stroke=0, fill=1)

        company_name = clean(company.get("Company Name")) or "SMS Controls & Automation"
        c.setFillColor(WHITE)
        c.setFont("Helvetica-Bold", 9.5)
        c.drawString(M+4*mm, y-6.2*mm, company_name)
        c.drawRightString(W-M-4*mm, y-6.2*mm, "DELIVERY CHALLAN")

        # Same logo position/size as invoice
        self.draw_logo_centered(M + 3*mm, y - header_h + 8*mm, 34*mm, 24*mm)

        cx = M + 41*mm
        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold", 13.5)
        c.drawString(cx, y-17*mm, company_name)
        c.setFillColor(GREY)
        c.setFont("Helvetica", 7.8)
        c.drawString(cx, y-22*mm, clean(company.get("Tagline")))

        c.setFillColor(TEXT)
        c.setFont("Helvetica", 7.1)
        c.drawString(cx, y-27*mm, clean(company.get("Address Line 1")) or clean(company.get("Address")))
        c.drawString(cx, y-32*mm, clean(company.get("Address Line 2")))

        gst_line = f"GSTIN: {clean(company.get('GSTIN'))} | State: {clean(company.get('State'))}, Code: {clean(company.get('State Code'))}"
        c.setFont("Helvetica", 6.9)
        c.drawString(cx, y-36*mm, gst_line)

        contact_parts = []
        if clean(company.get("Email")):
            contact_parts.append(f"Email: {clean(company.get('Email'))}")
        phone_value = clean(company.get("Phone")) or clean(company.get("Mobile")) or clean(company.get("Contact")) or "+91-8826059159 / 8595231536"
        if phone_value:
            contact_parts.append(f"Phone: {phone_value}")
        if contact_parts:
            c.drawString(cx, y-41*mm, " | ".join(contact_parts))

        # Same details card geometry as invoice
        card_w = 58*mm
        card_h = 27*mm
        card_x = W - M - card_w - 4*mm
        card_y = y - header_h + 5.5*mm

        c.setFillColor(SOFT_BLUE)
        c.setStrokeColor(BLUE)
        c.setLineWidth(0.7)
        c.roundRect(card_x, card_y, card_w, card_h, 3, stroke=1, fill=1)

        c.setFillColor(NAVY)
        c.rect(card_x, card_y+card_h-7.5*mm, card_w, 7.5*mm, stroke=0, fill=1)

        c.setFillColor(WHITE)
        c.setFont("Helvetica-Bold", 7.8)
        c.drawCentredString(card_x+card_w/2, card_y+card_h-5.1*mm, "CHALLAN DETAILS")

        date_val = challan.get("Date")
        if hasattr(date_val, "strftime"):
            date_val = date_val.strftime("%d-%b-%Y")

        label_x = card_x + 5*mm
        value_x = card_x + 27*mm
        row1 = card_y + card_h - 12.5*mm
        row_gap = 5.6*mm

        rows = [
            ("Challan No.", clean(challan.get("Challan No.")) or "-"),
            ("Date", clean(date_val) or "-"),
            ("Ref./Order", clean(challan.get("Ref./Order No.")) or "-"),
        ]
        for idx, (label, value) in enumerate(rows):
            yy = row1 - idx*row_gap
            c.setFillColor(GREY)
            c.setFont("Helvetica-Bold", 7.1)
            c.drawString(label_x, yy, label)
            c.setFillColor(NAVY if idx == 0 else TEXT)
            c.setFont("Helvetica-Bold" if idx == 0 else "Helvetica", 7.7)
            c.drawString(value_x, yy, value[:20])

        return y - header_h - 5*mm

    def draw_party_boxes(self, top_y):
        W, M = self.W, self.M
        gap = 7*mm
        box_h = 46*mm
        col_w = (W - 2*M - gap) / 2
        y = top_y - box_h
        self.party_box(M, y, col_w, box_h, "BILL TO", "Bill To")
        self.party_box(M + col_w + gap, y, col_w, box_h, "SHIP TO (CONSIGNEE)", "Ship To")
        return y - 5*mm

    def party_box(self, x, y, w, h, title, prefix):
        c = self.c
        c.setFillColor(WHITE)
        c.setStrokeColor(BORDER)
        c.roundRect(x, y, w, h, 3, stroke=1, fill=1)
        c.setFillColor(NAVY)
        c.roundRect(x, y+h-8.5*mm, w, 8.5*mm, 3, stroke=0, fill=1)
        c.rect(x, y+h-8.5*mm, w, 3*mm, stroke=0, fill=1)
        c.setFillColor(WHITE)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(x + 4*mm, y+h-5.8*mm, title)

        name = clean(self.customer.get(f"{prefix} Name")) or clean(self.customer.get("Customer Name"))
        lines = [
            (name, True),
            (clean(self.customer.get(f"{prefix} Address 1")), False),
            (clean(self.customer.get(f"{prefix} Address 2")), False),
            (clean(self.customer.get(f"{prefix} State")), False),
            (f"GSTIN      : {clean(self.customer.get(f'{prefix} GSTIN'))}", False),
            (f"State       : {clean(self.customer.get(f'{prefix} State Code'))}", False),
        ]
        yy = y + h - 15*mm
        for text, bold in lines:
            if not text or text.endswith(": "):
                continue
            c.setFillColor(TEXT)
            c.setFont("Helvetica-Bold" if bold else "Helvetica", 8.0 if bold else 7.6)
            c.drawString(x + 4*mm, yy, text[:62])
            yy -= 5.5*mm

    def draw_items(self, top_y):
        W, M = self.W, self.M
        style = ParagraphStyle("item", fontName="Helvetica", fontSize=7.8, leading=10, textColor=TEXT)
        headers = ["S. No.", "Description of Goods", "HSN / SAC", "Qty.", "Unit", "Remarks"]
        rows = [headers]
        total_qty = 0.0
        unit = ""
        for idx, it in enumerate(self.items, start=1):
            qty = it.get("Qty", "")
            try:
                total_qty += float(qty)
            except Exception:
                pass
            unit = clean(it.get("Unit")) or unit
            rows.append([
                clean(it.get("S.No")) or str(idx),
                Paragraph(clean(it.get("Description")), style),
                clean(it.get("HSN/SAC")) or clean(it.get("HSN")),
                clean(qty),
                clean(it.get("Unit")),
                clean(it.get("Remarks")) or "-",
            ])
        total_display = clean(int(total_qty) if float(total_qty).is_integer() else total_qty) if total_qty else ""
        rows.append(["", "TOTAL", "", total_display, unit, ""])

        item_h = 68*mm
        y = top_y - item_h
        col_widths = [13*mm, 72*mm, 27*mm, 20*mm, 22*mm, 35*mm]
        tbl = Table(rows, colWidths=col_widths, rowHeights=[9*mm] + [48*mm]*len(self.items) + [11*mm])
        tbl.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.45, BORDER),
            ("BACKGROUND", (0,0), (-1,0), NAVY),
            ("TEXTCOLOR", (0,0), (-1,0), WHITE),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,0), 8.0),
            ("FONTNAME", (0,1), (-1,-1), "Helvetica"),
            ("FONTSIZE", (0,1), (-1,-1), 7.8),
            ("ALIGN", (0,1), (0,-1), "CENTER"),
            ("ALIGN", (2,1), (-1,-1), "CENTER"),
            ("FONTNAME", (1,-1), (-1,-1), "Helvetica-Bold"),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("TOPPADDING", (0,0), (-1,-1), 5),
            ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ]))
        tbl.wrapOn(self.c, W - 2*M, item_h)
        tbl.drawOn(self.c, M, y)
        return y - 5*mm

    def draw_info_blocks(self, top_y):
        c, W, M = self.c, self.W, self.M
        gap = 8*mm
        h = 35*mm
        col_w = (W - 2*M - gap) / 2
        y = top_y - h

        c.setFillColor(WHITE)
        c.setStrokeColor(BORDER)
        c.roundRect(M, y, col_w, h, 3, stroke=1, fill=1)
        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold", 8.3)
        c.drawString(M+4*mm, y+h-7*mm, "REMARKS")
        self.line(M+4*mm, y+h-10*mm, M+13*mm, y+h-10*mm, ACCENT, 1)
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 7.4)
        remarks = clean(self.challan.get("Remarks")) or "There is no commercial value involved in this transaction."
        delivery_text = f"Delivery: {clean(self.challan.get('Terms of Delivery')) or '-'}"

        def wrap_line(txt, max_chars=60):
            txt = clean(txt)
            words = txt.split()
            out, cur = [], ""
            for word in words:
                if len(cur) + len(word) + 1 <= max_chars:
                    cur = (cur + " " + word).strip()
                else:
                    if cur:
                        out.append(cur)
                    cur = word
            if cur:
                out.append(cur)
            return out or [""]

        yy = y+h-17*mm

        # Show only template remarks + delivery. Do not add duplicate hardcoded text.
        for j, ln in enumerate(wrap_line(remarks, 60)[:3]):
            prefix = "- " if j == 0 else "  "
            c.drawString(M+5*mm, yy, prefix + ln)
            yy -= 4.8*mm

        # Always show delivery line explicitly.
        if yy >= y+7*mm:
            yy -= 1.0*mm
            c.drawString(M+5*mm, yy, "- " + delivery_text[:78])

        x2 = M + col_w + gap
        c.setFillColor(WHITE)
        c.setStrokeColor(BORDER)
        c.roundRect(x2, y, col_w, h, 3, stroke=1, fill=1)
        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold", 8.3)
        c.drawString(x2+4*mm, y+h-7*mm, "MATERIAL MOVEMENT DETAILS")
        self.line(x2+4*mm, y+h-10*mm, x2+13*mm, y+h-10*mm, ACCENT, 1)

        total = 0.0
        for it in self.items:
            try:
                total += float(it.get("Amount", 0))
            except Exception:
                pass
        show_value = clean(self.settings.get("Show Value")).lower() != "no"
        amount_words = clean(self.settings.get("Amount In Words")) or "INR: Sixty-Three Thousand Six Hundred Rupees Only"
        c.setFillColor(TEXT)
        c.setFont("Helvetica", 7.8)
        c.drawString(x2+5*mm, y+h-18*mm, "Material Value (For Reference Only)")
        c.setFont("Helvetica-Bold", 8.4)
        c.drawRightString(x2+col_w-5*mm, y+h-18*mm, f"Rs. {money(total)}" if show_value else "-")
        c.setFont("Helvetica", 7.8)
        c.drawString(x2+5*mm, y+h-27*mm, "Amount in Words")
        c.setFont("Helvetica-Bold", 7.8)
        c.drawString(x2+5*mm, y+h-32*mm, amount_words[:72])
        return y - 5*mm

    def draw_signature(self, top_y):
        c, W, M = self.c, self.W, self.M
        h = 32*mm
        y = top_y - h
        c.setFillColor(WHITE)
        c.setStrokeColor(BORDER)
        c.roundRect(M, y, W-2*M, h, 3, stroke=1, fill=1)
        mid = M + (W - 2*M) * 0.47
        self.line(mid, y, mid, y+h, BORDER, 0.4)

        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold", 8.3)
        c.drawString(M+4*mm, y+h-7*mm, "PREPARED BY")
        c.drawString(mid+4*mm, y+h-7*mm, "FOR SMS CONTROLS & AUTOMATION")
        self.line(M+4*mm, y+h-10*mm, M+13*mm, y+h-10*mm, ACCENT, 1)
        self.line(mid+4*mm, y+h-10*mm, mid+13*mm, y+h-10*mm, ACCENT, 1)

        c.setFillColor(TEXT)
        c.setFont("Helvetica", 7.8)
        c.drawString(M+4*mm, y+18*mm, "Name")
        c.drawString(M+24*mm, y+18*mm, ":")
        self.line(M+30*mm, y+18*mm, mid-10*mm, y+18*mm, colors.HexColor("#9AA6B2"), 0.4)
        c.drawString(M+4*mm, y+4*mm, "Date")
        c.drawString(M+24*mm, y+4*mm, ":")
        self.line(M+30*mm, y+4*mm, mid-10*mm, y+4*mm, colors.HexColor("#9AA6B2"), 0.4)

        self.line(mid+16*mm, y+14*mm, mid+51*mm, y+14*mm, colors.black, 0.5)
        self.line(W-M-52*mm, y+14*mm, W-M-12*mm, y+14*mm, colors.black, 0.5)
        c.setFont("Helvetica", 7.5)
        c.drawCentredString(mid+33.5*mm, y+7*mm, "Verified By")
        c.drawCentredString(W-M-32*mm, y+7*mm, "Authorized Signatory")
        return y

    def draw_footer(self):
        c, W, M = self.c, self.W, self.M
        c.setStrokeColor(BORDER)
        c.setLineWidth(0.5)
        c.line(M, 14*mm, W-M, 14*mm)
        email = clean(self.company.get("Email")) or "info@smscontrols.com"
        c.setFillColor(GREY)
        c.setFont("Helvetica", 7.1)
        c.drawCentredString(W/2, 9*mm, f"This is a Computer Generated Delivery Challan | SMS Controls & Automation | {email}")

    def build(self):
        y = self.draw_header()
        y = self.draw_party_boxes(y)
        y = self.draw_items(y)
        y = self.draw_info_blocks(y)
        y = self.draw_signature(y)
        self.draw_footer()
        self.c.save()


def _docx_set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), fill.replace("#", ""))

def _docx_set_cell_border(cell, color="D5E0EC", size="6"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right"):
        tag = "w:" + edge
        element = tcBorders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            tcBorders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color.replace("#", ""))

def _docx_set_cell_margins(cell, top=80, start=100, bottom=80, end=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in [("top", top), ("start", start), ("bottom", bottom), ("end", end)]:
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")

def _docx_set_table_width(table, width_inches):
    table.autofit = False
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), str(int(width_inches * 1440)))
    tblW.set(qn("w:type"), "dxa")

def _docx_run(paragraph, text, size=9, bold=False, color="0B2545"):
    run = paragraph.add_run(clean(text))
    run.font.name = "Aptos"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color.replace("#", ""))
    return run

def _docx_cell_text(cell, text, size=8.5, bold=False, color="0B2545", align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(0)
    _docx_run(p, text, size=size, bold=bold, color=color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _font(size, bold=False):
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
                return ImageFont.truetype(fp, size)
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
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()

def _crop_logo_padding(logo):
    logo = logo.convert("RGBA")
    alpha_bbox = logo.getbbox()
    if alpha_bbox:
        logo = logo.crop(alpha_bbox)
    try:
        rgb_white = Image.new("RGBA", logo.size, (255, 255, 255, 255))
        diff = ImageChops.difference(logo, rgb_white)
        bbox = diff.getbbox()
        if bbox:
            logo = logo.crop(bbox)
    except Exception:
        pass
    return logo

def _render_challan_docx_header_image(company, challan, out_path):
    """Render a Challan DOCX header image using the same invoice-style geometry as the PDF."""
    if Image is None:
        return None

    company_name = clean(company.get("Company Name")) or "SMS Controls & Automation"
    tagline = clean(company.get("Tagline")) or "Industrial Automation | AC/DC Drives | Repairs | Retrofitting"
    address1 = clean(company.get("Address Line 1")) or clean(company.get("Address")) or "Office No.2, Plot No.241, Loha Mandi"
    address2 = clean(company.get("Address Line 2")) or "Ghaziabad, Uttar Pradesh - 201009"
    gst = clean(company.get("GSTIN")) or "09JSMPS2386Q1ZB"
    state = clean(company.get("State")) or "Uttar Pradesh"
    code = clean(company.get("State Code")) or "09"
    email = clean(company.get("Email")) or "info@smscontrols.com"
    phone = clean(company.get("Phone")) or clean(company.get("Mobile")) or clean(company.get("Contact")) or "+91-8826059159 / 8595231536"

    date_val = challan.get("Date")
    if hasattr(date_val, "strftime"):
        date_val = date_val.strftime("%d-%b-%Y")

    SCALE = 3
    BW, BH = 1500, 350
    W, H = BW * SCALE, BH * SCALE
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    navy = "#0D344F"
    blue = "#1E96D4"
    soft = "#EAF2FB"
    text_col = "#0B2545"
    muted = "#5B6B82"
    white = "#FFFFFF"

    def sc(v):
        return int(round(v * SCALE))

    def font(size, bold=False):
        return _font(sc(size), bold)

    def rect(t):
        return tuple(sc(v) for v in t)

    def draw_fit(txt, xy, max_width, size, bold, fill, min_size=20):
        s = size
        f = font(s, bold)
        while s > min_size:
            try:
                bb = d.textbbox((0, 0), txt, font=f)
                if (bb[2] - bb[0]) <= max_width:
                    break
            except Exception:
                break
            s -= 1
            f = font(s, bold)
        d.text(xy, txt, font=f, fill=fill)

    # Same proven invoice/quotation-style header container
    d.rounded_rectangle(rect((2, 2, BW-3, BH-3)), radius=sc(7), outline=navy, width=sc(2), fill=white)
    d.rectangle(rect((2, 2, BW-3, 72)), fill=navy)

    d.text((sc(32), sc(25)), company_name, font=font(31, True), fill=white)
    doc_title = "DELIVERY CHALLAN"
    tf = font(31, True)
    tb = d.textbbox((0, 0), doc_title, font=tf)
    d.text((W - sc(32) - (tb[2] - tb[0]), sc(25)), doc_title, font=tf, fill=white)

    # Logo
    logo_box = rect((68, 108, 288, 308))
    logo_path = first_existing_logo()
    if logo_path and Path(logo_path).exists():
        try:
            logo = _crop_logo_padding(Image.open(logo_path))
            target_w = logo_box[2] - logo_box[0]
            target_h = logo_box[3] - logo_box[1]
            scale = min(target_w / logo.width, target_h / logo.height)
            logo = logo.resize((max(1, int(logo.width * scale)), max(1, int(logo.height * scale))), Image.LANCZOS)
            lx = logo_box[0] + (target_w - logo.width) // 2
            ly = logo_box[1] + (target_h - logo.height) // 2
            img.paste(logo, (lx, ly), logo)
        except Exception:
            d.text((sc(110), sc(190)), "SMS", font=font(62, True), fill=navy)
    else:
        d.text((sc(110), sc(190)), "SMS", font=font(62, True), fill=navy)

    # Challan details card position first so the company block can be fitted before it.
    card_w, card_h = sc(440), sc(230)
    card_x, card_y = W - sc(72) - card_w, sc(88)

    # Company block - fitted to available width before details card.
    cx = sc(335)
    max_company_width = card_x - cx - sc(42)
    draw_fit(company_name, (cx, sc(102)), max_company_width, 40, True, navy, 26)
    draw_fit(tagline, (cx, sc(154)), max_company_width, 22, False, muted, 16)
    draw_fit(address1, (cx, sc(198)), max_company_width, 20, False, text_col, 15)
    draw_fit(address2, (cx, sc(233)), max_company_width, 20, False, text_col, 15)
    draw_fit(f"GSTIN: {gst} | State: {state}, Code: {code}", (cx, sc(268)), max_company_width, 19, False, text_col, 14)
    draw_fit(f"Email: {email} | Phone: {phone}", (cx, sc(303)), max_company_width, 18, False, text_col, 13)

    # Challan details card
    d.rounded_rectangle((card_x, card_y, card_x + card_w, card_y + card_h), radius=sc(6), outline=blue, width=sc(2), fill=soft)
    d.rectangle((card_x, card_y, card_x + card_w, card_y + sc(58)), fill=navy)

    head = "CHALLAN DETAILS"
    hf = font(27, True)
    hb = d.textbbox((0, 0), head, font=hf)
    d.text((card_x + (card_w - (hb[2]-hb[0]))//2, card_y + sc(17)), head, font=hf, fill=white)

    rows = [
        ("Challan No.", clean(challan.get("Challan No.")) or "-"),
        ("Date", clean(date_val) or "-"),
        ("Ref./Order", clean(challan.get("Ref./Order No.")) or "-"),
    ]
    y = card_y + sc(88)
    for label, value in rows:
        d.text((card_x + sc(38), y), label, font=font(20, True), fill=muted)
        d.text((card_x + sc(205), y), ":", font=font(20, True), fill=text_col)
        d.text((card_x + sc(235), y), value, font=font(20, True if label == "Challan No." else False), fill=text_col)
        y += sc(48)

    img.save(out_path)
    return out_path

def _docx_header(doc, company, challan):
    """Add a stable invoice-style header image to the Challan DOCX.

    This replaces the previous native-table DOCX header that distorted/overlapped in Word.
    """
    section = doc.sections[0]
    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.34)
    section.right_margin = Inches(0.34)

    import tempfile
    header_path = Path(tempfile.gettempdir()) / f"sms_challan_docx_header_{os.getpid()}_{safe_filename(clean(challan.get('Challan No.')))}.png"
    rendered = _render_challan_docx_header_image(company, challan, header_path)

    if rendered and Path(rendered).exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        p.add_run().add_picture(str(rendered), width=Inches(7.25))
        try:
            Path(rendered).unlink(missing_ok=True)
        except Exception:
            pass
        return

    # Fallback only if Pillow is unavailable
    company_name = clean(company.get("Company Name")) or "SMS Controls & Automation"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _docx_run(p, company_name, 16, True, "0D344F")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _docx_run(p, "DELIVERY CHALLAN", 14, True, "0D344F")

def _docx_party_box(cell, title, customer, prefix):
    _docx_set_cell_border(cell, "D5E0EC", "6")
    _docx_set_cell_margins(cell, 90, 110, 90, 110)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(3)
    _docx_run(p, title, 8.5, True, "0D344F")
    name = clean(customer.get(f"{prefix} Name")) or clean(customer.get("Customer Name"))
    lines = [
        name,
        clean(customer.get(f"{prefix} Address 1")),
        clean(customer.get(f"{prefix} Address 2")),
        clean(customer.get(f"{prefix} State")),
        f"GSTIN: {clean(customer.get(f'{prefix} GSTIN'))}",
        f"State: {clean(customer.get(f'{prefix} State Code'))}",
    ]
    for i, line in enumerate(lines):
        if not line or line.endswith(": "):
            continue
        pp = cell.add_paragraph()
        pp.paragraph_format.space_after = Pt(0)
        _docx_run(pp, line, 8.0 if i == 0 else 7.6, i == 0, "0B2545")

def _docx_build_challan(docx_path, company, customer, challan, items, settings):
    doc = Document()
    _docx_header(doc, company, challan)

    doc.add_paragraph("")

    party = doc.add_table(rows=1, cols=2)
    party.alignment = WD_TABLE_ALIGNMENT.CENTER
    _docx_set_table_width(party, 7.59)
    _docx_party_box(party.rows[0].cells[0], "BILL TO", customer, "Bill To")
    _docx_party_box(party.rows[0].cells[1], "SHIP TO", customer, "Ship To")

    doc.add_paragraph("")

    item_table = doc.add_table(rows=1, cols=6)
    item_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _docx_set_table_width(item_table, 7.59)
    widths = [0.58, 3.40, 0.94, 0.68, 0.68, 1.31]
    headers = ["S. No.", "Description of Goods", "HSN/SAC", "Qty", "Unit", "Remarks"]
    for i, h in enumerate(headers):
        item_table.columns[i].width = Inches(widths[i])
        _docx_cell_text(item_table.cell(0, i), h, 8.0, True, "FFFFFF", WD_ALIGN_PARAGRAPH.CENTER)
        _docx_set_cell_shading(item_table.cell(0, i), "0D344F")
        _docx_set_cell_border(item_table.cell(0, i), "D5E0EC", "6")
    total_qty = 0.0
    unit = ""
    for idx, it in enumerate(items, 1):
        row = item_table.add_row().cells
        qty = it.get("Qty", "")
        try:
            total_qty += float(qty)
        except Exception:
            pass
        unit = clean(it.get("Unit")) or unit
        vals = [
            clean(it.get("S.No")) or str(idx),
            clean(it.get("Description")),
            clean(it.get("HSN/SAC")) or clean(it.get("HSN")),
            clean(qty),
            clean(it.get("Unit")),
            clean(it.get("Remarks")) or "-",
        ]
        for i, val in enumerate(vals):
            row[i].width = Inches(widths[i])
            _docx_cell_text(row[i], val, 7.8, False, "0B2545", WD_ALIGN_PARAGRAPH.CENTER if i != 1 else WD_ALIGN_PARAGRAPH.LEFT)
            _docx_set_cell_border(row[i], "D5E0EC", "6")

    total = item_table.add_row().cells
    vals = ["", "TOTAL", "", clean(int(total_qty) if total_qty and total_qty.is_integer() else total_qty), unit, ""]
    for i, val in enumerate(vals):
        _docx_cell_text(total[i], val, 7.8, i == 1, "0B2545", WD_ALIGN_PARAGRAPH.CENTER if i != 1 else WD_ALIGN_PARAGRAPH.LEFT)
        _docx_set_cell_border(total[i], "D5E0EC", "6")
        _docx_set_cell_shading(total[i], "F7FAFD")

    doc.add_paragraph("")

    remarks_tbl = doc.add_table(rows=1, cols=2)
    remarks_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _docx_set_table_width(remarks_tbl, 7.59)
    for cell in remarks_tbl.rows[0].cells:
        _docx_set_cell_border(cell, "D5E0EC", "6")
        _docx_set_cell_margins(cell, 100, 120, 100, 120)
    left, right = remarks_tbl.rows[0].cells
    _docx_run(left.paragraphs[0], "REMARKS", 8.3, True, "0D344F")
    remarks = clean(challan.get("Remarks")) or "There is no commercial value involved in this transaction."
    pp = left.add_paragraph()
    _docx_run(pp, remarks, 7.8, False, "0B2545")
    pp = left.add_paragraph()
    _docx_run(pp, f"Delivery: {clean(challan.get('Terms of Delivery')) or '-'}", 7.8, False, "0B2545")

    _docx_run(right.paragraphs[0], "MATERIAL MOVEMENT DETAILS", 8.3, True, "0D344F")
    total_amount = 0.0
    for it in items:
        try:
            total_amount += float(it.get("Amount", 0))
        except Exception:
            pass
    pp = right.add_paragraph()
    _docx_run(pp, f"Material Value (For Reference Only): Rs. {money(total_amount)}", 7.8, False, "0B2545")
    pp = right.add_paragraph()
    _docx_run(pp, clean(settings.get("Amount In Words")) or "INR: Sixty-Three Thousand Six Hundred Rupees Only", 7.8, True, "0B2545")

    doc.add_paragraph("")

    sig = doc.add_table(rows=1, cols=2)
    sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    _docx_set_table_width(sig, 7.59)
    for cell in sig.rows[0].cells:
        _docx_set_cell_border(cell, "D5E0EC", "6")
        _docx_set_cell_margins(cell, 160, 130, 160, 130)
    _docx_run(sig.rows[0].cells[0].paragraphs[0], "PREPARED BY", 8.5, True, "0D344F")
    for label in ["Name: ____________________", "Date: _____________________"]:
        pp = sig.rows[0].cells[0].add_paragraph()
        _docx_run(pp, label, 7.8, False, "0B2545")
    _docx_run(sig.rows[0].cells[1].paragraphs[0], "FOR SMS CONTROLS & AUTOMATION", 8.5, True, "0D344F")
    sig.rows[0].cells[1].add_paragraph("")
    sig.rows[0].cells[1].add_paragraph("")
    pp = sig.rows[0].cells[1].add_paragraph()
    pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _docx_run(pp, "Authorized Signatory", 7.8, False, "0B2545")

    section = doc.sections[0]
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _docx_run(footer, f"This is a Computer Generated Delivery Challan | SMS Controls & Automation | {clean(company.get('Email'))}", 7.1, False, "5B6B82")

    doc.save(docx_path)
    return Path(docx_path)

def generate_challan(input_file, output_dir=None):
    input_file = Path(input_file)
    normalized = normalize_input_file(input_file)
    company, customer, challan, items, settings = read_template(normalized)
    out_dir = Path(output_dir) if output_dir else Path(__file__).parent / "output"
    out_dir.mkdir(parents=True, exist_ok=True)
    challan_no = safe_filename(challan.get("Challan No.") or "Challan")
    customer_name = safe_filename(customer.get("Customer Name") or customer.get("Bill To Name") or "Customer")
    base = out_dir / f"DC-{challan_no}_{customer_name}_Delivery_Challan"

    pdf_path = base.with_suffix(".pdf")
    docx_path = base.with_suffix(".docx")

    pdf = ChallanPDF(pdf_path, company, customer, challan, items, settings)
    pdf.build()

    _docx_build_challan(docx_path, company, customer, challan, items, settings)

    return [pdf_path, docx_path]

if __name__ == "__main__":
    logger = setup_logger("delivery_challan_generator")
    try:
        if len(sys.argv) > 1:
            input_file = Path(sys.argv[1])
        else:
            input_file = Path(__file__).parent / "templates" / "SMS_Delivery_Challan_Template.xlsx"
        output_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else None
        results = generate_challan(input_file, output_dir)
        for result in results:
            print(result)
    except Exception:
        logger.exception("Delivery challan generation failed")
        raise
