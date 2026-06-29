from __future__ import annotations

"""SMS Commercial Document Generator - v1.0.0 GA.

Root fixes in this launcher:
1. Cards and buttons now have separate click zones, so the Generate button is not
   hidden behind the larger card click area.
2. Default templates are resolved from the document registry, not from old flat
   template names like Quotation.xlsx / Invoice.xlsx.
3. Template validation API is called correctly. It returns a DocumentDefinition,
   not (valid, errors).
4. Browse Template auto-detects document type and switches selected card when possible.
5. Canvas buttons/cards now give visual press/selection feedback and status text.
"""

import sys
import subprocess
import traceback
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

APP_NAME = "SMS Commercial Document Generator"
APP_VERSION = "1.0.0"
COMPANY_NAME = "SMS Controls & Automation"
PROJECT_ROOT = Path(__file__).resolve().parent

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# UI metadata only. Real templates/output folders come from registry_service.
DOCUMENTS = [
    ("Quotation", "quotation", "#2563EB", "Create professional quotations"),
    ("Tax Invoice", "invoice", "#16A34A", "Generate GST tax invoices"),
    ("Proforma Invoice", "proforma", "#7C3AED", "Create proforma invoices"),
    ("Delivery Challan", "challan", "#F97316", "Generate delivery challans"),
    ("Salary Slip", "salary_slip", "#0D9488", "Create salary slips"),
    ("Purchase Order", "purchase_order", "#B45309", "Create supplier purchase orders"),
]

COL = {
    "navy": "#071D3A",
    "navy2": "#0B2A52",
    "blue": "#2563EB",
    "blue_dark": "#1D4ED8",
    "bg": "#F5F7FB",
    "card": "#FFFFFF",
    "card_hover": "#F8FBFF",
    "border": "#D9E3F0",
    "text": "#0B1F3A",
    "muted": "#64748B",
    "white": "#FFFFFF",
    "green": "#22C55E",
    "green_dark": "#15803D",
    "orange": "#F97316",
    "soft": "#EFF6FF",
    "selected": "#DBEAFE",
    "pressed": "#BFDBFE",
}


def open_path(path: Path):
    try:
        path = Path(path)
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        elif sys.platform.startswith("win"):
            import os
            os.startfile(str(path))
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as exc:
        messagebox.showerror("Open Failed", str(exc))


class ProfessionalDashboard:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.selected = 0
        self.template_file: Path | None = None
        self.output_folder: Path | None = None
        self.last_generated: Path | None = None
        self.status = "Ready"
        self.status_color = COL["green"]
        self.pressed_zone: str | None = None
        # click tuple: x1, y1, x2, y2, command, zone_id
        self.clicks: list[tuple[int, int, int, int, callable, str]] = []
        self.hover_zone: str | None = None

        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("1280x780")
        self.root.minsize(1050, 650)
        self.root.configure(bg=COL["bg"])

        self.canvas = tk.Canvas(self.root, bg=COL["bg"], highlightthickness=0, bd=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<Motion>", self.on_motion)
        self.canvas.bind("<Configure>", lambda _e: self.draw())

        self.load_default_selection()
        self.draw()
        self.root.update_idletasks()
        self.root.update()
        print(f"UI_VISIBLE items={len(self.canvas.find_all())} version={APP_VERSION}", flush=True)

    # ---------- registry helpers ----------
    def document_name(self, idx: int | None = None) -> str:
        return DOCUMENTS[self.selected if idx is None else idx][0]

    def get_document(self, idx: int | None = None):
        from app_core.services.registry_service import get_document
        return get_document(self.document_name(idx))

    def load_default_selection(self):
        doc = self.get_document(self.selected)
        self.template_file = Path(doc.default_template)
        self.output_folder = Path(doc.output_subdir)

    # ---------- drawing primitives ----------
    def rect(self, x1, y1, x2, y2, fill, outline="", width=1):
        return self.canvas.create_rectangle(x1, y1, x2, y2, fill=fill, outline=outline, width=width)

    def text(self, x, y, txt, size=12, fill=None, anchor="nw", weight="normal", width=None):
        return self.canvas.create_text(
            x, y, text=txt, fill=fill or COL["text"], anchor=anchor,
            font=("Arial", size, weight), width=width
        )

    def click_zone(self, x1, y1, x2, y2, cmd, zone_id: str):
        self.clicks.append((x1, y1, x2, y2, cmd, zone_id))

    def button(self, x1, y1, x2, y2, label, cmd, fill=None, zone_id="button"):
        base = fill or COL["blue"]
        active = COL["pressed"] if self.pressed_zone == zone_id else base
        text_color = COL["text"] if self.pressed_zone == zone_id else COL["white"]
        self.rect(x1, y1, x2, y2, active, active)
        self.text((x1 + x2) // 2, (y1 + y2) // 2, label, 12, text_color, "center", "bold")
        self.click_zone(x1, y1, x2, y2, cmd, zone_id)

    # ---------- layout ----------
    def draw(self):
        self.clicks.clear()
        c = self.canvas
        c.delete("all")
        w = max(c.winfo_width(), 1280)
        h = max(c.winfo_height(), 780)

        self.rect(0, 0, w, h, COL["bg"], COL["bg"])
        self.rect(0, 0, 270, h, COL["navy"], COL["navy"])
        self.rect(270, 0, w, 92, "#FFFFFF", "#E2E8F0")
        self.rect(270, h - 42, w, h, COL["navy"], COL["navy"])

        # Brand
        self.rect(28, 28, 78, 78, "#FFFFFF", "#FFFFFF")
        self.text(53, 53, "S", 28, COL["blue"], "center", "bold")
        self.text(92, 30, "SMS CONTROLS", 16, COL["white"], "nw", "bold")
        self.text(92, 55, "& AUTOMATION", 16, COL["white"], "nw", "bold")
        self.text(28, 95, "Commercial Document Generator", 10, "#BFD2EA")

        # Sidebar
        y = 140
        self.nav_item(y, "Dashboard", True, lambda: self.set_status("Dashboard selected"), "nav_dashboard"); y += 48
        for i, d in enumerate(DOCUMENTS):
            self.nav_item(y, d[0], i == self.selected, lambda idx=i: self.select(idx), f"nav_{i}"); y += 48
        y += 20
        self.nav_item(y, "Open Output Folder", False, self.open_output, "nav_output"); y += 48
        self.nav_item(y, "Settings", False, self.settings, "nav_settings"); y += 48
        self.nav_item(y, "About", False, self.about, "nav_about")
        self.rect(22, h - 140, 248, h - 70, COL["navy2"], "#24496F")
        self.text(42, h - 122, f"Version {APP_VERSION}", 11, COL["white"], weight="bold")
        self.text(42, h - 95, f"● {self.status}", 10, self.status_color, weight="bold", width=185)

        # Topbar
        self.text(310, 25, "Welcome!", 24, COL["text"], weight="bold")
        self.text(310, 60, "Generate professional commercial documents quickly and efficiently", 11, COL["muted"])
        self.button(w - 325, 25, w - 205, 62, "Settings", self.settings, "#1E3A5F", "top_settings")
        self.button(w - 190, 25, w - 35, 62, "Open Output", self.open_output, COL["blue"], "top_output")

        # Content
        self.text(310, 120, "DOCUMENTS", 12, COL["blue"], weight="bold")
        self.rect(410, 130, w - 385, 131, COL["border"], COL["border"])

        start_x, start_y = 310, 160
        card_w, card_h, gap = 235, 210, 20
        for i, _d in enumerate(DOCUMENTS):
            row = 0 if i < 3 else 1
            col = i if i < 3 else i - 3
            x = start_x + col * (card_w + gap)
            y = start_y + row * (card_h + gap)
            self.doc_card(i, x, y, card_w, card_h)

        rx = w - 335
        self.text(rx, 120, "QUICK ACTIONS", 12, COL["blue"], weight="bold")
        self.action(rx, 160, "Browse Template", "Select Excel template", self.browse_template, "action_browse", COL["blue"])
        self.action(rx, 235, "Select Output Folder", "Choose output location", self.browse_output, "action_output", "#0D9488")
        self.action(rx, 310, "Generate Document", "Create selected document", self.generate_selected, "action_generate", COL["green_dark"])
        self.action(rx, 385, "Preview Last Document", "Open last generated file", self.preview_last, "action_preview", "#7C3AED")

        self.text(rx, 485, "SELECTED DOCUMENT", 12, COL["blue"], weight="bold")
        self.rect(rx, 515, w - 35, 640, COL["card"], COL["border"])
        name, _key, color, _desc = DOCUMENTS[self.selected]
        self.rect(rx + 16, 535, rx + 54, 573, COL["soft"], COL["soft"])
        self.text(rx + 35, 554, name[0], 16, color, "center", "bold")
        self.text(rx + 68, 532, name, 15, COL["text"], weight="bold")
        self.text(rx + 18, 580, f"Template: {Path(self.template_file).name if self.template_file else '-'}", 9, COL["muted"], width=275)
        self.text(rx + 18, 607, f"Output: {Path(self.output_folder).name if self.output_folder else '-'}", 9, COL["muted"], width=275)

        self.text(310, h - 28, f"● {self.status}", 10, self.status_color, weight="bold", width=240)
        self.text(560, h - 28, "Templates: loaded", 10, COL["white"])
        self.text(730, h - 28, f"Version: {APP_VERSION}", 10, COL["white"])
        self.text(w - 240, h - 28, datetime.now().strftime("%d %b %Y  •  %I:%M %p"), 10, COL["white"])

    def nav_item(self, y, label, selected, cmd, zone_id):
        bg = COL["blue"] if selected else ("#0A2446" if self.hover_zone == zone_id else COL["navy"])
        self.rect(14, y, 256, y + 38, bg, bg)
        self.text(38, y + 19, "■", 10, COL["white"], "center", "bold")
        self.text(70, y + 19, label, 11, COL["white"], "w", "bold" if selected else "normal")
        self.click_zone(14, y, 256, y + 38, cmd, zone_id)

    def doc_card(self, idx, x, y, cw, ch):
        name, _key, color, desc = DOCUMENTS[idx]
        zone_id = f"card_{idx}"
        border = color if idx == self.selected else COL["border"]
        fill = COL["selected"] if self.pressed_zone == zone_id else (COL["card_hover"] if self.hover_zone == zone_id else COL["card"])
        self.rect(x, y, x + cw, y + ch, fill, border, 2 if idx == self.selected else 1)
        # Add card zone FIRST so the button zone added later wins when reversed.
        self.click_zone(x, y, x + cw, y + ch, lambda i=idx: self.select(i), zone_id)
        self.rect(x + 25, y + 25, x + 95, y + 95, "#EFF6FF", "#EFF6FF")
        self.text(x + 60, y + 60, name[0], 28, color, "center", "bold")
        self.text(x + 25, y + 115, name, 15, COL["text"], weight="bold")
        self.text(x + 25, y + 145, desc, 10, COL["muted"], width=180)
        self.button(x + 25, y + 170, x + 160, y + 202, "Generate  →", lambda i=idx: self.generate(i), color, f"generate_{idx}")

    def action(self, x, y, title, sub, cmd, zone_id, color):
        fill = COL["pressed"] if self.pressed_zone == zone_id else (COL["card_hover"] if self.hover_zone == zone_id else COL["card"])
        self.rect(x, y, x + 300, y + 58, fill, COL["border"])
        self.rect(x, y, x + 6, y + 58, color, color)
        self.rect(x + 18, y + 14, x + 48, y + 44, "#EFF6FF", "#EFF6FF")
        self.text(x + 33, y + 29, title[0], 12, color, "center", "bold")
        self.text(x + 62, y + 14, title, 11, COL["text"], weight="bold")
        self.text(x + 62, y + 35, sub, 9, COL["muted"])
        self.text(x + 280, y + 29, "›", 20, color, "center", "bold")
        self.click_zone(x, y, x + 300, y + 58, cmd, zone_id)

    # ---------- event handling ----------
    def hit(self, x: int, y: int):
        for x1, y1, x2, y2, cmd, zone_id in reversed(self.clicks):
            if x1 <= x <= x2 and y1 <= y <= y2:
                return cmd, zone_id
        return None, None

    def on_click(self, event):
        cmd, zone_id = self.hit(event.x, event.y)
        if not cmd:
            return
        self.pressed_zone = zone_id
        self.draw()
        self.root.update_idletasks()
        self.root.after(120, self.clear_press)
        cmd()

    def clear_press(self):
        self.pressed_zone = None
        self.draw()

    def on_motion(self, event):
        _cmd, zone_id = self.hit(event.x, event.y)
        if zone_id != self.hover_zone:
            self.hover_zone = zone_id
            self.canvas.configure(cursor="hand2" if zone_id else "")
            self.draw()

    # ---------- actions ----------
    def set_status(self, text: str, color: str | None = None):
        self.status = text
        self.status_color = color or COL["green"]
        self.draw()
        self.root.update_idletasks()

    def select(self, idx):
        self.selected = idx
        self.load_default_selection()
        self.set_status(f"Selected {self.document_name(idx)}")

    def browse_template(self):
        file = filedialog.askopenfilename(
            title="Select Excel Template",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("All", "*.*")]
        )
        if not file:
            return
        self.template_file = Path(file)
        try:
            from app_core.services.template_service import TemplateService
            detected = TemplateService().infer_document_type(self.template_file)
            if detected:
                for idx, item in enumerate(DOCUMENTS):
                    if item[0] == detected:
                        self.selected = idx
                        doc = self.get_document(idx)
                        self.output_folder = Path(doc.output_subdir)
                        self.set_status(f"Template selected: {detected}")
                        return
            self.set_status("Template selected")
        except Exception:
            self.set_status("Template selected")

    def browse_output(self):
        initial = str(self.output_folder or PROJECT_ROOT / "output")
        folder = filedialog.askdirectory(title="Select Output Folder", initialdir=initial)
        if folder:
            self.output_folder = Path(folder)
            self.output_folder.mkdir(parents=True, exist_ok=True)
            self.set_status("Output folder selected")

    def open_output(self):
        if self.output_folder is None:
            self.load_default_selection()
        Path(self.output_folder).mkdir(parents=True, exist_ok=True)
        self.set_status("Opening output folder")
        open_path(Path(self.output_folder))

    def settings(self):
        messagebox.showinfo(
            "Settings",
            f"{APP_NAME}\nVersion {APP_VERSION}\n\nSelected: {self.document_name()}\n"
            f"Template:\n{self.template_file}\n\nOutput folder:\n{self.output_folder}"
        )

    def about(self):
        messagebox.showinfo("About", f"{APP_NAME}\nVersion {APP_VERSION}\n{COMPANY_NAME}")

    def generate_selected(self):
        self.generate(self.selected)

    def generate(self, idx):
        # Fix: card Generate must always use the document represented by that card.
        # Previously a previously selected/browsed template was reused, so clicking
        # Invoice after Proforma could still generate Proforma.
        if idx != self.selected:
            self.selected = idx
            self.load_default_selection()
        elif self.template_file is None:
            self.load_default_selection()

        name = self.document_name(idx)
        try:
            from app_core.services.template_service import TemplateService
            from app_core.services.generation_service import DocumentGenerationService

            document = self.get_document(idx)
            template = Path(self.template_file) if self.template_file else Path(document.default_template)
            output = Path(self.output_folder) if self.output_folder else Path(document.output_subdir)

            if not template.exists():
                messagebox.showerror("Template Missing", f"Template file not found:\n{template}")
                self.set_status("Template missing", "#DC2626")
                return

            self.set_status(f"Generating {name}...", COL["orange"])
            self.root.update()

            # Correct API usage: validate returns the matching DocumentDefinition or raises.
            validated_document = TemplateService().validate(document, template)
            if validated_document.display_name != document.display_name:
                # Auto-switch if selected template is for another document type.
                for auto_idx, item in enumerate(DOCUMENTS):
                    if item[0] == validated_document.display_name:
                        self.selected = auto_idx
                        name = item[0]
                        break

            output.mkdir(parents=True, exist_ok=True)
            generated = [Path(p) for p in DocumentGenerationService().generate(validated_document, template, output) if p]
            self.last_generated = generated[0] if generated else None
            self.template_file = template
            self.output_folder = output
            self.set_status(f"Generated {name}", COL["green"])
            msg = "Generated successfully:\n\n" + "\n".join(str(p) for p in generated)
            messagebox.showinfo("Generated", msg)
        except Exception as exc:
            print(traceback.format_exc(), file=sys.stderr, flush=True)
            self.set_status("Generation failed", "#DC2626")
            messagebox.showerror("Generation Failed", f"{type(exc).__name__}: {exc}")

    def preview_last(self):
        if self.last_generated and self.last_generated.exists():
            self.set_status("Opening last document")
            open_path(self.last_generated)
        else:
            messagebox.showinfo("Preview", "No generated document available yet.")


def main():
    print(f"APP_START file={Path(__file__).resolve()} cwd={Path.cwd()} python={sys.version.split()[0]} version={APP_VERSION}", flush=True)
    root = tk.Tk()
    ProfessionalDashboard(root)
    root.mainloop()


if __name__ == "__main__":
    main()
