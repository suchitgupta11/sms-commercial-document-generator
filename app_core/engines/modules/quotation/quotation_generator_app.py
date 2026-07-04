
import sys
import subprocess
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from platform_utils import (
    setup_logger,
    detect_environment,
    convert_docx_to_pages_mac,
    convert_docx_to_pdf,
    open_folder,
    is_mac,
    is_windows,
    LOG_DIR,
)

APP_TITLE = "SMS Commercial Quotation Generator"
BASE_DIR = Path(__file__).resolve().parent
GENERATOR = BASE_DIR / "generate_proposal.py"
DEFAULT_TEMPLATE = BASE_DIR / "templates" / "Proposal_Data_Template.xlsx"
OUTPUT_DIR = BASE_DIR / "output"
LOGO_PATH = BASE_DIR / "assets" / "logo_from_doc.png"

# Premium industrial/corporate palette
NAVY = "#041E42"
NAVY_2 = "#0D2742"
BLUE = "#0052CC"
BLUE_PRESS = "#003D99"
TEAL = "#008DA6"
TEAL_PRESS = "#006C80"
GREEN = "#00875A"
GREEN_PRESS = "#006644"
ORANGE = "#FF8B00"
ORANGE_PRESS = "#C76D00"
RED = "#DE350B"
RED_PRESS = "#BF2600"

BG = "#EEF5FF"
CARD = "#FFFFFF"
CARD_ALT = "#F8FBFF"
TEXT = "#172B4D"
MUTED = "#5E6C84"
BORDER = "#C7D8EA"

SOFT_BLUE = "#E6F0FF"
SOFT_GREEN = "#E3FCEF"
SOFT_ORANGE = "#FFF4E5"
SOFT_GREY = "#F4F6F8"
SELECTED = "#DCEBFF"

SUCCESS = "#0F766E"
ERROR = "#B91C1C"

class QuotationGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1060x800")
        self.root.minsize(1020, 760)
        self.root.configure(bg=BG)
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

        self.logger = setup_logger("quotation_generator_app")
        self.logger.info("Application started V10")

        self.input_file = tk.StringVar(value=str(DEFAULT_TEMPLATE if DEFAULT_TEMPLATE.exists() else ""))
        self.status = tk.StringVar(value="Ready to generate quotation.")
        self.is_running = False

        self.out_docx = tk.BooleanVar(value=True)
        self.out_pdf = tk.BooleanVar(value=True)
        self.out_pages = tk.BooleanVar(value=False)

        self.env_text = tk.StringVar(value="Checking system status...")
        self.last_pdf_path = None
        self.output_cards = {}

        self.configure_styles()
        self.build_ui()
        self.refresh_environment()
        self.refresh_output_cards()

    def configure_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Root.TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Header.TFrame", background=NAVY)
        style.configure("Title.TLabel", background=NAVY, foreground="white", font=("Helvetica", 25, "bold"))
        style.configure("Company.TLabel", background=NAVY, foreground="#BFD7FF", font=("Helvetica", 11, "bold"))
        style.configure("HeaderSub.TLabel", background=NAVY, foreground="#D7E8FF", font=("Helvetica", 11))
        style.configure("Section.TLabel", background=CARD, foreground=NAVY, font=("Helvetica", 13, "bold"))
        style.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=("Helvetica", 9))
        style.configure("Status.TLabel", background=BG, foreground=TEXT, font=("Helvetica", 10, "bold"))
        style.configure("Accent.Horizontal.TProgressbar", troughcolor="#D9E6F2", background=BLUE)

    def button(self, parent, text, command, bg, fg="white", active_bg=None, font=("Helvetica", 11, "bold"), padx=16, pady=10, width=None):
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg or bg,
            activeforeground=fg,
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            font=font,
            padx=padx,
            pady=pady,
            width=width,
            cursor="hand2",
        )

    def build_ui(self):
        self.build_header()

        shell = tk.Frame(self.root, bg=BG)
        shell.pack(fill="both", expand=True, padx=28, pady=18)

        left = tk.Frame(shell, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        right = tk.Frame(shell, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        right.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

        shell.columnconfigure(0, weight=3)
        shell.columnconfigure(1, weight=2)
        shell.rowconfigure(0, weight=1)

        self.build_left_panel(left)
        self.build_right_panel(right)

        footer = tk.Frame(self.root, bg=BG)
        footer.pack(fill="x", padx=28, pady=(0, 18))
        tk.Label(footer, textvariable=self.status, bg=BG, fg=TEXT, font=("Helvetica", 10, "bold")).pack(anchor="w")

    def build_header(self):
        header = tk.Frame(self.root, bg=NAVY)
        header.pack(fill="x")

        inner = tk.Frame(header, bg=NAVY)
        inner.pack(fill="x", padx=28, pady=22)

        logo_frame = tk.Frame(inner, bg=NAVY)
        logo_frame.pack(side="left", padx=(0, 18))

        self.logo_img = None
        try:
            if LOGO_PATH.exists():
                self.logo_img = tk.PhotoImage(file=str(LOGO_PATH))
                w = self.logo_img.width()
                if w > 120:
                    factor = max(1, int(w / 110))
                    self.logo_img = self.logo_img.subsample(factor, factor)
                tk.Label(logo_frame, image=self.logo_img, bg=NAVY).pack()
            else:
                tk.Label(logo_frame, text="SMS", bg=NAVY, fg="white", font=("Helvetica", 22, "bold")).pack()
        except Exception:
            tk.Label(logo_frame, text="SMS", bg=NAVY, fg="white", font=("Helvetica", 22, "bold")).pack()

        text_frame = tk.Frame(inner, bg=NAVY)
        text_frame.pack(side="left", fill="x", expand=True)
        ttk.Label(text_frame, text="SMS Controls & Automation", style="Company.TLabel").pack(anchor="w")
        ttk.Label(text_frame, text="Commercial Quotation Generator", style="Title.TLabel").pack(anchor="w", pady=(2, 0))
        ttk.Label(text_frame, text="Production Edition • Generate premium quotations in DOCX and PDF", style="HeaderSub.TLabel").pack(anchor="w", pady=(5, 0))

        tk.Label(inner, text="V10", bg="#123A5C", fg="white", font=("Helvetica", 12, "bold"), padx=16, pady=8).pack(side="right")

    def section_title(self, parent, title, subtitle=None):
        tk.Label(parent, text=title, bg=CARD, fg=NAVY, font=("Helvetica", 13, "bold")).pack(anchor="w")
        if subtitle:
            tk.Label(parent, text=subtitle, bg=CARD, fg=MUTED, font=("Helvetica", 9), wraplength=620, justify="left").pack(anchor="w", pady=(4, 8))

    def build_left_panel(self, parent):
        pad = tk.Frame(parent, bg=CARD)
        pad.pack(fill="both", expand=True, padx=24, pady=22)

        self.section_title(pad, "1. Select Quotation Data", "Choose the Excel or Apple Numbers file that contains customer, item, pricing and terms data.")

        input_card = tk.Frame(pad, bg=SOFT_BLUE, highlightbackground="#BED4F7", highlightthickness=1)
        input_card.pack(fill="x", pady=(0, 18))

        tk.Label(input_card, text="Quotation Data File", bg=SOFT_BLUE, fg=NAVY, font=("Helvetica", 10, "bold")).pack(anchor="w", padx=14, pady=(12, 4))

        file_row = tk.Frame(input_card, bg=SOFT_BLUE)
        file_row.pack(fill="x", padx=14, pady=(0, 12))

        tk.Entry(
            file_row,
            textvariable=self.input_file,
            bg="white",
            fg=TEXT,
            relief="flat",
            font=("Helvetica", 10),
        ).pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 12))

        # Fixed Browse button with proper full label and width
        self.button(file_row, "📂 Browse File", self.browse_file, BLUE, "white", BLUE_PRESS, width=14, padx=10, pady=9).pack(side="left")

        tk.Label(
            input_card,
            text="Supported: Mac (.xlsx, .xlsm, .numbers)  |  Windows (.xlsx, .xlsm, .xls)",
            bg=SOFT_BLUE,
            fg=MUTED,
            font=("Helvetica", 9),
        ).pack(anchor="w", padx=14, pady=(0, 12))

        self.section_title(pad, "2. Select Output Format", "DOCX and PDF are selected by default. Click any card to change the output.")

        card_row = tk.Frame(pad, bg=CARD)
        card_row.pack(fill="x", pady=(0, 18))
        self.create_output_card(card_row, "docx", "Editable Word Copy", ".docx", "For internal editing and future changes", self.out_docx, 0)
        self.create_output_card(card_row, "pdf", "Customer PDF", ".pdf", "Best format for sharing with customers", self.out_pdf, 1)
        self.create_output_card(card_row, "pages", "Apple Pages", ".pages", "Optional Mac native copy", self.out_pages, 2)

        self.section_title(pad, "3. Generate Quotation", None)

        purpose = tk.Frame(pad, bg=SOFT_GREEN, highlightbackground="#ABF5D1", highlightthickness=1)
        purpose.pack(fill="x", pady=(8, 18))
        tk.Label(purpose, text="What this software does", bg=SOFT_GREEN, fg=GREEN, font=("Helvetica", 10, "bold")).pack(anchor="w", padx=14, pady=(10, 2))
        tk.Label(
            purpose,
            text="It reads quotation data, performs calculations, applies the premium proposal design, and creates the selected output files.",
            bg=SOFT_GREEN,
            fg=TEXT,
            wraplength=620,
            justify="left",
            font=("Helvetica", 10),
        ).pack(anchor="w", padx=14, pady=(0, 10))

        btn_row = tk.Frame(pad, bg=CARD)
        btn_row.pack(fill="x", pady=(2, 0))

        self.generate_btn = self.button(btn_row, "Generate Quotation", self.generate_async, BLUE, "white", BLUE_PRESS, font=("Helvetica", 13, "bold"), padx=24, pady=13)
        self.generate_btn.pack(side="left")

        self.button(btn_row, "Preview PDF", self.preview_pdf, GREEN, "white", GREEN_PRESS, padx=18, pady=13).pack(side="left", padx=(10, 0))
        self.button(btn_row, "Open Output", lambda: open_folder(OUTPUT_DIR), TEAL, "white", TEAL_PRESS, padx=18, pady=13).pack(side="left", padx=(10, 0))

        self.progress = ttk.Progressbar(pad, mode="indeterminate", style="Accent.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(18, 0))

    def create_output_card(self, parent, key, title, ext, desc, var, col):
        frame = tk.Frame(parent, bg=CARD, highlightbackground=BORDER, highlightthickness=1, bd=0, cursor="hand2")
        frame.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 10, 0), ipadx=8, ipady=8)
        parent.columnconfigure(col, weight=1)

        check = tk.Label(frame, text="☐", bg=CARD, fg=MUTED, font=("Helvetica", 20, "bold"))
        check.pack(anchor="w", padx=12, pady=(10, 0))
        title_label = tk.Label(frame, text=title, bg=CARD, fg=NAVY, font=("Helvetica", 12, "bold"))
        title_label.pack(anchor="w", padx=12, pady=(2, 0))
        ext_label = tk.Label(frame, text=ext, bg=CARD, fg=BLUE, font=("Helvetica", 11, "bold"))
        ext_label.pack(anchor="w", padx=12)
        desc_label = tk.Label(frame, text=desc, bg=CARD, fg=MUTED, font=("Helvetica", 9), wraplength=150, justify="left")
        desc_label.pack(anchor="w", padx=12, pady=(2, 12))

        self.output_cards[key] = {"frame": frame, "check": check, "var": var}
        def click(_=None):
            var.set(not var.get())
            self.refresh_output_cards()
        frame.bind("<Button-1>", click)
        for child in frame.winfo_children():
            child.bind("<Button-1>", click)

    def refresh_output_cards(self):
        for key, data in self.output_cards.items():
            selected = data["var"].get()
            frame = data["frame"]
            check = data["check"]
            bg = SELECTED if selected else CARD
            fg = BLUE if selected else MUTED
            frame.configure(bg=bg, highlightbackground=BLUE if selected else BORDER, highlightthickness=2 if selected else 1)
            check.configure(text="☑" if selected else "☐", bg=bg, fg=fg)
            for child in frame.winfo_children():
                try:
                    child.configure(bg=bg)
                except Exception:
                    pass

    def build_right_panel(self, parent):
        pad = tk.Frame(parent, bg=CARD)
        pad.pack(fill="both", expand=True, padx=24, pady=22)

        self.section_title(pad, "System Status", "Simple readiness check for quotation generation and PDF export.")

        status_card = tk.Frame(pad, bg=SOFT_BLUE, highlightbackground="#BED4F7", highlightthickness=1)
        status_card.pack(fill="x", pady=(0, 18))
        self.env_label = tk.Label(status_card, textvariable=self.env_text, bg=SOFT_BLUE, fg=TEXT, justify="left", anchor="w", font=("Helvetica", 10))
        self.env_label.pack(fill="x", padx=14, pady=12)

        self.section_title(pad, "Generation Progress", None)
        self.step_labels = {}
        steps = [
            ("input", "Validate input file"),
            ("docx", "Generate Word quotation"),
            ("pages", "Create Apple Pages copy"),
            ("pdf", "Create customer PDF"),
            ("done", "Complete"),
        ]
        step_box = tk.Frame(pad, bg=SOFT_GREY, highlightbackground="#DFE4EA", highlightthickness=1)
        step_box.pack(fill="x", pady=(8, 18))
        for key, text in steps:
            lbl = tk.Label(step_box, text=f"○ {text}", bg=SOFT_GREY, fg=MUTED, font=("Helvetica", 10), anchor="w")
            lbl.pack(fill="x", padx=14, pady=5)
            self.step_labels[key] = lbl

        self.section_title(pad, "Quick Actions", None)
        tk.Label(
            pad,
            text="Use Preview PDF after generation to open the latest customer PDF. Output files are saved safely in the output folder.",
            bg=CARD,
            fg=MUTED,
            wraplength=350,
            justify="left",
            font=("Helvetica", 9),
        ).pack(anchor="w", pady=(4, 12))

        self.button(pad, "Preview PDF", self.preview_pdf, GREEN, "white", GREEN_PRESS, padx=18, pady=10).pack(anchor="w", fill="x")
        self.button(pad, "Open Output Folder", lambda: open_folder(OUTPUT_DIR), TEAL, "white", TEAL_PRESS, padx=18, pady=10).pack(anchor="w", fill="x", pady=(10, 0))
        self.button(pad, "Exit", self.close_app, RED, "white", RED_PRESS, padx=18, pady=10).pack(anchor="w", fill="x", pady=(12, 0))

    def refresh_environment(self):
        env = detect_environment()
        lines = [f"✓ Operating System: {env['os']}"]
        if is_mac():
            lines += [
                f"{'✓' if env['numbers'] else '⚠'} Apple Numbers input",
                f"{'✓' if env['pages'] else '⚠'} Apple Pages PDF export",
                f"{'✓' if env['libreoffice'] else '○'} LibreOffice fallback",
            ]
        elif is_windows():
            lines += [
                f"{'✓' if env['microsoft_excel'] else '⚠'} Microsoft Excel handling",
                f"{'✓' if env['microsoft_word'] else '⚠'} Microsoft Word PDF export",
                f"{'✓' if env['libreoffice'] else '○'} LibreOffice fallback",
            ]
            self.out_pages.set(False)
        else:
            lines += [f"{'✓' if env['libreoffice'] else '⚠'} LibreOffice fallback"]
            self.out_pages.set(False)

        self.env_text.set("\n".join(lines))
        self.logger.info("Environment status: %s", env)

    def browse_file(self):
        filetypes = [
            ("Supported Files", "*.xlsx *.xlsm *.xls *.numbers"),
            ("Excel Files", "*.xlsx *.xlsm *.xls"),
            ("Apple Numbers Files", "*.numbers"),
            ("All Files", "*.*"),
        ]
        path = filedialog.askopenfilename(title="Select Quotation Data File", filetypes=filetypes)
        if path:
            self.input_file.set(path)
            self.logger.info("Input file selected: %s", path)

    def mark_step(self, key, state):
        if key not in self.step_labels:
            return
        prefix = {"pending": "○", "running": "●", "done": "✓", "fail": "✗"}.get(state, "○")
        colors = {"pending": MUTED, "running": BLUE, "done": SUCCESS, "fail": ERROR}
        text = self.step_labels[key].cget("text")
        label_text = text[2:] if len(text) > 2 else text
        self.step_labels[key].configure(text=f"{prefix} {label_text}", fg=colors.get(state, MUTED))

    def reset_steps(self):
        for key in self.step_labels:
            self.mark_step(key, "pending")

    def set_running(self, running):
        self.is_running = running
        if running:
            self.generate_btn.configure(state="disabled", text="Generating...")
            self.progress.start(12)
        else:
            self.generate_btn.configure(state="normal", text="Generate Quotation")
            self.progress.stop()

    def generate_async(self):
        if self.is_running:
            return
        threading.Thread(target=self.generate, daemon=True).start()

    def generate(self):
        input_path = Path(self.input_file.get().strip())
        if not input_path.exists():
            self.root.after(0, lambda: messagebox.showerror("Missing File", "Please select a valid quotation data file."))
            return
        if not (self.out_docx.get() or self.out_pdf.get() or self.out_pages.get()):
            self.root.after(0, lambda: messagebox.showerror("Output Required", "Please select at least one output format."))
            return

        self.root.after(0, lambda: self.set_running(True))
        self.root.after(0, self.reset_steps)
        self.status.set("Generating quotation...")
        self.logger.info("Generation started: %s", input_path)

        generated_files = []
        warnings = []
        self.last_pdf_path = None

        try:
            self.root.after(0, lambda: self.mark_step("input", "done"))
            self.root.after(0, lambda: self.mark_step("docx", "running"))
            result = subprocess.run([sys.executable, str(GENERATOR), str(input_path)], cwd=str(BASE_DIR), capture_output=True, text=True)
            if result.returncode != 0:
                raise RuntimeError(result.stderr or result.stdout or "Quotation generation failed.")

            output_docx = Path(result.stdout.strip().splitlines()[-1])
            if not output_docx.exists():
                raise RuntimeError("DOCX generation completed but output file was not found.")

            self.root.after(0, lambda: self.mark_step("docx", "done"))
            self.logger.info("DOCX generated: %s", output_docx)

            if self.out_docx.get():
                generated_files.append(output_docx)

            if self.out_pages.get():
                self.root.after(0, lambda: self.mark_step("pages", "running"))
                try:
                    pages_file = convert_docx_to_pages_mac(output_docx, logger=self.logger)
                    generated_files.append(pages_file)
                    self.root.after(0, lambda: self.mark_step("pages", "done"))
                except Exception as e:
                    warnings.append("PAGES export failed:\n" + str(e))
                    self.logger.exception("PAGES export failed")
                    self.root.after(0, lambda: self.mark_step("pages", "fail"))

            if self.out_pdf.get():
                self.root.after(0, lambda: self.mark_step("pdf", "running"))
                try:
                    pdf_file = convert_docx_to_pdf(output_docx, logger=self.logger)
                    self.last_pdf_path = Path(pdf_file)
                    generated_files.append(pdf_file)
                    self.root.after(0, lambda: self.mark_step("pdf", "done"))
                except Exception as e:
                    warnings.append("PDF export failed:\n" + str(e))
                    self.logger.exception("PDF export failed")
                    self.root.after(0, lambda: self.mark_step("pdf", "fail"))

            if not generated_files:
                generated_files.append(output_docx)
                warnings.append("Only DOCX was generated as fallback.")

            self.root.after(0, lambda: self.mark_step("done", "done"))

            def success_ui():
                self.set_running(False)
                self.status.set("Completed with warnings." if warnings else "Quotation generated successfully.")
                message = "Generated files:\n\n" + "\n".join(str(p) for p in generated_files)
                if warnings:
                    message += "\n\nWarnings:\n\n" + "\n\n".join(warnings)
                    messagebox.showwarning("Completed with Warnings", message)
                else:
                    messagebox.showinfo("Success", message)

            self.root.after(0, success_ui)

        except Exception as e:
            self.logger.exception("Generation failed")
            self.root.after(0, lambda: self.mark_step("done", "fail"))
            def error_ui():
                self.set_running(False)
                self.status.set("Failed. Please check logs.")
                messagebox.showerror("Error", str(e))
            self.root.after(0, error_ui)

    def preview_pdf(self):
        pdf_path = self.last_pdf_path
        if not pdf_path or not Path(pdf_path).exists():
            pdfs = sorted(list(OUTPUT_DIR.glob("*.pdf")), key=lambda p: p.stat().st_mtime, reverse=True)
            if pdfs:
                pdf_path = pdfs[0]
        if not pdf_path or not Path(pdf_path).exists():
            messagebox.showinfo("Preview PDF", "No PDF file found yet. Generate a quotation with Customer PDF selected first.")
            return
        try:
            if is_mac():
                subprocess.run(["open", str(pdf_path)])
            elif is_windows():
                import os
                os.startfile(str(pdf_path))
            else:
                subprocess.run(["xdg-open", str(pdf_path)])
            self.status.set(f"Preview opened: {Path(pdf_path).name}")
        except Exception as e:
            messagebox.showerror("Preview Failed", str(e))

    def close_app(self):
        if self.is_running and not messagebox.askyesno("Generation Running", "Generation is still running. Do you want to close the app?"):
            return
        self.logger.info("Application closed")
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    QuotationGeneratorApp(root)
    root.mainloop()
