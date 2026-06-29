
from pathlib import Path
import os
import platform
import subprocess
import tempfile
import shutil
try:
    from app_core.config.settings import BASE_DIR
except Exception:
    BASE_DIR = Path(__file__).parent.resolve()

def is_mac():
    return platform.system().lower() == "darwin"

def is_windows():
    return platform.system().lower().startswith("win")

def open_folder(path):
    path = str(Path(path).resolve())
    if is_mac():
        subprocess.run(["open", path], check=False)
    elif is_windows():
        os.startfile(path)
    else:
        subprocess.run(["xdg-open", path], check=False)

def open_file(path):
    path = str(Path(path).resolve())
    if is_mac():
        subprocess.run(["open", path], check=False)
    elif is_windows():
        os.startfile(path)
    else:
        subprocess.run(["xdg-open", path], check=False)

def find_libreoffice():
    candidates = []
    # Optional bundled LibreOffice locations used by packaged builds.
    bundled = BASE_DIR / "vendor" / "libreoffice"
    candidates += [
        str(bundled / "program" / "soffice.exe"),
        str(bundled / "program" / "soffice"),
        str(bundled / "LibreOffice.app" / "Contents" / "MacOS" / "soffice"),
    ]
    if is_mac():
        candidates += [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/opt/homebrew/bin/soffice",
            "/usr/local/bin/soffice",
        ]
    elif is_windows():
        candidates += [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    candidates += ["soffice", "libreoffice"]
    for c in candidates:
        p = shutil.which(c) if c in ("soffice", "libreoffice") else c
        if p and Path(p).exists():
            return p
    return None

def convert_numbers_to_xlsx(input_path, work_dir=None):
    """Convert Apple Numbers file to XLSX using Numbers automation on macOS."""
    input_path = Path(input_path).resolve()
    if not is_mac():
        raise RuntimeError(".numbers input is supported only on macOS with Apple Numbers installed.")
    if work_dir is None:
        work_dir = Path(tempfile.mkdtemp(prefix="cdg_numbers_"))
    else:
        work_dir = Path(work_dir)
        work_dir.mkdir(parents=True, exist_ok=True)
    output_path = work_dir / (input_path.stem + ".xlsx")

    script = """
set inputFile to POSIX file "{input_file}"
set outputFile to POSIX file "{output_file}"
tell application "Numbers"
    set theDoc to open inputFile
    export theDoc to outputFile as Microsoft Excel
    close theDoc saving no
end tell
""".format(input_file=str(input_path).replace('"', '\\"'), output_file=str(output_path).replace('"', '\\"'))

    res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError("Failed to convert .numbers to .xlsx using Apple Numbers. " + (res.stderr or res.stdout))
    if not output_path.exists():
        raise RuntimeError("Numbers conversion completed but XLSX file was not created.")
    return output_path

def convert_xls_to_xlsx(input_path, work_dir=None):
    input_path = Path(input_path).resolve()
    if work_dir is None:
        work_dir = Path(tempfile.mkdtemp(prefix="cdg_xls_"))
    else:
        work_dir = Path(work_dir)
        work_dir.mkdir(parents=True, exist_ok=True)
    output_path = work_dir / (input_path.stem + ".xlsx")

    if is_mac():
        script = """
set inputFile to POSIX file "{input_file}"
set outputFile to POSIX file "{output_file}"
tell application "Numbers"
    set theDoc to open inputFile
    export theDoc to outputFile as Microsoft Excel
    close theDoc saving no
end tell
""".format(input_file=str(input_path).replace('"', '\\"'), output_file=str(output_path).replace('"', '\\"'))
        res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        if res.returncode == 0 and output_path.exists():
            return output_path

    soffice = find_libreoffice()
    if soffice:
        res = subprocess.run([
            soffice, "--headless", "--convert-to", "xlsx", "--outdir", str(work_dir), str(input_path)
        ], capture_output=True, text=True)
        if res.returncode == 0 and output_path.exists():
            return output_path
        candidates = list(work_dir.glob(input_path.stem + "*.xlsx"))
        if candidates:
            return candidates[0]
        raise RuntimeError("LibreOffice XLS conversion failed. " + (res.stderr or res.stdout))

    raise RuntimeError("Unable to convert .xls to .xlsx. Install Apple Numbers on Mac or LibreOffice.")

def normalize_spreadsheet_input(input_path, work_dir=None):
    input_path = Path(input_path)
    ext = input_path.suffix.lower()
    if ext == ".xlsx":
        return input_path
    if ext == ".numbers":
        return convert_numbers_to_xlsx(input_path, work_dir=work_dir)
    if ext == ".xls":
        return convert_xls_to_xlsx(input_path, work_dir=work_dir)
    raise RuntimeError("Unsupported template format. Please select .xlsx, .xls, or .numbers.")


def convert_docx_to_pdf(input_docx, output_dir=None):
    """Convert DOCX to PDF using LibreOffice, or Apple Pages on macOS."""
    input_docx = Path(input_docx).resolve()
    if output_dir is None:
        output_dir = input_docx.parent
    else:
        output_dir = Path(output_dir).resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

    expected_pdf = output_dir / (input_docx.stem + ".pdf")

    # 1. LibreOffice path
    soffice = find_libreoffice()
    if soffice:
        res = subprocess.run([
            soffice, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(input_docx)
        ], capture_output=True, text=True)
        if res.returncode == 0 and expected_pdf.exists():
            return expected_pdf
        candidates = list(output_dir.glob(input_docx.stem + "*.pdf"))
        if candidates:
            return candidates[0]

    # 2. Apple Pages fallback
    if is_mac():
        script = """
set inputFile to POSIX file "{input_file}"
set outputFile to POSIX file "{output_file}"
tell application "Pages"
    set theDoc to open inputFile
    export theDoc to outputFile as PDF
    close theDoc saving no
end tell
""".format(input_file=str(input_docx).replace('"', '\\"'), output_file=str(expected_pdf).replace('"', '\\"'))
        res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        if res.returncode == 0 and expected_pdf.exists():
            return expected_pdf
        raise RuntimeError("DOCX to PDF conversion failed. Install LibreOffice or ensure Apple Pages is installed. " + (res.stderr or res.stdout))

    raise RuntimeError("DOCX to PDF conversion failed. Install LibreOffice to generate PDF from DOCX.")

# Backward-compatible aliases used by the quotation engine.
def normalize_input_file(input_path, logger=None):
    return normalize_spreadsheet_input(input_path)


def setup_logger(name="commercial_document_generator"):
    try:
        from logging_config import setup_logging
        return setup_logging().getChild(name)
    except Exception:
        import logging
        logging.basicConfig(level=logging.INFO)
        return logging.getLogger(name)
