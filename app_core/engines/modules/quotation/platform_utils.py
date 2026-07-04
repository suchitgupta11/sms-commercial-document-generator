
import os
import sys
import shutil
import platform
import subprocess
import tempfile
import logging
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"
OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

def setup_logger(name="proposal_app"):
    log_file = LOG_DIR / f"app_{datetime.now().strftime('%Y%m%d')}.log"
    error_file = LOG_DIR / f"error_{datetime.now().strftime('%Y%m%d')}.log"

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s")

    info_handler = logging.FileHandler(log_file, encoding="utf-8")
    info_handler.setLevel(logging.INFO)
    info_handler.setFormatter(formatter)

    error_handler = logging.FileHandler(error_file, encoding="utf-8")
    error_handler.setLevel(logging.ERROR)
    error_handler.setFormatter(formatter)

    logger.addHandler(info_handler)
    logger.addHandler(error_handler)
    return logger

def get_os_name():
    return platform.system()

def is_mac():
    return platform.system().lower() == "darwin"

def is_windows():
    return platform.system().lower() == "windows"

def app_exists_mac(app_name):
    """
    Robust macOS app detection.
    Checks common app paths, Spotlight metadata, and Launch Services via `open -Ra`.
    """
    if not is_mac():
        return False

    common_paths = [
        Path(f"/Applications/{app_name}.app"),
        Path(f"/System/Applications/{app_name}.app"),
        Path.home() / "Applications" / f"{app_name}.app",
    ]

    for p in common_paths:
        if p.exists():
            return True

    try:
        result = subprocess.run(
            ["mdfind", f"kMDItemCFBundleIdentifier == 'com.apple.iWork.{app_name}'"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0 and result.stdout.strip():
            return True
    except Exception:
        pass

    try:
        result = subprocess.run(
            ["open", "-Ra", app_name],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return True
    except Exception:
        pass

    return False

def find_libreoffice():
    candidates = []
    if is_mac():
        candidates.extend([
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/libreoffice",
        ])
    elif is_windows():
        candidates.extend([
            r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
            r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
        ])
    else:
        candidates.extend(["libreoffice", "soffice"])

    for c in candidates:
        if Path(c).exists() or shutil.which(c):
            return c
    return shutil.which("libreoffice") or shutil.which("soffice")

def has_microsoft_word():
    if not is_windows():
        return False
    try:
        import win32com.client
        word = win32com.client.DispatchEx("Word.Application")
        word.Quit()
        return True
    except Exception:
        return False

def has_microsoft_excel():
    if not is_windows():
        return False
    try:
        import win32com.client
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Quit()
        return True
    except Exception:
        return False

def detect_environment():
    return {
        "os": get_os_name(),
        "numbers": is_mac() and app_exists_mac("Numbers"),
        "pages": is_mac() and app_exists_mac("Pages"),
        "libreoffice": bool(find_libreoffice()),
        "microsoft_word": has_microsoft_word(),
        "microsoft_excel": has_microsoft_excel(),
    }

def run_applescript(script_text):
    return subprocess.run(["osascript", "-e", script_text], capture_output=True, text=True)

def convert_numbers_to_xlsx(numbers_path, logger=None):
    if not is_mac():
        raise RuntimeError(".numbers input is supported only on macOS with Apple Numbers installed.")
    if not app_exists_mac("Numbers"):
        raise RuntimeError("Apple Numbers is not installed. Please use .xlsx/.xlsm input or install Numbers.")

    numbers_path = Path(numbers_path).resolve()
    if not numbers_path.exists():
        raise RuntimeError(f"Numbers file not found: {numbers_path}")

    temp_dir = Path(tempfile.mkdtemp(prefix="proposal_numbers_"))
    xlsx_out = temp_dir / (numbers_path.stem + ".xlsx")

    applescript = (
        f'set inputFile to POSIX file "{numbers_path}" as alias\n'
        f'set outputFile to POSIX file "{xlsx_out}"\n'
        'tell application "Numbers"\n'
        '    open inputFile\n'
        '    delay 1\n'
        '    set theDoc to front document\n'
        '    export theDoc to outputFile as Microsoft Excel\n'
        '    close theDoc saving no\n'
        'end tell\n'
    )

    if logger:
        logger.info("Converting Numbers input to XLSX: %s", numbers_path)

    result = run_applescript(applescript)
    if result.returncode != 0:
        raise RuntimeError("Numbers conversion failed. Allow macOS Automation permission if prompted.\n" + (result.stderr or result.stdout))
    if not xlsx_out.exists():
        raise RuntimeError("Numbers conversion completed but XLSX file was not created.")
    return xlsx_out

def convert_xls_to_xlsx_with_excel(xls_path, logger=None):
    if not is_windows():
        raise RuntimeError(".xls conversion with Microsoft Excel is supported on Windows only.")
    try:
        import win32com.client
    except Exception:
        raise RuntimeError("pywin32 is required for .xls conversion on Windows. Install with: python -m pip install pywin32")

    xls_path = Path(xls_path).resolve()
    out_path = xls_path.with_suffix(".xlsx")

    if logger:
        logger.info("Converting XLS to XLSX using Microsoft Excel: %s", xls_path)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(xls_path))
        wb.SaveAs(str(out_path), FileFormat=51)
        wb.Close(False)
    finally:
        excel.Quit()

    if not out_path.exists():
        raise RuntimeError("Excel conversion completed but XLSX file was not created.")
    return out_path

def normalize_input_file(input_path, logger=None):
    input_path = Path(input_path)
    suffix = input_path.suffix.lower()
    if suffix == ".numbers":
        return convert_numbers_to_xlsx(input_path, logger=logger)
    if suffix in (".xlsx", ".xlsm"):
        return input_path
    if suffix == ".xls":
        if is_windows():
            return convert_xls_to_xlsx_with_excel(input_path, logger=logger)
        raise RuntimeError(".xls input is supported on Windows with Microsoft Excel. On Mac, please save as .xlsx or .numbers.")
    raise RuntimeError("Input must be .xlsx, .xlsm, .xls, or .numbers.")

def convert_docx_to_pages_mac(docx_path, logger=None):
    if not is_mac():
        raise RuntimeError(".pages output is supported only on macOS.")
    if not app_exists_mac("Pages"):
        raise RuntimeError("Apple Pages is not installed.")

    docx_path = Path(docx_path).resolve()
    pages_path = docx_path.with_suffix(".pages").resolve()

    applescript = (
        f'set inputFile to POSIX file "{docx_path}" as alias\n'
        f'set outputFile to POSIX file "{pages_path}"\n'
        'tell application "Pages"\n'
        '    open inputFile\n'
        '    delay 1\n'
        '    set theDoc to front document\n'
        '    save theDoc in outputFile\n'
        '    close theDoc saving yes\n'
        'end tell\n'
    )

    if logger:
        logger.info("Converting DOCX to Pages: %s", pages_path)

    result = run_applescript(applescript)
    if result.returncode != 0:
        raise RuntimeError("Pages conversion failed. Allow macOS Automation permission if prompted.\n" + (result.stderr or result.stdout))
    if not pages_path.exists():
        raise RuntimeError("Pages conversion completed but .pages file was not created.")
    return pages_path

def convert_docx_to_pdf_mac_pages(docx_path, logger=None):
    if not is_mac():
        raise RuntimeError("Apple Pages PDF export is supported only on macOS.")
    if not app_exists_mac("Pages"):
        raise RuntimeError("Apple Pages is not installed.")

    docx_path = Path(docx_path).resolve()
    pdf_path = docx_path.with_suffix(".pdf").resolve()

    applescript = (
        f'set inputFile to POSIX file "{docx_path}" as alias\n'
        f'set outputFile to POSIX file "{pdf_path}"\n'
        'tell application "Pages"\n'
        '    open inputFile\n'
        '    delay 1\n'
        '    set theDoc to front document\n'
        '    export theDoc to outputFile as PDF\n'
        '    close theDoc saving no\n'
        'end tell\n'
    )

    if logger:
        logger.info("Exporting PDF using Apple Pages: %s", pdf_path)

    result = run_applescript(applescript)
    if result.returncode != 0:
        raise RuntimeError("Pages PDF export failed. Allow macOS Automation permission if prompted.\n" + (result.stderr or result.stdout))
    if not pdf_path.exists():
        raise RuntimeError("Pages PDF export completed but PDF file was not created.")
    return pdf_path

def convert_docx_to_pdf_windows_word(docx_path, logger=None):
    if not is_windows():
        raise RuntimeError("Microsoft Word PDF export is supported only on Windows.")
    try:
        import win32com.client
    except Exception:
        raise RuntimeError("pywin32 is required for Word PDF export. Install with: python -m pip install pywin32")

    docx_path = Path(docx_path).resolve()
    pdf_path = docx_path.with_suffix(".pdf").resolve()

    if logger:
        logger.info("Exporting PDF using Microsoft Word: %s", pdf_path)

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path))
        doc.ExportAsFixedFormat(str(pdf_path), 17)
        doc.Close(False)
    finally:
        word.Quit()

    if not pdf_path.exists():
        raise RuntimeError("Word PDF export completed but PDF file was not created.")
    return pdf_path

def convert_docx_to_pdf_libreoffice(docx_path, logger=None):
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("LibreOffice is not installed.")
    docx_path = Path(docx_path).resolve()
    out_dir = docx_path.parent

    if logger:
        logger.info("Exporting PDF using LibreOffice: %s", docx_path)

    cmd = [str(soffice), "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(result.stderr or result.stdout or "LibreOffice PDF conversion failed.")
    pdf_path = docx_path.with_suffix(".pdf")
    if not pdf_path.exists():
        raise RuntimeError("LibreOffice conversion completed but PDF file was not created.")
    return pdf_path

def convert_docx_to_pdf(docx_path, logger=None):
    errors = []

    # Prefer headless LibreOffice first when available.
    # This avoids opening Pages/Word UI during normal generation.
    try:
        return convert_docx_to_pdf_libreoffice(docx_path, logger=logger)
    except Exception as e:
        errors.append("LibreOffice: " + str(e))

    # macOS fallback: Pages is used only when LibreOffice is not available.
    # Pages may briefly open the document for export, but it is closed automatically.
    if is_mac():
        try:
            return convert_docx_to_pdf_mac_pages(docx_path, logger=logger)
        except Exception as e:
            errors.append("Apple Pages: " + str(e))

    # Windows fallback: Word is used invisibly via COM automation.
    if is_windows():
        try:
            return convert_docx_to_pdf_windows_word(docx_path, logger=logger)
        except Exception as e:
            errors.append("Microsoft Word: " + str(e))

    raise RuntimeError("Unable to generate PDF.\n\n" + "\n\n".join(errors))

def open_folder(folder):
    folder = Path(folder)
    try:
        if is_mac():
            subprocess.run(["open", str(folder)])
        elif is_windows():
            os.startfile(str(folder))
        else:
            subprocess.run(["xdg-open", str(folder)])
    except Exception:
        pass
