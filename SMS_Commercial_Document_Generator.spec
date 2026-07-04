# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None
ROOT = Path.cwd()

added_data = [
    (str(ROOT / 'app'), 'app'),
    (str(ROOT / 'app_core'), 'app_core'),
    (str(ROOT / 'ui'), 'ui'),
    (str(ROOT / 'modules'), 'modules'),
    (str(ROOT / 'resources'), 'resources'),
    (str(ROOT / 'assets'), 'assets'),
    (str(ROOT / 'VERSION.txt'), '.'),
    *collect_data_files('docx'),
]

vendor_lo = ROOT / 'vendor' / 'libreoffice'
if vendor_lo.exists():
    added_data.append((str(vendor_lo), 'vendor/libreoffice'))

hiddenimports = [
    *collect_submodules('docx'),
    'customtkinter',
    'openpyxl',
    'reportlab',
    'PIL',
    'app_core.engines.modules.quotation.generate_proposal',
    'app_core.engines.modules.invoice.generate_invoice',
    'app_core.engines.modules.proforma.generate_proforma_invoice',
    'app_core.engines.modules.challan.generate_challan',
    'app_core.engines.modules.purchase_order.generate_purchase_order',
    'app_core.engines.modules.salary_slip.generate_salary_slip',
]

a = Analysis(
    ['commercial_document_generator_app.py'],
    pathex=[str(ROOT)],
    binaries=[],
    datas=added_data,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy.tests', 'pytest'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SMS Commercial Document Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ROOT / 'assets' / 'app_icon.ico') if (ROOT / 'assets' / 'app_icon.ico').exists() else None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SMS Commercial Document Generator',
)

if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='SMS Commercial Document Generator.app',
        icon=str(ROOT / 'assets' / 'app_icon.icns') if (ROOT / 'assets' / 'app_icon.icns').exists() else None,
        bundle_identifier='com.smscontrols.commercialdocumentgenerator',
        info_plist={
            'CFBundleName': 'SMS Commercial Document Generator',
            'CFBundleDisplayName': 'SMS Commercial Document Generator',
            'CFBundleShortVersionString': '1.0.0',
            'CFBundleVersion': '1.0.0',
            'NSHighResolutionCapable': True,
        },
    )
