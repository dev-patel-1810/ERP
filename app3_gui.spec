# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = []
datas += collect_data_files('tkcalendar')
datas += collect_data_files('babel')
datas += collect_data_files('flask')
datas += collect_data_files('jinja2')

hiddenimports = [
    'tkcalendar',
    'lxml',
    'babel',
    'babel.numbers',
    'babel.dates',
    'pandas',
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.cell._writer',
    'reportlab',
    'reportlab.graphics.barcode',
    'flask',
    'werkzeug',
    'werkzeug.serving',
    'werkzeug.debug',
    'jinja2',
    'jinja2.ext',
    'itsdangerous',
    'click',
    'threading',
    'sqlite3',
    'tkinter',
    'tkinter.ttk',
    'tkinter.messagebox',
    'tkinter.filedialog',
]

a = Analysis(
    ['app3_gui.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='RegistersSystem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,        # windowed app, no console box
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    onefile=True,
)