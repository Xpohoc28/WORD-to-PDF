# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['test2.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'tkinter',
        'tkinter.filedialog',
        'tkinter.ttk',
        'tkinter.messagebox',
        'docx2pdf',
        'adobe.pdfservices.operation.auth.credentials',
        'adobe.pdfservices.operation.execution_context',
        'adobe.pdfservices.operation.io.file_ref',
        'adobe.pdfservices.operation.pdfops.convert_pdf_operation',
        'queue',
        'threading',
        'os',
        'comtypes',
        'comtypes.client',
        'win32com',
        'win32com.client',
        'win32com.client.gencache',
        'pythoncom',
        'docx2pdf.converter',
        'sys',
        'subprocess',
        'concurrent.futures',
        'logging'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DocToPdfConverter',
    debug=True,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='C:\\Users\\Xpohoc28\\Coding\\gang\\icon.ico'
)