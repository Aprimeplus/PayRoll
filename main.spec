# -*- mode: python ; coding: utf-8 -*--- "C:/Users/Nitro V15/AppData/Local/Programs/Python/Python313/python.exe" -m PyInstaller main.spec

block_cipher = None

# รายชื่อไฟล์ที่ต้องการรวมเข้าไปในโปรแกรม (ต้นทาง, ปลายทาง)
added_files = [
    ('resources', 'resources'),               # โฟลเดอร์ resources (ฟอนต์)
    ('company_logo.png', '.'),                # รูปโลโก้ (ถ้ามี)
    ('approve_wh3_081156.pdf', '.'),          # แบบฟอร์ม 50 ทวิ
    ('pnd1k.pdf', '.'),                       # แบบฟอร์ม ภ.ง.ด. 1ก
]

# รายชื่อ Library ที่อาจต้องบังคับรวม (Hidden Imports)
hidden_imports = [
    'psycopg2',
    'tksheet',
    'reportlab',
    'fpdf',
    'pypdf',
    'bahttext',
    'pandas',
    'tkinter',
    'PIL.ImageTk', 
    'PIL.Image',
    'babel.numbers'
]

a = Analysis(
    ['main.py'],             # ไฟล์หลักของโปรแกรม
    pathex=[],
    binaries=[],
    datas=added_files,
    hiddenimports=hidden_imports,
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
    [],
    exclude_binaries=True,
    name='APLUS_HR_Payroll',   # ชื่อไฟล์ .exe ที่จะได้
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,            # แก้เป็น True ถ้าอยากเห็นจอดำ (เอาไว้ดู Error), False คือซ่อนจอดำ
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='company_logo.png'   # (Optional) ถ้ามีไฟล์ .ico ให้เปลี่ยนเป็นชื่อไฟล์ .ico
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='APLUS_HR_Payroll',   # ชื่อโฟลเดอร์ที่จะสร้างใน dist/
)