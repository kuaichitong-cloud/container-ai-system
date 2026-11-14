# PyInstaller Windows one-file spec; run on Windows x64 with Python 3.11/3.12
from PyInstaller.utils.hooks import collect_all
from pathlib import Path
root = Path(SPECPATH)
datas=[]; binaries=[]; hidden=[]
for pkg in ["PIL", "docx", "pytesseract"]:
    d,b,h=collect_all(pkg); datas+=d; binaries+=b; hidden+=h
# Keep OCR executable and language data next to the frozen app payload.
tess = root / "tesseract"
if tess.exists(): datas.append((str(tess), "tesseract"))
a = Analysis([str(root / "集装箱资料智能生成系统.py")], pathex=[str(root)], binaries=binaries, datas=datas, hiddenimports=hidden, hookspath=[], hooksconfig={}, runtime_hooks=[], excludes=[], noarchive=False)
pyz=PYZ(a.pure)
exe=EXE(pyz,a.scripts,a.binaries,a.datas,[],name="集装箱资料智能生成系统",debug=False,bootloader_ignore_signals=False,strip=False,upx=True,console=False,disable_windowed_traceback=False)
