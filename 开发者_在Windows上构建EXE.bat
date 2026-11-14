@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo [1/4] 检查 Python（仅开发者构建时需要）...
where python >nul 2>nul || (echo 请在开发电脑安装 Python 3.11/3.12 x64 后重试。& pause & exit /b 1)
if not exist "tesseract\tesseract.exe" (echo 缺少 tesseract\tesseract.exe 和 chi_sim.traineddata，请先放入便携版 Tesseract OCR。& pause & exit /b 1)
echo [2/4] 安装构建依赖...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo [3/4] 构建单文件 EXE...
python -m PyInstaller --noconfirm --clean "集装箱资料智能生成系统.spec"
if errorlevel 1 goto failed
echo [4/4] 完成：dist\集装箱资料智能生成系统.exe
pause
exit /b 0
:failed
echo 构建失败，请检查错误信息。
pause
exit /b 1
