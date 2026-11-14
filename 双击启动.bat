@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist "集装箱资料智能生成系统.exe" (start "" "%~dp0集装箱资料智能生成系统.exe" & exit /b 0)
echo 当前文件夹还没有 Windows EXE 成品。
echo 请先阅读“重要说明_如何得到双击即用的EXE.txt”完成一次云端构建。
pause
