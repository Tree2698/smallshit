@echo off
setlocal

echo [1/3] Installing build dependency...
pip install pyinstaller

echo [2/3] Cleaning old build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [3/3] Building app...
pyinstaller --noconfirm --clean main.spec

echo.
echo Build finished.
echo Output folder: dist\xiaolaoxiang
pause
