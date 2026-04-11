@echo off
setlocal
python -m pip install -r requirements.txt
pyinstaller --noconfirm --clean main.spec
endlocal
