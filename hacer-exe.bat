@echo off
REM Genera dist\Recibos.exe (un solo archivo, no necesita Python en la otra PC)
cd /d "%~dp0"
python -m pip install --upgrade pyinstaller
python -m PyInstaller --noconsole --onefile --clean --name Recibos recibos.py
echo.
echo Listo. El ejecutable quedo en dist\Recibos.exe
pause
