@echo off
REM Genera dist\Recibos.exe (un solo archivo, no necesita Python en la otra PC).
REM Si al lado hay un client_secret*.json de Google, lo mete dentro del .exe
REM para que funcione "Iniciar sesion con Gmail".
cd /d "%~dp0"
python -m pip install --upgrade pyinstaller
set EXTRA=
for %%f in (client_secret*.json) do set EXTRA=--add-data "%%f;."
if not defined EXTRA echo AVISO: no hay client_secret.json, el .exe no tendra "Iniciar sesion con Gmail".
python -m PyInstaller --noconsole --onefile --clean --name Recibos %EXTRA% recibos.py
echo.
echo Listo. El ejecutable quedo en dist\Recibos.exe
pause
