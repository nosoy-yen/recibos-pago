@echo off
cd /d "%~dp0"
where pythonw >nul 2>&1 && (start "" pythonw recibos.py) || (python recibos.py)
