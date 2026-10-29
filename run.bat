@echo off
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" create_listing.py %*
) else (
    python create_listing.py %*
)
pause
