@echo off
cd /d "%~dp0"
if "%~1"=="" (
    echo Enter or paste the product title:
    set /p TITLE=""
) else (
    set TITLE=%*
)
python create_listing.py --tags "%TITLE%"
pause
