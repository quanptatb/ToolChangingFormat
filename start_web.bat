@echo off
cd /d "%~dp0"
echo Dang mo web dinh dang BOM...
if exist "BomFormatterWeb.exe" (
    BomFormatterWeb.exe --host 127.0.0.1 --port 8100
) else (
    python main.py --host 127.0.0.1 --port 8100
)
pause
