@echo off
setlocal
cd /d "%~dp0"
echo Dang mo web dinh dang BOM...

if exist "BomFormatterWeb.exe" (
    BomFormatterWeb.exe --host 127.0.0.1 --port 8100
    goto end
)

rem Khi chay trong thu muc ma nguon, uu tien moi truong Python cuc bo.
if exist ".venv\Scripts\python.exe" set "BOM_PYTHON=.venv\Scripts\python.exe"
if not defined BOM_PYTHON if exist ".venv\bin\python.exe" set "BOM_PYTHON=.venv\bin\python.exe"
if not defined BOM_PYTHON if exist ".build-venv\Scripts\python.exe" set "BOM_PYTHON=.build-venv\Scripts\python.exe"
if not defined BOM_PYTHON if exist ".build-venv\bin\python.exe" set "BOM_PYTHON=.build-venv\bin\python.exe"
if not defined BOM_PYTHON set "BOM_PYTHON=python"

"%BOM_PYTHON%" -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Dang cai thu vien openpyxl, vui long doi...
    "%BOM_PYTHON%" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Khong cai duoc thu vien. Hay kiem tra ket noi Internet va Python.
        goto end
    )
)

"%BOM_PYTHON%" main.py --host 127.0.0.1 --port 8100

:end
pause
