@echo off
setlocal
cd /d "%~dp0"

echo Tao moi truong build...
python -m venv .build-venv
if errorlevel 1 goto failed

echo Cai thu vien can thiet...
".build-venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto failed
".build-venv\Scripts\python.exe" -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto failed

echo Dong goi file exe...
".build-venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean BomFormatterWeb.spec
if errorlevel 1 goto failed

copy /Y start_web.bat dist\BomFormatterWeb\start_web.bat >nul
copy /Y start_lan.bat dist\BomFormatterWeb\start_lan.bat >nul

echo.
echo Da build xong.
echo Thu muc deploy: dist\BomFormatterWeb
echo Copy toan bo thu muc dist\BomFormatterWeb sang may khac, chay start_web.bat.
goto end

:failed
echo.
echo Build that bai. Kiem tra Python, Internet va quyen cai pip.
exit /b 1

:end
pause
