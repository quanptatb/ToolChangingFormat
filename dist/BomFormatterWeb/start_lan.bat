@echo off
cd /d "%~dp0"
echo Chay web cho cac may trong cung mang LAN.
echo Dia chi IP cua may nay:
ipconfig | findstr /i "IPv4"
echo.
echo Neu Windows Firewall hoi, chon Allow access.
if exist "BomFormatterWeb.exe" (
    BomFormatterWeb.exe --host 0.0.0.0 --port 8000
) else (
    python main.py --host 0.0.0.0 --port 8000
)
pause
