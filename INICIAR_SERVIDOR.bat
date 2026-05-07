@echo off
title CHUMACERO ESPADA - Control de Inventarios
color 1F
cls

echo.
echo  ============================================================
echo   CHUMACERO ESPADA S.R.L. - CONTROL DE INVENTARIOS v2.0
echo  ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado. Instala Python 3.10 o superior.
    echo  Descarga: https://www.python.org/downloads/
    pause & exit
)

echo  Verificando dependencias...
pip install flask openpyxl --quiet --disable-pip-version-check

for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set IP=%%a & goto :found
)
:found
set IP=%IP: =%

echo.
echo  ============================================================
echo   Servidor iniciado correctamente
echo  ============================================================
echo.
echo   Local:       http://localhost:5000
echo   En red WiFi: http://%IP%:5000
echo.
echo   Comparte la IP con otras laptops en el mismo WiFi.
echo   Presiona Ctrl+C para detener el servidor.
echo  ============================================================
echo.

timeout /t 2 /nobreak >nul
start http://localhost:5000
python app.py
pause
