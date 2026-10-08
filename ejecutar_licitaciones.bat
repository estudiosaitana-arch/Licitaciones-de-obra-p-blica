@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Actualizacion diaria. Para quitar el filtro de importe, borra --min 500000 --max 10000000
where python >/dev/null 2>nul
if %errorlevel%==0 (set PY=python) else (set PY=py)
%PY% licitaciones_obras.py --min 500000 --max 10000000 %*
if exist licitaciones_en_plazo.csv start "" "licitaciones_en_plazo.csv"
echo.
pause
