@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Carga inicial: descarga los ficheros mensuales de los 2 ultimos meses (solo hace falta una vez)
where python >/dev/null 2>nul
if %errorlevel%==0 (set PY=python) else (set PY=py)
%PY% licitaciones_obras.py --meses 2 --min 500000 --max 10000000
if exist licitaciones_en_plazo.csv start "" "licitaciones_en_plazo.csv"
echo.
pause
