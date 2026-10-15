@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Ejecucion manual. Para quitar el filtro de importe minimo, borra --min 500000
rem Python: lanzador "py" o, en su defecto, "python" del PATH
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY (where python >nul 2>nul && set "PY=python")
if not defined PY (
  echo No se encuentra Python. Instalalo desde python.org marcando "Add Python to PATH".
  pause
  exit /b 1
)
rem openpyxl es necesario para el Excel; si falta, se instala una vez
%PY% -c "import openpyxl" >nul 2>nul || %PY% -m pip install --user --disable-pip-version-check openpyxl
%PY% licitaciones_obras.py --min 500000 --log registro.log --abrir siempre %*
echo.
pause
