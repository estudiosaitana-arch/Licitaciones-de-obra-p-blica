@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Programa la ejecucion automatica de lunes a viernes. Hora por defecto 07:45.
rem Para otra hora: programar_tarea.bat 08:30
set "CARPETA=%~dp0"
set "HORA_TAREA=%~1"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$c = Get-Content -Raw -Encoding UTF8 -LiteralPath (Join-Path $env:CARPETA 'programar_tarea.ps1'); Invoke-Expression $c"
echo.
pause
