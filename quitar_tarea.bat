@echo off
chcp 65001 >nul
rem Elimina la ejecucion automatica programada con programar_tarea.bat
powershell -NoProfile -Command "try { Unregister-ScheduledTask -TaskName 'Licitaciones PLACSP' -Confirm:$false -ErrorAction Stop; Write-Host 'Tarea eliminada.' } catch { Write-Host 'No habia ninguna tarea programada con ese nombre.' }"
echo.
pause
