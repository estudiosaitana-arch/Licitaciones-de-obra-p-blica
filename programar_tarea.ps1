# Registra en el Programador de tareas de Windows la ejecución automática de licitaciones_obras.py.
# Se lanza desde programar_tarea.bat. No necesita permisos de administrador.
#
# Qué deja configurado:
#   - Tarea "Licitaciones PLACSP", de lunes a viernes a la hora indicada (por defecto 07:45).
#   - Si el ordenador estaba apagado a esa hora, se ejecuta en cuanto se inicia sesión.
#   - Se ejecuta con pythonw (sin ventana). La salida queda en registro.log.
#   - Si hay licitaciones nuevas o con cambios, abre licitaciones_en_plazo.xlsx al terminar.

$ErrorActionPreference = 'Stop'
$Carpeta = $env:CARPETA.TrimEnd('\')
$Nombre  = 'Licitaciones PLACSP'
$Hora    = if ($env:HORA_TAREA) { $env:HORA_TAREA } else { '07:45' }
$Dias    = 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'

# Argumentos de cada ejecución. Mismo rango de importe que ejecutar_licitaciones.bat.
# Para no abrir el Excel automáticamente, cambia "--abrir novedades" por "--abrir nunca".
$Argumentos = '--min 500000 --max 10000000 --log registro.log --abrir novedades'

function Falla($texto) {
    Write-Host ''
    Write-Host "ERROR: $texto" -ForegroundColor Red
    exit 1
}

Write-Host ''
Write-Host '=== Programar la actualización diaria de licitaciones ===' -ForegroundColor Cyan
Write-Host "Carpeta: $Carpeta"

if ($Hora -notmatch '^([01]?\d|2[0-3]):[0-5]\d$') { Falla "Hora no válida: '$Hora'. Usa el formato HH:MM, por ejemplo 08:30." }
$script = Join-Path $Carpeta 'licitaciones_obras.py'
if (-not (Test-Path -LiteralPath $script)) { Falla "No está licitaciones_obras.py en $Carpeta." }

# 1. Localizar Python (lanzador py o python del PATH). Se comprueba que funciona de verdad,
#    porque en Windows "python" puede ser un acceso a la Microsoft Store sin Python instalado.
$python = $null
foreach ($candidato in @('py', 'python')) {
    try {
        $salida = & $candidato -c "import sys; print(sys.executable if sys.version_info >= (3, 8) else '')" 2>$null
        $ruta = ($salida | Select-Object -Last 1)
        if ($LASTEXITCODE -eq 0 -and $ruta -and (Test-Path -LiteralPath $ruta.Trim())) { $python = $ruta.Trim(); break }
    } catch { }
}
if (-not $python) { Falla 'No se encuentra Python 3.8 o superior. Instálalo desde python.org (marcando "Add Python to PATH") y vuelve a ejecutar.' }
$pythonw = Join-Path (Split-Path -LiteralPath $python) 'pythonw.exe'
if (-not (Test-Path -LiteralPath $pythonw)) {
    Write-Host 'Aviso: no hay pythonw.exe; la ejecución programada mostrará una ventana mientras dura.' -ForegroundColor Yellow
    $pythonw = $python
}
Write-Host "Python:  $python"

# 2. openpyxl, necesario para el Excel (sin él, el script genera el CSV)
$ErrorActionPreference = 'Continue'
& $python -c "import openpyxl" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host 'Instalando openpyxl (necesario para el Excel)...'
    & $python -m pip install --user --disable-pip-version-check openpyxl
    if ($LASTEXITCODE -ne 0) {
        Write-Host 'Aviso: no se pudo instalar openpyxl. La tarea funcionará, pero generará CSV en lugar de Excel.' -ForegroundColor Yellow
        Write-Host '       Instálalo a mano con:  py -m pip install --user openpyxl' -ForegroundColor Yellow
    }
} else {
    Write-Host 'openpyxl: instalado'
}
$ErrorActionPreference = 'Stop'

# 3. Registrar la tarea (sustituye la anterior si ya existía)
try {
    $accion     = New-ScheduledTaskAction -Execute $pythonw -Argument ('"' + $script + '" ' + $Argumentos) -WorkingDirectory $Carpeta
    $disparador = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $Dias -At $Hora
    $ajustes    = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                    -ExecutionTimeLimit (New-TimeSpan -Hours 1) -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName $Nombre -Action $accion -Trigger $disparador -Settings $ajustes -Force `
        -Description 'Actualiza licitaciones_en_plazo.xlsx con las obras en plazo de la Plataforma de Contratación del Sector Público.' | Out-Null
} catch {
    Falla ("No se pudo registrar la tarea: " + $_.Exception.Message + "`nSi tu equipo está gestionado por la empresa, puede que la política no permita crear tareas programadas.")
}

Write-Host ''
Write-Host "Tarea '$Nombre' programada: de lunes a viernes a las $Hora." -ForegroundColor Green
Write-Host 'Si el ordenador está apagado a esa hora, se ejecuta al iniciar sesión.'
Write-Host 'Resultado:  licitaciones_en_plazo.xlsx   |   Registro de cada ejecución:  registro.log'
Write-Host 'Cambiar la hora:  programar_tarea.bat 08:30      Quitar la tarea:  quitar_tarea.bat'

if (-not (Test-Path -LiteralPath (Join-Path $Carpeta 'base_licitaciones.json'))) {
    Write-Host ''
    Write-Host 'Aún no hay base de licitaciones. Conviene ejecutar antes carga_inicial.bat (una sola vez).' -ForegroundColor Yellow
}

Write-Host ''
$respuesta = Read-Host '¿Lanzar la tarea ahora para comprobar que funciona? (S/N)'
if ($respuesta -match '^[sS]') {
    Start-ScheduledTask -TaskName $Nombre
    Write-Host 'Lanzada. Tarda entre 1 y 3 minutos; al terminar queda anotada en registro.log.'
}
