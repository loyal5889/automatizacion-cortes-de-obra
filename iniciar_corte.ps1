param(
    [switch]$Manual,
    [switch]$Diagnostico,
    [switch]$Pruebas,
    [switch]$SinPausa
)

# Busca instalaciones reales; evita el alias de Microsoft Store.
# No modifica el PATH, el registro ni la politica permanente de PowerShell.
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
[Console]::InputEncoding = $utf8
$OutputEncoding = $utf8
$env:PYTHONIOENCODING = 'utf-8'
$candidates = New-Object 'System.Collections.Generic.List[string]'

function Add-PythonCandidate([string]$Path) {
    if ($Path -and (Test-Path -LiteralPath $Path -PathType Leaf) -and
        $Path -notmatch '\\Microsoft\\WindowsApps\\' -and -not $candidates.Contains($Path)) {
        $candidates.Add($Path)
    }
}

function Find-Python {
    Add-PythonCandidate (Join-Path $PSScriptRoot '.venv\Scripts\python.exe')
    Add-PythonCandidate $env:DARCA_PYTHON
    foreach ($name in @('python', 'python3')) {
        foreach ($command in @(Get-Command $name -All -CommandType Application -ErrorAction SilentlyContinue)) {
            Add-PythonCandidate $command.Source
        }
    }
    foreach ($key in @('HKCU:\Software\Python\PythonCore\*\InstallPath',
                       'HKLM:\Software\Python\PythonCore\*\InstallPath',
                       'HKLM:\Software\WOW6432Node\Python\PythonCore\*\InstallPath')) {
        foreach ($entry in @(Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)) {
            Add-PythonCandidate $entry.ExecutablePath
            if ($entry.'(default)') { Add-PythonCandidate (Join-Path $entry.'(default)' 'python.exe') }
        }
    }
    foreach ($command in @(Get-Command py -All -CommandType Application -ErrorAction SilentlyContinue)) {
        try {
            $located = @(& $command.Source -3 -c 'import sys; print(sys.executable)' 2>$null)
            if ($LASTEXITCODE -eq 0 -and $located.Count) { Add-PythonCandidate ([string]$located[-1]) }
        } catch { }
    }
    if ($env:LOCALAPPDATA) {
        $localPrograms = Join-Path $env:LOCALAPPDATA 'Programs\Python'
        foreach ($folder in @(Get-ChildItem -LiteralPath $localPrograms -Directory -ErrorAction SilentlyContinue)) {
            Add-PythonCandidate (Join-Path $folder.FullName 'python.exe')
        }
    }
    # Ultimo recurso: runtime existente en este equipo. No se instala ni modifica.
    if ($env:USERPROFILE) {
        Add-PythonCandidate (Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe')
    }
    foreach ($candidate in $candidates) {
        try {
            $check = @(& $candidate -c "import sys, openpyxl, PIL; assert sys.version_info >= (3,11); assert (3,1,5) <= tuple(int(x) for x in openpyxl.__version__.split('.')[:3]) < (4,); print('DARCA_OK')" 2>$null)
            if ($LASTEXITCODE -eq 0 -and $check -contains 'DARCA_OK') { return $candidate }
        } catch { }
    }
    throw 'No se encontro Python 3.11+ con openpyxl 3.1.5 o posterior (menor que 4) y Pillow. Consulte la seccion Inicio en Windows del README. Puede indicar un ejecutable con DARCA_PYTHON.'
}

$exitCode = 1
try {
    $pythonExe = Find-Python
    Write-Host "Python seleccionado: $pythonExe"
    Push-Location -LiteralPath $PSScriptRoot
    try {
        if ($Diagnostico) {
            & $pythonExe -c "import sys, openpyxl, PIL; print('Python:', sys.version.split()[0]); print('openpyxl:', openpyxl.__version__); print('Listo para ejecutar el generador de cortes.')"
        } elseif ($Pruebas) {
            & $pythonExe -m unittest discover -s tests -v
        } elseif ($Manual) {
            & $pythonExe (Join-Path $PSScriptRoot 'generar_corte.py') --manual
        } else {
            & $pythonExe (Join-Path $PSScriptRoot 'generar_corte.py')
        }
        $exitCode = $LASTEXITCODE
    } finally {
        Pop-Location
    }
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
} finally {
    if (-not $SinPausa) { [void](Read-Host 'Presione Enter para cerrar') }
}
exit $exitCode
