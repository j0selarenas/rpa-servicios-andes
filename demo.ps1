# Demo: ejecución, repetición el mismo día y reproceso desde landing.
# Uso (desde la carpeta del proyecto, con .venv ya creado):  .\demo.ps1

Set-Location $PSScriptRoot
$py = ".\.venv\Scripts\python.exe"
$hoy = Get-Date -Format "yyyy-MM-dd"

function Paso($titulo, [string[]]$argumentos) {
    Write-Host ""
    Write-Host "== $titulo ==" -ForegroundColor Cyan
    & $py -m rpa_andes @argumentos
    if ($LASTEXITCODE -ne 0) { Write-Host "Terminó con código $LASTEXITCODE" -ForegroundColor Red; exit $LASTEXITCODE }
}

Paso "1. Ejecución: captura UF y dólar del SII y procesa el Excel" @("ejecutar", "--entrada", "solicitudes.xlsx")
Paso "2. Misma ejecución otra vez: reutiliza la captura y no duplica" @("ejecutar", "--entrada", "solicitudes.xlsx")
Paso "3. Reproceso de $hoy desde landing, sin conexión al SII" @("reprocesar", "--fecha", $hoy)
Paso "Snapshots guardados" @("snapshots")
