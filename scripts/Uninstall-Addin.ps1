# ===============================================================================
# SCRIPT: Uninstall-Addin.ps1
# PROPÓSITO: Desinstala limpiamente el complemento Morfo de Excel
# ===============================================================================

[CmdletBinding()]
param(
    [string]$AddinName = "Morfo.xlam"
)

$ErrorActionPreference = 'SilentlyContinue'

Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host ' DESINSTALADOR DE MORFO PARA EXCEL                        ' -ForegroundColor Cyan
Write-Host '==========================================================' -ForegroundColor Cyan

# 1. Cerrar Excel
$excelProc = Get-Process excel -ErrorAction SilentlyContinue
if ($null -ne $excelProc) {
    Write-Host 'Cerrando Excel...' -ForegroundColor Yellow
    Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

$appData = [Environment]::GetFolderPath('ApplicationData')
$namesToClean = @($AddinName, "HerramientasDeTexto.xlam")

foreach ($name in $namesToClean) {
    $destAddIn = Join-Path $appData ('Microsoft\AddIns\' + $name)
    $destXLStart = Join-Path $appData ('Microsoft\Excel\XLSTART\' + $name)

    if (Test-Path $destAddIn) { Remove-Item $destAddIn -Force }
    if (Test-Path $destXLStart) { Remove-Item $destXLStart -Force }
}

# 2. Limpiar entradas OPEN en Options
$optionsKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
if (Test-Path $optionsKey) {
    $props = Get-ItemProperty -Path $optionsKey
    $openKeys = @('OPEN') + (1..20 | ForEach-Object { 'OPEN' + $_ })
    foreach ($k in $openKeys) {
        if ($null -ne $props.$k) {
            foreach ($name in $namesToClean) {
                if ($props.$k -like ('*' + $name + '*')) {
                    Remove-ItemProperty -Path $optionsKey -Name $k -Force -ErrorAction SilentlyContinue
                    Write-Host ('Eliminada referencia en registro: ' + $k) -ForegroundColor Yellow
                }
            }
        }
    }
}

Write-Host ''
Write-Host 'Complemento Morfo desinstalado correctamente.' -ForegroundColor Green
Write-Host '==========================================================' -ForegroundColor Green
