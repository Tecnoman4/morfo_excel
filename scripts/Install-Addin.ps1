# ===============================================================================
# SCRIPT: Install-Addin.ps1
# PROPÓSITO: Instala, desbloquea y registra de forma permanente el complemento
#            en Microsoft Excel para que funcione en cualquier libro (.xlsx, .xlsm, .csv)
# ===============================================================================

[CmdletBinding()]
param(
    [string]$AddinName = "Morfo.xlam",
    [string]$TabPosition = "Final"
)

$ErrorActionPreference = 'Stop'

Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host ' INSTALADOR AUTOMATICO DE MORFO PARA MICROSOFT EXCEL      ' -ForegroundColor Cyan
Write-Host '==========================================================' -ForegroundColor Cyan

$projectRoot = Split-Path -Parent $PSScriptRoot
$sourceAddin = Join-Path $projectRoot ('dist\' + $AddinName)

# 1. Compilar complemento con la posición de pestaña deseada
Write-Host "Compilando complemento con posición de pestaña: $TabPosition..." -ForegroundColor Yellow
& "$PSScriptRoot\Build-Addin.ps1" -OutputFileName $AddinName -TabPosition $TabPosition

# 2. Cerrar Excel si está abierto para evitar bloqueos de archivos en uso
$excelProc = Get-Process excel -ErrorAction SilentlyContinue
if ($null -ne $excelProc) {
    Write-Host 'Cerrando instancias de Excel abiertas para actualizar archivos...' -ForegroundColor Yellow
    Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

# 3. Definir rutas de instalación
$appData = [Environment]::GetFolderPath('ApplicationData')
$addInsDir = Join-Path $appData 'Microsoft\AddIns'
$xlStartDir = Join-Path $appData 'Microsoft\Excel\XLSTART'

if (-not (Test-Path $addInsDir)) { New-Item -ItemType Directory -Path $addInsDir -Force | Out-Null }

$destAddIn = Join-Path $addInsDir $AddinName
$destXLStart = Join-Path $xlStartDir $AddinName

# Limpiar cualquier copia en XLSTART para EVITAR el error "Excel no puede abrir dos libros con el mismo nombre"
if (Test-Path $destXLStart) {
    Remove-Item $destXLStart -Force -ErrorAction SilentlyContinue
    Write-Host '   -> Eliminada copia redundante en XLSTART para evitar conflicto de nombres.' -ForegroundColor Green
}

# Limpiar versión anterior si existía
$oldAddIn = Join-Path $addInsDir "HerramientasDeTexto.xlam"
$oldXLStart = Join-Path $xlStartDir "HerramientasDeTexto.xlam"
if (Test-Path $oldAddIn) { Remove-Item $oldAddIn -Force -ErrorAction SilentlyContinue }
if (Test-Path $oldXLStart) { Remove-Item $oldXLStart -Force -ErrorAction SilentlyContinue }

Write-Host '1. Copiando complemento Morfo a la carpeta oficial de AddIns...' -ForegroundColor Yellow
Copy-Item -Path $sourceAddin -Destination $destAddIn -Force

# Desbloquear archivos descargados o generados (eliminar marca Zone.Identifier de Windows)
Unblock-File -Path $destAddIn -ErrorAction SilentlyContinue
Write-Host '   -> Archivo copiado y desbloqueado correctamente en Microsoft\AddIns.' -ForegroundColor Green

# 4. Configurar Ubicaciones de Confianza (Trusted Locations) en el Registro
Write-Host '2. Configurando Ubicaciones de Confianza (Trusted Locations)...' -ForegroundColor Yellow
$trustedRoot = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations'

$locations = @(
    @{ Key = 'UserAddins'; Path = $addInsDir; Desc = 'Complementos de usuario' },
    @{ Key = 'UserXLSTART'; Path = $xlStartDir; Desc = 'Inicio automatico de Excel' },
    @{ Key = 'DevAnsiExcel'; Path = $projectRoot; Desc = 'Carpeta de desarrollo del complemento' }
)

foreach ($loc in $locations) {
    $locPath = Join-Path $trustedRoot $loc.Key
    if (-not (Test-Path $locPath)) {
        New-Item -Path $locPath -Force | Out-Null
    }
    Set-ItemProperty -Path $locPath -Name 'Path' -Value $loc.Path -Force
    Set-ItemProperty -Path $locPath -Name 'AllowSubfolders' -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $locPath -Name 'Description' -Value $loc.Desc -Force
    Write-Host ('   -> Confianza otorgada a: ' + $loc.Path) -ForegroundColor Green
}

# 5. Registrar el inicio automático en HKCU:\Software\Microsoft\Office\16.0\Excel\Options
Write-Host '3. Configurando carga automatica en opciones de Excel...' -ForegroundColor Yellow
$optionsKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
if (Test-Path $optionsKey) {
    $props = Get-ItemProperty -Path $optionsKey
    $openKeys = @('OPEN') + (1..20 | ForEach-Object { 'OPEN' + $_ })
    
    # Limpiar entradas de versiones anteriores o desinstaladas
    foreach ($k in $openKeys) {
        if ($null -ne $props.$k -and $props.$k -like '*HerramientasDeTexto*') {
            Remove-ItemProperty -Path $optionsKey -Name $k -Force -ErrorAction SilentlyContinue
        }
    }
    
    $props = Get-ItemProperty -Path $optionsKey
    $alreadyRegistered = $false
    $firstFreeKey = $null

    foreach ($k in $openKeys) {
        if ($null -ne $props.$k) {
            if ($props.$k -like ('*' + $AddinName + '*')) {
                $alreadyRegistered = $true
                Write-Host ('   -> Ya registrado previamente en: ' + $k) -ForegroundColor Green
                break
            }
        } elseif ($null -eq $firstFreeKey) {
            $firstFreeKey = $k
        }
    }

    if (-not $alreadyRegistered -and ($null -ne $firstFreeKey)) {
        $regValue = '/R "' + $destAddIn + '"'
        Set-ItemProperty -Path $optionsKey -Name $firstFreeKey -Value $regValue -Force
        Write-Host ('   -> Registrado exitosamente en Excel Options como: ' + $firstFreeKey) -ForegroundColor Green
    }
}

# 6. Registrar en el catalogo de AddIns de Excel vía COM
Write-Host '4. Activando complemento en el catalogo de Excel...' -ForegroundColor Yellow
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    $dummyWb = $excel.Workbooks.Add()
    $addinObj = $excel.AddIns.Add($destAddIn, $false)
    $addinObj.Installed = $true
    $dummyWb.Close($false)
    $excel.Quit()
    Write-Host '   -> Activado en la lista de Complementos de Excel.' -ForegroundColor Green
}
catch {
    Write-Host ('   Nota COM: ' + $_.Exception.Message) -ForegroundColor Gray
    if ($null -ne $excel) { $excel.Quit() }
}
finally {
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ''
Write-Host '==========================================================' -ForegroundColor Green
Write-Host ' ¡INSTALACION DE MORFO COMPLETADA CON EXITO!              ' -ForegroundColor Green
Write-Host ' El complemento funcionara en TODOS los libros (.xlsx, etc)' -ForegroundColor Green
Write-Host ' 1. Abre Microsoft Excel normalmente                     ' -ForegroundColor Green
Write-Host ' 2. Ve a la pestaña "Morfo"                               ' -ForegroundColor Green
Write-Host ' 3. O haz clic derecho sobre cualquier celda o rango     ' -ForegroundColor Green
Write-Host '==========================================================' -ForegroundColor Green
