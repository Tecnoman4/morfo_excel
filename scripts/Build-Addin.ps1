# ===============================================================================
# SCRIPT: Build-Addin.ps1
# PROPOSITO: Compila el complemento de Excel (.xlam) inyectando los modulos VBA
#            y el archivo de configuracion RibbonX / Menu Contextual (CustomUI XML).
# ===============================================================================

[CmdletBinding()]
param(
    [string]$OutputFileName = "Morfo.xlam"
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $projectRoot "src"
$customUIDir = Join-Path $projectRoot "customUI"
$distDir = Join-Path $projectRoot "dist"
$outputAddinPath = Join-Path $distDir $OutputFileName

Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host ' INICIANDO COMPILACION DEL COMPLEMENTO DE EXCEL (.XLAM)   ' -ForegroundColor Cyan
Write-Host '==========================================================' -ForegroundColor Cyan

# 1. Asegurar directorio de distribucion
if (-not (Test-Path $distDir)) {
    New-Item -ItemType Directory -Path $distDir | Out-Null
}

if (Test-Path $outputAddinPath) {
    Remove-Item $outputAddinPath -Force
}

# 2. Configurar acceso seguro a VBOM en el registro para permitir la automatizacion COM
$secKeys = @(
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Security",
    "HKCU:\Software\Microsoft\Office\15.0\Excel\Security",
    "HKCU:\Software\Microsoft\Office\14.0\Excel\Security"
)

foreach ($key in $secKeys) {
    if (Test-Path (Split-Path $key)) {
        if (-not (Test-Path $key)) {
            New-Item -Path $key -Force | Out-Null
        }
        Set-ItemProperty -Path $key -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    }
}

# 3. Iniciar instancia oculta de Excel
Write-Host "Iniciando Microsoft Excel para compilar proyecto VBA..." -ForegroundColor Yellow
$excel = $null
$wb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 1 # msoAutomationSecurityLow

    # Crear libro nuevo
    $wb = $excel.Workbooks.Add()
    $wb.VBProject.Name = "Morfo"

    # Importar modulos .bas
    $basFiles = Get-ChildItem -Path $srcDir -Filter "*.bas"
    foreach ($bas in $basFiles) {
        Write-Host " -> Importando modulo: $($bas.Name)" -ForegroundColor Green
        $wb.VBProject.VBComponents.Import($bas.FullName) | Out-Null
    }

    # Importar formularios .frm de forma segura creando componentes MSForm nativos (Type 3)
    $frmFiles = Get-ChildItem -Path $srcDir -Filter "*.frm"
    foreach ($frm in $frmFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($frm.Name)
        Write-Host " -> Importando formulario MSForm: $baseName" -ForegroundColor Green
        
        $rawLines = Get-Content -Path $frm.FullName -Encoding UTF8
        $codeOnly = @()
        $startCode = $false
        $formCaption = "Morfo - Limpieza y Transformación Avanzada"
        
        foreach ($line in $rawLines) {
            if ($line -match '^\s*Caption\s*=\s*"(.*)"') {
                $formCaption = $matches[1]
            }
            if ($line -like "Option Explicit*" -or $line -like "Private Sub*" -or $line -like "Public Sub*") {
                $startCode = $true
            }
            if ($startCode) {
                $codeOnly += $line
            }
        }
        
        # 3 = vbext_ct_MSForm
        $formComp = $wb.VBProject.VBComponents.Add(3)
        $formComp.Name = $baseName
        try {
            $formComp.Properties.Item("Caption").Value = $formCaption
        } catch { }
        
        if ($codeOnly.Count -gt 0) {
            $codeBody = $codeOnly -join "`r`n"
            $formComp.CodeModule.AddFromString($codeBody)
        }
    }

    # Los eventos de apertura y cierre se gestionan automáticamente con Auto_Open y Auto_Close en modContextMenuFallback

    # Marcar como AddIn antes de guardar
    $wb.IsAddin = $true

    # Guardar como .xlam (xlOpenXMLAddIn = 55)
    Write-Host "Guardando estructura base como .xlam..." -ForegroundColor Yellow
    $xlOpenXMLAddIn = 55
    $wb.SaveAs($outputAddinPath, $xlOpenXMLAddIn)
    $wb.Close($false)
    $excel.Quit()
}
catch {
    Write-Host "Error durante la compilacion en Excel: $($_.Exception.Message)" -ForegroundColor Red
    if ($null -ne $wb) { $wb.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    throw
}
finally {
    if ($null -ne $wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# 4. Inyectar CustomUI14.xml en el paquete OpenXML (ZIP)
Write-Host "Inyectando cinta de opciones (Ribbon) y menu contextual XML..." -ForegroundColor Yellow
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$customUIXmlPath = Join-Path $customUIDir "customUI14.xml"
if (-not (Test-Path $customUIXmlPath)) {
    throw "No se encontro el archivo $customUIXmlPath"
}

# Abrir el ZIP del archivo .xlam
$zip = [System.IO.Compression.ZipFile]::Open($outputAddinPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    # 4.1 Anadir archivo customUI/customUI14.xml
    $customEntry = $zip.GetEntry("customUI/customUI14.xml")
    if ($null -ne $customEntry) {
        $customEntry.Delete()
    }
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $customUIXmlPath, "customUI/customUI14.xml") | Out-Null

    # 4.2 Actualizar _rels/.rels para registrar la relacion de CustomUI
    $relsEntry = $zip.GetEntry("_rels/.rels")
    if ($null -ne $relsEntry) {
        $stream = $relsEntry.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        $relsContent = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()

        $customUIRelId = "rIdCustomUI14"
        $relXmlSnippet = '<Relationship Id="' + $customUIRelId + '" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'

        if ($relsContent -notmatch "customUI/customUI14.xml") {
            $updatedRels = $relsContent -replace "</Relationships>", "$relXmlSnippet`r`n</Relationships>"
            $relsEntry.Delete()

            $newRelsEntry = $zip.CreateEntry("_rels/.rels")
            $writeStream = $newRelsEntry.Open()
            $writer = New-Object System.IO.StreamWriter($writeStream, [System.Text.Encoding]::UTF8)
            $writer.Write($updatedRels)
            $writer.Flush()
            $writer.Close()
            $writeStream.Close()
        }
    }
}
finally {
    $zip.Dispose()
}

Write-Host '==========================================================' -ForegroundColor Green
Write-Host ' COMPILACION EXITOSA!                                     ' -ForegroundColor Green
Write-Host (' Archivo generado: ' + $outputAddinPath) -ForegroundColor Green
Write-Host '==========================================================' -ForegroundColor Green
