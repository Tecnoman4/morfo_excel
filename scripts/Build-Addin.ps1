# ===============================================================================
# SCRIPT: Build-Addin.ps1
# PROPOSITO: Compila el complemento de Excel (.xlam) inyectando los modulos VBA
#            y el archivo de configuracion RibbonX / Menu Contextual (CustomUI XML).
# ===============================================================================

[CmdletBinding()]
param(
    [string]$OutputFileName = "Morfo.xlam",
    [string]$TabPosition = "Final"
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
        }
        catch { }
        
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
    # 4.1 Añadir archivo customUI/customUI14.xml con la posición elegida
    $customXmlRaw = Get-Content -Path $customUIXmlPath -Raw -Encoding UTF8
    $insertAttr = ""
    switch ($TabPosition) {
        "Inicio" { $insertAttr = ' insertAfterMso="TabHome"' }
        "Home" { $insertAttr = ' insertAfterMso="TabHome"' }
        "Insertar" { $insertAttr = ' insertAfterMso="TabInsert"' }
        "Datos" { $insertAttr = ' insertAfterMso="TabData"' }
        "Data" { $insertAttr = ' insertAfterMso="TabData"' }
        "Revisar" { $insertAttr = ' insertAfterMso="TabReview"' }
        "Review" { $insertAttr = ' insertAfterMso="TabReview"' }
        "Vista" { $insertAttr = ' insertAfterMso="TabView"' }
        "View" { $insertAttr = ' insertAfterMso="TabView"' }
        "Programador" { $insertAttr = ' insertAfterMso="TabDeveloper"' }
        "Developer" { $insertAttr = ' insertAfterMso="TabDeveloper"' }
        default { $insertAttr = "" } # Final (al final de la cinta)
    }
    $tabTagReplacement = "<tab id=`"tabMorfo`" label=`"Morfo`"$insertAttr>"
    $customXmlProcessed = $customXmlRaw -replace '<tab id="tabMorfo" label="Morfo"[^>]*>', $tabTagReplacement

    # Procesar iconos personalizados y nativos según icon_config.json
    function Update-XmlButtonIcon {
        param([string]$xmlText, [string]$buttonId, [string]$imageName, [string]$defaultMso)
        $pattern = "(?s)<button\b(?=[^>]*?\bid=`"$buttonId`")[^>]*?>"
        return [regex]::Replace($xmlText, $pattern, {
            param($m)
            $tag = $m.Value
            if ($imageName) {
                if ($tag -match 'imageMso="[^"]*"') {
                    return ($tag -replace 'imageMso="[^"]*"', "image=`"$imageName`"")
                } elseif ($tag -match 'image="[^"]*"') {
                    return ($tag -replace 'image="[^"]*"', "image=`"$imageName`"")
                } else {
                    return ($tag -replace '<button\b', "<button image=`"$imageName`"")
                }
            } else {
                if ($tag -match 'image="[^"]*"') {
                    return ($tag -replace 'image="[^"]*"', "imageMso=`"$defaultMso`"")
                } elseif ($tag -match 'imageMso="[^"]*"') {
                    return ($tag -replace 'imageMso="[^"]*"', "imageMso=`"$defaultMso`"")
                } else {
                    return ($tag -replace '<button\b', "<button imageMso=`"$defaultMso`"")
                }
            }
        })
    }

    $iconConfigPath = Join-Path $customUIDir "icon_config.json"
    $logosDir = Join-Path $customUIDir "logos"

    if (Test-Path $iconConfigPath) {
        try {
            $configJson = Get-Content -Path $iconConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($prop in $configJson.PSObject.Properties) {
                $btnId = $prop.Name
                $item = $prop.Value
                $assignedImg = [string]$item.image
                $ctxId = [string]$item.ctxId
                $defaultMso = [string]$item.defaultImageMso

                # Si no tiene imagen asignada explícita, buscar automáticamente por alias en customUI/logos
                if ([string]::IsNullOrWhiteSpace($assignedImg) -and (Test-Path $logosDir)) {
                    $aliases = @($btnId) + @($item.aliases)
                    foreach ($alias in $aliases) {
                        $candidate = Join-Path $logosDir ($alias + ".png")
                        if (Test-Path $candidate) {
                            $assignedImg = ($alias + ".png")
                            break
                        }
                    }
                }

                $imgBaseName = $null
                if (-not [string]::IsNullOrWhiteSpace($assignedImg)) {
                    $imgPath = Join-Path $logosDir $assignedImg
                    if (Test-Path $imgPath) {
                        $imgBaseName = [System.IO.Path]::GetFileNameWithoutExtension($assignedImg)
                    } else {
                        Write-Host "   -> Aviso: Imagen '$assignedImg' para $btnId no existe en customUI\logos, usando icono nativo." -ForegroundColor DarkYellow
                    }
                }

                $customXmlProcessed = Update-XmlButtonIcon -xmlText $customXmlProcessed -buttonId $btnId -imageName $imgBaseName -defaultMso $defaultMso
                if (-not [string]::IsNullOrWhiteSpace($ctxId)) {
                    $customXmlProcessed = Update-XmlButtonIcon -xmlText $customXmlProcessed -buttonId $ctxId -imageName $imgBaseName -defaultMso $defaultMso
                }

                if ($imgBaseName) {
                    Write-Host "   -> Botón '$($item.label)': asignado logo '$assignedImg'" -ForegroundColor Green
                }
            }
        } catch {
            Write-Host "   -> Aviso leyendo icon_config.json: $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }

    $customEntry = $zip.GetEntry("customUI/customUI14.xml")
    if ($null -ne $customEntry) {
        $customEntry.Delete()
    }
    $newXmlEntry = $zip.CreateEntry("customUI/customUI14.xml")
    $entryStream = $newXmlEntry.Open()
    $xmlWriter = New-Object System.IO.StreamWriter($entryStream, [System.Text.Encoding]::UTF8)
    $xmlWriter.Write($customXmlProcessed)
    $xmlWriter.Flush()
    $xmlWriter.Close()
    $entryStream.Close()
    Write-Host "   -> Pestaña 'Morfo' configurada en posición: $TabPosition" -ForegroundColor Green

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
        }
    }

    # 4.3 Inyectar logos e imágenes personalizadas (customUI/images) y sus relaciones
    $logosDir = Join-Path $customUIDir "logos"
    if (Test-Path $logosDir) {
        $imageFiles = Get-ChildItem -Path $logosDir -File
        if ($imageFiles.Count -gt 0) {
            Write-Host "Inyectando logos personalizados ($($imageFiles.Count))..." -ForegroundColor Yellow
            $relItems = @()
            foreach ($img in $imageFiles) {
                $baseImgName = [System.IO.Path]::GetFileNameWithoutExtension($img.Name)
                $targetInZip = "customUI/images/" + $img.Name
                
                # Eliminar entrada si ya existía
                $existingImg = $zip.GetEntry($targetInZip)
                if ($null -ne $existingImg) { $existingImg.Delete() }
                
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $img.FullName, $targetInZip) | Out-Null
                Write-Host "   -> Imagen agregada: $targetInZip (Id: $baseImgName)" -ForegroundColor Green
                
                $relItems += '<Relationship Id="' + $baseImgName + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/' + $img.Name + '"/>'
            }

            # Crear customUI/_rels/customUI14.xml.rels
            $relsXmlContent = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`r`n" +
                              "<Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">`r`n" +
                              ($relItems -join "`r`n") + "`r`n" +
                              "</Relationships>"

            $customUIRelsEntry = $zip.GetEntry("customUI/_rels/customUI14.xml.rels")
            if ($null -ne $customUIRelsEntry) { $customUIRelsEntry.Delete() }

            $newCustomUIRels = $zip.CreateEntry("customUI/_rels/customUI14.xml.rels")
            $relsStream = $newCustomUIRels.Open()
            $relsWriter = New-Object System.IO.StreamWriter($relsStream, [System.Text.Encoding]::UTF8)
            $relsWriter.Write($relsXmlContent)
            $relsWriter.Flush()
            $relsWriter.Close()
            $relsStream.Close()
            Write-Host "   -> Relaciones customUI14.xml.rels creadas correctamente." -ForegroundColor Green

            # 4.4 Asegurar que [Content_Types].xml incluya los tipos MIME de imágenes (png, etc.)
            $ctEntry = $zip.GetEntry("[Content_Types].xml")
            if ($null -ne $ctEntry) {
                $ctStream = $ctEntry.Open()
                $ctReader = New-Object System.IO.StreamReader($ctStream)
                $ctContent = $ctReader.ReadToEnd()
                $ctReader.Close()
                $ctStream.Close()

                $typesToAdd = ""
                if ($ctContent -notmatch 'Extension="png"') {
                    $typesToAdd += '<Default Extension="png" ContentType="image/png"/>'
                }
                if ($ctContent -notmatch 'Extension="jpg"') {
                    $typesToAdd += '<Default Extension="jpg" ContentType="image/jpeg"/>'
                }

                if ($typesToAdd -ne "") {
                    $updatedCt = $ctContent -replace "</Types>", "$typesToAdd</Types>"
                    $ctEntry.Delete()

                    $newCtEntry = $zip.CreateEntry("[Content_Types].xml")
                    $writeCtStream = $newCtEntry.Open()
                    $ctWriter = New-Object System.IO.StreamWriter($writeCtStream, [System.Text.Encoding]::UTF8)
                    $ctWriter.Write($updatedCt)
                    $ctWriter.Flush()
                    $ctWriter.Close()
                    $writeCtStream.Close()
                    Write-Host "   -> Tipos MIME de imágenes registrados en [Content_Types].xml" -ForegroundColor Green
                }
            }
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
