# ===============================================================================
# SCRIPT: Configure-Icons.ps1
# PROPÓSITO: Asistente interactivo para seleccionar y asignar logos e iconos
#            a cada una de las funciones del complemento Morfo en Excel.
# ===============================================================================

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$customUIDir = Join-Path $projectRoot "customUI"
$logosDir = Join-Path $customUIDir "logos"
$iconConfigPath = Join-Path $customUIDir "icon_config.json"

if (-not (Test-Path $logosDir)) {
    New-Item -ItemType Directory -Path $logosDir -Force | Out-Null
}

function Get-IconConfig {
    if (Test-Path $iconConfigPath) {
        return (Get-Content -Path $iconConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    return $null
}

function Set-IconConfig ($configObj) {
    $json = $configObj | ConvertTo-Json -Depth 5
    Set-Content -Path $iconConfigPath -Value $json -Encoding UTF8
}

function Select-LogoFile {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Selecciona el logo / imagen para Morfo"
    $dialog.Filter = "Imágenes PNG (*.png)|*.png|Todos los formatos (*.png;*.jpg;*.ico;*.bmp)|*.png;*.jpg;*.ico;*.bmp"
    $dialog.InitialDirectory = $logosDir

    $res = $dialog.ShowDialog()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
        $sourceFile = $dialog.FileName
        $destFile = Join-Path $logosDir ([System.IO.Path]::GetFileName($sourceFile))
        if ($sourceFile -ne $destFile) {
            Copy-Item -Path $sourceFile -Destination $destFile -Force
            Write-Host "   -> Archivo copiado a customUI\logos\" -ForegroundColor Green
        }
        return [System.IO.Path]::GetFileName($sourceFile)
    }
    return $null
}

# Bucle principal de interacción
while ($true) {
    Clear-Host
    Write-Host '==============================================================' -ForegroundColor Cyan
    Write-Host '         CONFIGURADOR DE LOGOS E ICONOS - MORFO EXCEL         ' -ForegroundColor Cyan
    Write-Host '==============================================================' -ForegroundColor Cyan
    Write-Host ''

    $cfg = Get-IconConfig
    if ($null -eq $cfg) {
        Write-Host "No se encontró el archivo icon_config.json en $customUIDir" -ForegroundColor Red
        pause
        break
    }

    # Mostrar estado actual de las funciones
    $keys = @($cfg.PSObject.Properties | ForEach-Object { $_.Name })
    $keyIndexMap = @{}
    $i = 1

    Write-Host 'Funciones y Logos asignados:' -ForegroundColor Yellow
    foreach ($k in $keys) {
        $item = $cfg.$k
        $img = $item.image
        $status = ""
        if ([string]::IsNullOrWhiteSpace($img)) {
            $status = "(Nativo de Office: " + $item.defaultImageMso + ")"
            $color = "DarkGray"
        } else {
            $imgPath = Join-Path $logosDir $img
            if (Test-Path $imgPath) {
                $status = "[$img]"
                $color = "Green"
            } else {
                $status = "[$img - ¡ARCHIVO NO ENCONTRADO!]"
                $color = "Red"
            }
        }
        $keyIndexMap[$i.ToString()] = $k
        $line = "  [{0}] {1,-18} -> {2}" -f $i, $item.label, $status
        Write-Host $line -ForegroundColor $color
        $i++
    }

    Write-Host ''
    Write-Host 'Logos disponibles en customUI\logos:' -ForegroundColor Yellow
    $availableImages = Get-ChildItem -Path $logosDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.(png|jpg|jpeg|ico|bmp)$' }
    if ($availableImages.Count -eq 0) {
        Write-Host '  (No hay archivos de imagen en customUI\logos aún)' -ForegroundColor DarkGray
    } else {
        foreach ($imgFile in $availableImages) {
            Write-Host "  * $($imgFile.Name) ($($imgFile.Length) bytes)" -ForegroundColor Gray
        }
    }

    Write-Host ''
    Write-Host '--------------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host " [1-$($keys.Count)] Seleccionar función para cambiar o asignar su logo" -ForegroundColor White
    Write-Host ' [G]   Guardar y Recompilar/Instalar en Excel ahora mismo' -ForegroundColor Cyan
    Write-Host ' [R]   Restaurar todas las funciones a iconos nativos de Office' -ForegroundColor DarkYellow
    Write-Host ' [0]   Salir' -ForegroundColor White
    Write-Host '--------------------------------------------------------------' -ForegroundColor DarkGray
    $op = Read-Host "Elige una opción"

    if ($op -eq "0") {
        break
    } elseif ($op -eq "G" -or $op -eq "g") {
        Write-Host ''
        Write-Host 'Compilando e instalando complemento Morfo con los logos seleccionados...' -ForegroundColor Yellow
        & "$PSScriptRoot\Install-Addin.ps1"
        Write-Host ''
        Write-Host 'Presiona Enter para continuar...' -ForegroundColor Cyan
        Read-Host
    } elseif ($op -eq "R" -or $op -eq "r") {
        $confirm = Read-Host "¿Seguro que deseas volver a los iconos nativos para todas las funciones? (s/n)"
        if ($confirm -eq "s" -or $confirm -eq "S") {
            foreach ($k in $keys) {
                $cfg.$k.image = ""
            }
            Set-IconConfig $cfg
            Write-Host 'Todos los botones se restablecieron a los iconos nativos de Office.' -ForegroundColor Green
            Start-Sleep -Seconds 1
        }
    } elseif ($keyIndexMap.ContainsKey($op)) {
        $selectedKey = $keyIndexMap[$op]
        $selectedItem = $cfg.$selectedKey

        Write-Host ''
        Write-Host "=== Asignar logo para: $($selectedItem.label) ===" -ForegroundColor Cyan
        Write-Host "Actual: $(if ($selectedItem.image) { $selectedItem.image } else { 'Nativo (' + $selectedItem.defaultImageMso + ')' })"
        Write-Host ''

        $imgList = @($availableImages)
        $subIdx = 1
        if ($imgList.Count -gt 0) {
            Write-Host 'Elige un archivo existente en customUI\logos:' -ForegroundColor Yellow
            foreach ($im in $imgList) {
                Write-Host "  [$subIdx] $($im.Name)"
                $subIdx++
            }
        }

        Write-Host ''
        Write-Host '  [B] Buscar y elegir imagen con el Explorador de Windows...' -ForegroundColor Cyan
        Write-Host '  [D] Quitar logo personalizado y usar icono nativo de Office' -ForegroundColor DarkYellow
        Write-Host '  [C] Cancelar / Volver' -ForegroundColor DarkGray
        Write-Host ''

        $subOp = Read-Host "Selecciona una opción"
        if ($subOp -eq "C" -or $subOp -eq "c") {
            continue
        } elseif ($subOp -eq "D" -or $subOp -eq "d") {
            $selectedItem.image = ""
            Set-IconConfig $cfg
            Write-Host "Restablecido a icono nativo." -ForegroundColor Green
            Start-Sleep -Seconds 1
        } elseif ($subOp -eq "B" -or $subOp -eq "b") {
            Write-Host "Abriendo selector de archivos..." -ForegroundColor Yellow
            $chosen = Select-LogoFile
            if ($null -ne $chosen) {
                $selectedItem.image = $chosen
                Set-IconConfig $cfg
                Write-Host "Asignado: $chosen" -ForegroundColor Green
                Start-Sleep -Seconds 1
            }
        } elseif ([int]::TryParse($subOp, [ref]$null)) {
            $num = [int]$subOp
            if ($num -ge 1 -and $num -le $imgList.Count) {
                $chosen = $imgList[$num - 1].Name
                $selectedItem.image = $chosen
                Set-IconConfig $cfg
                Write-Host "Asignado: $chosen" -ForegroundColor Green
                Start-Sleep -Seconds 1
            }
        }
    }
}
