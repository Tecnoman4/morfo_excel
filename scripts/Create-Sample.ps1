$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()
$ws = $wb.Sheets.Item(1)
$ws.Name = "Datos de Prueba"

# Encabezados
$ws.Range("A1").Value2 = "ID"
$ws.Range("B1").Value2 = "Texto Original (Desordenado / Sucio)"
$ws.Range("C1").Value2 = "Resultado Esperado / Notas"
$ws.Range("A1:C1").Font.Bold = $true
$ws.Range("A1:C1").Interior.Color = 14277081 # Gris suave

# Fila 2: Caso de Mayusculas / Espacios
$ws.Range("A2").Value2 = 1
$ws.Range("B2").Value2 = "   juan   carlos   g" + [char]0xF3 + "mez   valle   "
$ws.Range("C2").Value2 = "Probar: Quitar Espacios y Nombre Propio"

# Fila 3: Caso de Tildes y Minusculas
$ws.Range("A3").Value2 = 2
$ws.Range("B3").Value2 = ([char]0xC1 + "NGEL P" + [char]0xC9 + "REZ - DIRECCI" + [char]0xD3 + "N DE LOG" + [char]0xCD + "STICA (M" + [char]0xC9 + "XICO)")
$ws.Range("C3").Value2 = "Probar: Quitar Acentos / minusculas"

# Fila 4: Caso de Caracteres Raros / Salto de linea
$ws.Range("A4").Value2 = 3
$ws.Range("B4").Value2 = ("ERP_REF" + [char]10 + "00984" + [char]0x200B + "#!?" + [char]13)
$ws.Range("C4").Value2 = "Probar: Quitar Raros de Bases de Datos"

# Fila 5 a 8: Rango continuo de filas para procesar en lote
$ws.Range("A5").Value2 = 4
$ws.Range("B5").Value2 = ("  mar" + [char]0xED + "a de los " + [char]0xE1 + "ngeles pe" + [char]0xF1 + "a  ")
$ws.Range("C5").Value2 = "Probar fila individual o en bloque con las de abajo"

$ws.Range("A6").Value2 = 5
$ws.Range("B6").Value2 = ("  jos" + [char]0xE9 + " luis mu" + [char]0xF1 + "iz  ")
$ws.Range("C6").Value2 = "Probar seleccion conjunta B5:B8 con Limpieza Total"

$ws.Range("A7").Value2 = 6
$ws.Range("B7").Value2 = ("  ra" + [char]0xFA + "l gonz" + [char]0xE1 + "lez s" + [char]0xE1 + "nchez  ")
$ws.Range("C7").Value2 = "Nombre Propio + Quitar Acentos"

$ws.Range("A8").Value2 = 7
$ws.Range("B8").Value2 = ("   bogot" + [char]0xE1 + ",   medell" + [char]0xED + "n,   asunci" + [char]0xF3 + "n,   canc" + [char]0xFA + "n   ")
$ws.Range("C8").Value2 = "Espacios + Acentos + MAYUSCULAS"

# Fila 9: Caso de Tipo Oracion
$ws.Range("A9").Value2 = 8
$ws.Range("B9").Value2 = "ESTE ES UN TEXTO EN MAY" + [char]0xDA + "SCULAS PARA CONVERTIR A TIPO ORACI" + [char]0xD3 + "N."
$ws.Range("C9").Value2 = "Probar: Tipo Oracion (Solo 1a letra mayuscula)"

# Ajustar ancho de columnas
$ws.Columns.Item(1).AutoFit() | Out-Null
$ws.Columns.Item(2).ColumnWidth = 52
$ws.Columns.Item(3).ColumnWidth = 48

$outSample = Join-Path (Split-Path -Parent $PSScriptRoot) "Ejemplo_Datos_Prueba.xlsx"
if (Test-Path $outSample) { Remove-Item $outSample -Force }
$wb.SaveAs($outSample)
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host ("Archivo de prueba generado: " + $outSample) -ForegroundColor Green
