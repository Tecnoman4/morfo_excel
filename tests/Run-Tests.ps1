# ===============================================================================
# SCRIPT: Run-Tests.ps1
# PROPOSITO: Verifica que el complemento instalado funcione en libros nuevos
# ===============================================================================

$ErrorActionPreference = 'Stop'

Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host ' VALIDANDO COMPLEMENTO INSTALADO EN MICROSOFT EXCEL       ' -ForegroundColor Cyan
Write-Host '==========================================================' -ForegroundColor Cyan

$excel = $null
$testWb = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1 # msoAutomationSecurityLow

    $addin = $excel.AddIns | Where-Object { $_.Name -like "*Morfo*" }
    if ($null -eq $addin) {
        throw "No se encontro Morfo en la lista de AddIns de Excel."
    }
    $addinWb = $excel.Workbooks.Open($addin.FullName)
    $runPrefix = "'" + $addin.Name + "'!"
    Write-Host ('Complemento cargado en memoria COM: ' + $addin.Name) -ForegroundColor Green

    # Crear libro de prueba
    $testWb = $excel.Workbooks.Add()
    $ws = $testWb.Sheets.Item(1)

    # 1. Probar que las funciones logicas y de transformacion se ejecutan
    Write-Host '1. Probando transformacion Mayusculas/Minusculas/Nombre Propio en celda...' -ForegroundColor Yellow
    $ws.Range('A1').Value2 = 'juan perez del valle'
    $ws.Range('A1').Select() | Out-Null
    $excel.Run($runPrefix + 'EjecutarMayusculas')
    $resMayus = $ws.Range('A1').Value2
    Write-Host ('   Resultado Mayusculas: ' + $resMayus) -ForegroundColor Green
    if ($resMayus -ne 'JUAN PEREZ DEL VALLE') { throw 'Fallo Mayusculas in-place' }

    $excel.Run($runPrefix + 'EjecutarMinusculas')
    $resMinus = $ws.Range('A1').Value2
    Write-Host ('   Resultado Minusculas: ' + $resMinus) -ForegroundColor Green
    if ($resMinus -ne 'juan perez del valle') { throw 'Fallo Minusculas in-place' }

    $excel.Run($runPrefix + 'EjecutarNombrePropio')
    $resProper = $ws.Range('A1').Value2
    Write-Host ('   Resultado Nombre Propio: ' + $resProper) -ForegroundColor Green
    if ($resProper -ne 'Juan Perez Del Valle') { throw 'Fallo Nombre Propio in-place' }

    $excel.Run($runPrefix + 'EjecutarTipoOracion')
    $resOracion = $ws.Range('A1').Value2
    Write-Host ('   Resultado Tipo Oracion: ' + $resOracion) -ForegroundColor Green
    if ($resOracion -ne 'Juan perez del valle') { throw 'Fallo Tipo Oracion in-place' }

    # 2. Probar rango de filas continuo (A5:A9)
    Write-Host '2. Probando rango de filas continuo (A5:A9)...' -ForegroundColor Yellow
    $strMex = '   m' + [char]0xE9 + 'xico   '
    $strPeru = '   per' + [char]0xFA + '   '
    $strEsp = '   espa' + [char]0xF1 + 'a   '
    
    $ws.Range('A5').Value2 = $strMex
    $ws.Range('A6').Value2 = '   colombia   '
    $ws.Range('A7').Value2 = '   argentina   '
    $ws.Range('A8').Value2 = $strPeru
    $ws.Range('A9').Value2 = $strEsp

    $ws.Range('A5:A9').Select() | Out-Null
    $excel.Run($runPrefix + 'EjecutarLimpiezaTotal')
    $excel.Run($runPrefix + 'EjecutarMayusculas')

    $r5 = $ws.Range('A5').Value2
    $r8 = $ws.Range('A8').Value2
    $r9 = $ws.Range('A9').Value2

    $expectedEsp = 'ESPA' + [char]0xD1 + 'A'
    Write-Host ('   Fila 5: ' + $r5 + ' (esperado: MEXICO)') -ForegroundColor Green
    Write-Host ('   Fila 8: ' + $r8 + ' (esperado: PERU)') -ForegroundColor Green
    Write-Host ('   Fila 9: ' + $r9 + ' (esperado: ' + $expectedEsp + ')') -ForegroundColor Green

    if ($r5 -ne 'MEXICO' -or $r8 -ne 'PERU' -or $r9 -ne $expectedEsp) {
        throw 'Fallo en transformacion de rango de filas'
    }

    # 3. Probar fila individual
    Write-Host '3. Probando fila individual aislada (B15:D15)...' -ForegroundColor Yellow
    $ws.Range('B15').Value2 = '  ana  '
    $ws.Range('C15').Value2 = ('  mar' + [char]0xED + 'a  ')
    $ws.Range('D15').Value2 = ('  g' + [char]0xF3 + 'mez  ')

    $ws.Range('B15:D15').Select() | Out-Null
    $excel.Run($runPrefix + 'EjecutarLimpiarEspacios')
    $excel.Run($runPrefix + 'EjecutarNombrePropio')

    $b15 = $ws.Range('B15').Value2
    $c15 = $ws.Range('C15').Value2
    $d15 = $ws.Range('D15').Value2

    Write-Host ('   Fila 15: ' + $b15 + ' ' + $c15 + ' ' + $d15) -ForegroundColor Green
    if ($b15 -ne 'Ana') {
        throw 'Fallo en fila individual'
    }

    # 4. Probar eliminacion de caracteres raros de base de datos
    Write-Host '4. Probando eliminacion de caracteres de control y raros...' -ForegroundColor Yellow
    $dirtyStr = 'Producto' + [char]7 + [char]31 + '123' + [char]0x200B + ' '
    $ws.Range('A20').Value2 = $dirtyStr
    $ws.Range('A20').Select() | Out-Null
    $excel.Run($runPrefix + 'EjecutarQuitarRaros')
    $rDirty = $ws.Range('A20').Value2
    Write-Host ('   Resultado Caracteres Raros: [' + $rDirty + '] (esperado: [Producto123 ])') -ForegroundColor Green
    if ($rDirty -ne 'Producto123 ') {
        throw 'Fallo en eliminacion de caracteres raros'
    }

    Write-Host ''
    Write-Host '==========================================================' -ForegroundColor Green
    Write-Host ' TODAS LAS PRUEBAS DE INTEGRACION PASARON EXITOSAMENTE!   ' -ForegroundColor Green
    Write-Host '==========================================================' -ForegroundColor Green
}
catch {
    Write-Host ('Error en la validacion: ' + $_.Exception.Message) -ForegroundColor Red
    throw
}
finally {
    if ($null -ne $testWb) { $testWb.Close($false) }
    if ($null -ne $addinWb) { $addinWb.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $testWb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testWb) | Out-Null }
    if ($null -ne $addinWb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($addinWb) | Out-Null }
    if ($null -ne $excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
