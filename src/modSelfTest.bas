Attribute VB_Name = "modSelfTest"
Option Explicit

' ===============================================================================
' MÓDULO: modSelfTest
' PROPÓSITO: Suite de pruebas unitarias y de integración para validar
'            transformaciones de fila individual, rangos y limpieza.
' ===============================================================================

Public Function RunTestSuite() As String
    Dim logPath As String
    Dim fNum As Integer
    Dim testStr As String
    Dim outStr As String
    
    logPath = "C:\Users\DELL\Documents\Ansi_excel\tests\debug.log"
    On Error Resume Next
    Kill logPath
    On Error GoTo 0
    
    fNum = FreeFile
    Open logPath For Output As #fNum
    Print #fNum, "Iniciando Test Suite..."
    
    ' 1. Prueba Mayúsculas
    Print #fNum, "Step 1: Mayusculas"
    outStr = modTextTransform.TextToUpper("juan perez del valle")
    If outStr <> "JUAN PEREZ DEL VALLE" Then
        Print #fNum, "FAIL 1"
        Close #fNum
        RunTestSuite = "FAIL: Mayúsculas falló. Obtenido: " & outStr
        Exit Function
    End If
    
    ' 2. Prueba Minúsculas
    Print #fNum, "Step 2: Minusculas"
    outStr = modTextTransform.TextToLower("JUAN PEREZ DEL VALLE")
    If outStr <> "juan perez del valle" Then
        Print #fNum, "FAIL 2"
        Close #fNum
        RunTestSuite = "FAIL: Minúsculas falló. Obtenido: " & outStr
        Exit Function
    End If
    
    ' 3. Prueba Nombre Propio
    Print #fNum, "Step 3: Nombre Propio"
    outStr = modTextTransform.TextToProper("juan perez del valle")
    If outStr <> "Juan Perez Del Valle" Then
        Print #fNum, "FAIL 3"
        Close #fNum
        RunTestSuite = "FAIL: Nombre Propio falló. Obtenido: " & outStr
        Exit Function
    End If
    
    ' 4. Prueba Limpiar Espacios
    Print #fNum, "Step 4: Limpiar Espacios"
    testStr = "   texto   con   muchos   espacios   " & Chr$(160) & "  "
    outStr = modDataClean.CleanAllSpaces(testStr)
    If outStr <> "texto con muchos espacios" Then
        Print #fNum, "FAIL 4"
        Close #fNum
        RunTestSuite = "FAIL: Espacios falló. Obtenido: [" & outStr & "]"
        Exit Function
    End If
    
    ' 5. Prueba Quitar Acentos con protección de Ñ
    Print #fNum, "Step 5: Acentos con Enie"
    testStr = "Ángel, Camión, Pingüino, Corazón y Niños"
    outStr = modDataClean.RemoveAccentsFromString(testStr, preserveEnie:=True)
    If outStr <> "Angel, Camion, Pinguino, Corazon y Niños" Then
        Print #fNum, "FAIL 5"
        Close #fNum
        RunTestSuite = "FAIL: Acentos (con Ñ) falló. Obtenido: " & outStr
        Exit Function
    End If
    
    ' 6. Prueba Quitar Acentos convirtiendo Ñ en N
    Print #fNum, "Step 6: Acentos sin Enie"
    outStr = modDataClean.RemoveAccentsFromString(testStr, preserveEnie:=False)
    If outStr <> "Angel, Camion, Pinguino, Corazon y Ninos" Then
        Print #fNum, "FAIL 6"
        Close #fNum
        RunTestSuite = "FAIL: Acentos (sin Ñ) falló. Obtenido: " & outStr
        Exit Function
    End If
    
    ' 7. Prueba Caracteres Raros / Control de Bases de Datos
    Print #fNum, "Step 7: Caracteres raros"
    testStr = "Base" & Chr$(0) & Chr$(31) & "de" & Chr$(10) & "Datos" & ChrW$(&H200B) & "!"
    outStr = modDataClean.RemoveRareCharsFromString(testStr, removeLineBreaks:=True)
    If outStr <> "Base de Datos !" Then
        Print #fNum, "FAIL 7"
        Close #fNum
        RunTestSuite = "FAIL: Caracteres raros falló. Obtenido: [" & outStr & "]"
        Exit Function
    End If
    
    ' 8. Prueba en Rango de Filas
    Print #fNum, "Step 8: Rango de filas en hoja"
    Dim wbTemp As Workbook
    Dim wsTemp As Worksheet
    
    Set wbTemp = Application.Workbooks.Add
    Set wsTemp = wbTemp.Sheets(1)
    
    wsTemp.Range("A1").Value2 = "  méxico  "
    wsTemp.Range("A2").Value2 = "  colombia  "
    wsTemp.Range("A3").Value2 = "  argentina  "
    wsTemp.Range("A4").Value2 = "  perú  "
    wsTemp.Range("A5").Value2 = "  españa  "
    
    Print #fNum, "Step 8.1: ProcessRange A1:A5"
    modArrayProcessor.ProcessRange wsTemp.Range("A1:A5"), CLEAN_ALL, preserveEnie:=True, removeLineBreaks:=True
    modArrayProcessor.ProcessRange wsTemp.Range("A1:A5"), TRANSFORM_UPPER
    
    Print #fNum, "Step 8.2: Verificando valores"
    If wsTemp.Range("A1").Value2 <> "MEXICO" Or wsTemp.Range("A4").Value2 <> "PERU" Or wsTemp.Range("A5").Value2 <> "ESPAÑA" Then
        wbTemp.Close False
        Print #fNum, "FAIL 8.2"
        Close #fNum
        RunTestSuite = "FAIL: Rango de filas falló."
        Exit Function
    End If
    
    Print #fNum, "Step 8.3: Fila individual C10"
    wsTemp.Range("C10").Value2 = "   ana   maría   gómez   "
    modArrayProcessor.ProcessRange wsTemp.Range("C10"), CLEAN_SPACES
    modArrayProcessor.ProcessRange wsTemp.Range("C10"), TRANSFORM_PROPER
    
    If wsTemp.Range("C10").Value2 <> "Ana María Gómez" Then
        wbTemp.Close False
        Print #fNum, "FAIL 8.3"
        Close #fNum
        RunTestSuite = "FAIL: Fila individual C10 falló: " & wsTemp.Range("C10").Value2
        Exit Function
    End If
    
    wbTemp.Close False
    Print #fNum, "SUCCESS"
    Close #fNum
    
    RunTestSuite = "SUCCESS: Todas las pruebas unitarias y de rangos pasaron con éxito (8/8)."
End Function
