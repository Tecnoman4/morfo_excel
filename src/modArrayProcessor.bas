Attribute VB_Name = "modArrayProcessor"
Option Explicit

' ===============================================================================
' MÓDULO: modArrayProcessor
' PROPÓSITO: Motor de alto rendimiento para procesar celdas individuales,
'            filas completas, rangos continuos o discontinuos (múltiples áreas).
' ===============================================================================

Public Enum TransformationType
    TRANSFORM_UPPER = 1
    TRANSFORM_LOWER = 2
    TRANSFORM_PROPER = 3
    CLEAN_SPACES = 4
    REMOVE_ACCENTS = 5
    REMOVE_RARE_CHARS = 6
    CLEAN_ALL = 7
    TRANSFORM_SENTENCE = 8
End Enum

Public Sub ApplyTransformationToSelection(ByVal transType As TransformationType, Optional ByVal preserveEnie As Boolean = True, Optional ByVal removeLineBreaks As Boolean = False)
    Dim targetRange As Range
    
    If TypeName(Selection) <> "Range" Then
        If Application.Visible Then
            MsgBox "Por favor, seleccione una celda, fila o rango de celdas válido en la hoja.", vbExclamation, "Morfo"
        End If
        Exit Sub
    End If
    
    Set targetRange = Selection
    ProcessRange targetRange, transType, preserveEnie, removeLineBreaks
End Sub

Public Sub ProcessRange(ByVal rng As Range, ByVal transType As TransformationType, Optional ByVal preserveEnie As Boolean = True, Optional ByVal removeLineBreaks As Boolean = False)
    Dim ws As Worksheet
    Dim validRange As Range
    Dim area As Range
    Dim dataArr As Variant
    Dim r As Long, c As Long
    Dim numRows As Long, numCols As Long
    Dim cellValue As Variant
    Dim transformedValue As String
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalEnableEvents As Boolean
    Dim changedCount As Long
    
    Set ws = rng.Worksheet
    
    ' Optimización crucial: Si se seleccionan filas o columnas enteras, limitar a UsedRange
    If rng.CountLarge > 500000 Or rng.Rows.CountLarge = ws.Rows.CountLarge Or rng.Columns.CountLarge = ws.Columns.CountLarge Then
        If ws.UsedRange Is Nothing Then Exit Sub
        Set validRange = Intersect(rng, ws.UsedRange)
        If validRange Is Nothing Then Exit Sub
    Else
        Set validRange = rng
    End If
    
    ' Guardar configuración original del entorno Excel
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalEnableEvents = Application.EnableEvents
    
    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    changedCount = 0
    
    ' Procesar cada área seleccionada (soporta rangos no continuos con Ctrl)
    For Each area In validRange.Areas
        numRows = area.Rows.Count
        numCols = area.Columns.Count
        
        If numRows = 1 And numCols = 1 Then
            ' Caso de una sola celda
            If Not area.HasFormula And Not IsEmpty(area.Value2) Then
                cellValue = area.Value2
                If VarType(cellValue) = vbString Then
                    transformedValue = TransformString(CStr(cellValue), transType, preserveEnie, removeLineBreaks)
                    If transformedValue <> CStr(cellValue) Then
                        area.Value2 = transformedValue
                        changedCount = changedCount + 1
                    End If
                End If
            End If
        Else
            ' Procesamiento en memoria por matriz (Array) para filas individuales o rangos de filas
            dataArr = area.Value2
            
            If IsArray(dataArr) Then
                For r = 1 To numRows
                    For c = 1 To numCols
                        cellValue = dataArr(r, c)
                        If Not IsEmpty(cellValue) Then
                            If VarType(cellValue) = vbString Then
                                transformedValue = TransformString(CStr(cellValue), transType, preserveEnie, removeLineBreaks)
                                If transformedValue <> CStr(cellValue) Then
                                    dataArr(r, c) = transformedValue
                                    changedCount = changedCount + 1
                                End If
                            End If
                        End If
                    Next c
                Next r
                
                ' Volcar la matriz completa de vuelta a la hoja en una sola operación ultrarrápida
                area.Value2 = dataArr
            End If
        End If
    Next area

CleanUp:
    ' Restaurar configuración original
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents
    
    If Application.Visible Then
        If Err.Number <> 0 Then
            MsgBox "Ocurrió un error durante el procesamiento: " & Err.Description, vbCritical, "Error"
        Else
            Application.StatusBar = "Morfo: " & changedCount & " celdas actualizadas correctamente."
        End If
    End If
End Sub

Public Function TransformString(ByVal inputTxt As String, ByVal transType As TransformationType, Optional ByVal preserveEnie As Boolean = True, Optional ByVal removeLineBreaks As Boolean = False) As String
    Dim result As String
    result = inputTxt
    
    Select Case transType
        Case TRANSFORM_UPPER
            result = modTextTransform.TextToUpper(result)
            
        Case TRANSFORM_LOWER
            result = modTextTransform.TextToLower(result)
            
        Case TRANSFORM_PROPER
            result = modTextTransform.TextToProper(result)
            
        Case TRANSFORM_SENTENCE
            result = modTextTransform.TextToSentence(result)
            
        Case CLEAN_SPACES
            result = modDataClean.CleanAllSpaces(result)
            
        Case REMOVE_ACCENTS
            result = modDataClean.RemoveAccentsFromString(result, preserveEnie)
            
        Case REMOVE_RARE_CHARS
            result = modDataClean.RemoveRareCharsFromString(result, removeLineBreaks)
            
        Case CLEAN_ALL
            ' Limpieza integral recomendada para bases de datos:
            ' 1. Quitar caracteres de control y raros
            result = modDataClean.RemoveRareCharsFromString(result, removeLineBreaks)
            ' 2. Quitar acentos/tildes
            result = modDataClean.RemoveAccentsFromString(result, preserveEnie)
            ' 3. Limpiar espacios iniciales, finales y dobles
            result = modDataClean.CleanAllSpaces(result)
    End Select
    
    TransformString = result
End Function
