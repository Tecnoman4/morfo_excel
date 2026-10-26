Attribute VB_Name = "modTextTransform"
Option Explicit

' ===============================================================================
' MÓDULO: modTextTransform
' PROPÓSITO: Transformaciones de mayúsculas, minúsculas y nombre propio.
'            Soporta ejecución in-place en selecciones (fila individual, rangos)
'            y como funciones de fórmula en celdas (UDF).
' ===============================================================================

' --- Procedimientos de Acción In-Place (para botones del Ribbon / Menú Contextual) ---

Public Sub EjecutarMayusculas()
    modArrayProcessor.ApplyTransformationToSelection TRANSFORM_UPPER
End Sub

Public Sub EjecutarMinusculas()
    modArrayProcessor.ApplyTransformationToSelection TRANSFORM_LOWER
End Sub

Public Sub EjecutarNombrePropio()
    modArrayProcessor.ApplyTransformationToSelection TRANSFORM_PROPER
End Sub

Public Sub EjecutarTipoOracion()
    modArrayProcessor.ApplyTransformationToSelection TRANSFORM_SENTENCE
End Sub

' --- Funciones Lógicas Puras de Transformación ---

Public Function TextToUpper(ByVal txt As String) As String
    If Len(txt) = 0 Then
        TextToUpper = ""
        Exit Function
    End If
    ' UCase en VBA maneja caracteres Unicode y alfabeto español estándar
    TextToUpper = UCase$(txt)
End Function

Public Function TextToLower(ByVal txt As String) As String
    If Len(txt) = 0 Then
        TextToLower = ""
        Exit Function
    End If
    TextToLower = LCase$(txt)
End Function

Public Function TextToProper(ByVal txt As String) As String
    If Len(txt) = 0 Then
        TextToProper = ""
        Exit Function
    End If
    ' StrConv con vbProperCase convierte a tipo título respetando palabras
    TextToProper = StrConv(txt, vbProperCase)
End Function

Public Function TextToSentence(ByVal txt As String) As String
    If Len(txt) = 0 Then
        TextToSentence = ""
        Exit Function
    End If
    
    Dim res As String
    Dim i As Long, n As Long
    Dim ch As String
    Dim code As Long
    Dim capitalizeNext As Boolean
    
    res = LCase$(txt)
    n = Len(res)
    capitalizeNext = True
    
    For i = 1 To n
        ch = Mid$(res, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        
        If capitalizeNext Then
            ' Letras básicas a-z y vocales acentuadas / ñ
            If (ch >= "a" And ch <= "z") Or code = 225 Or code = 233 Or code = 237 Or code = 243 Or code = 250 Or code = 252 Or code = 241 Then
                Mid$(res, i, 1) = UCase$(ch)
                capitalizeNext = False
            End If
        Else
            ' Detectar fin de oración para capitalizar la siguiente
            If ch = "." Or ch = "?" Or ch = "!" Or ch = vbLf Or ch = vbCr Then
                capitalizeNext = True
            End If
        End If
    Next i
    
    TextToSentence = res
End Function

' --- Funciones UDF para usar directamente como fórmulas en Excel (=UDF_MAYUSCULAS(A2), etc.) ---

Public Function UDF_MAYUSCULAS(ByVal valor As Variant) As Variant
    If IsError(valor) Then
        UDF_MAYUSCULAS = valor
    ElseIf IsEmpty(valor) Then
        UDF_MAYUSCULAS = ""
    Else
        UDF_MAYUSCULAS = TextToUpper(CStr(valor))
    End If
End Function

Public Function UDF_MINUSCULAS(ByVal valor As Variant) As Variant
    If IsError(valor) Then
        UDF_MINUSCULAS = valor
    ElseIf IsEmpty(valor) Then
        UDF_MINUSCULAS = ""
    Else
        UDF_MINUSCULAS = TextToLower(CStr(valor))
    End If
End Function

Public Function UDF_NOMBRE_PROPIO(ByVal valor As Variant) As Variant
    If IsError(valor) Then
        UDF_NOMBRE_PROPIO = valor
    ElseIf IsEmpty(valor) Then
        UDF_NOMBRE_PROPIO = ""
    Else
        UDF_NOMBRE_PROPIO = TextToProper(CStr(valor))
    End If
End Function

Public Function UDF_TIPO_ORACION(ByVal valor As Variant) As Variant
    If IsError(valor) Then
        UDF_TIPO_ORACION = valor
    ElseIf IsEmpty(valor) Then
        UDF_TIPO_ORACION = ""
    Else
        UDF_TIPO_ORACION = TextToSentence(CStr(valor))
    End If
End Function
