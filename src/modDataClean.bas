Attribute VB_Name = "modDataClean"
Option Explicit

' ===============================================================================
' MÓDULO: modDataClean
' PROPÓSITO: Limpieza y normalización de datos para bases de datos:
'            - Espacios (extremos, duplicados, no separables Chr 160).
'            - Acentos y diacríticos (á, é, í, ó, ú, ü, etc.).
'            - Caracteres raros, invisibles, de control y no imprimibles.
'            - Limpieza integral 1-clic.
' ===============================================================================

' --- Procedimientos de Acción In-Place (para botones del Ribbon / Menú Contextual) ---

Public Sub EjecutarLimpiarEspacios()
    modArrayProcessor.ApplyTransformationToSelection CLEAN_SPACES
End Sub

Public Sub EjecutarQuitarAcentos()
    modArrayProcessor.ApplyTransformationToSelection REMOVE_ACCENTS, preserveEnie:=True
End Sub

Public Sub EjecutarQuitarRaros()
    modArrayProcessor.ApplyTransformationToSelection REMOVE_RARE_CHARS, removeLineBreaks:=True
End Sub

Public Sub EjecutarLimpiezaTotal()
    modArrayProcessor.ApplyTransformationToSelection CLEAN_ALL, preserveEnie:=True, removeLineBreaks:=True
End Sub

' --- 1. LÓGICA DE LIMPIEZA DE ESPACIOS ---

Public Function CleanAllSpaces(ByVal txt As String) As String
    Dim res As String
    If Len(txt) = 0 Then
        CleanAllSpaces = ""
        Exit Function
    End If
    
    res = txt
    
    ' Reemplazar espacios duros o web (ASCII 160 y otros invisibles de internet) por espacios estándar
    res = Replace(res, Chr$(160), " ")
    res = Replace(res, ChrW$(&HA0), " ")
    res = Replace(res, ChrW$(&H2007), " ") ' Figure space
    res = Replace(res, ChrW$(&H202F), " ") ' Narrow no-break space
    res = Replace(res, vbTab, " ")
    
    ' Recortar extremos
    res = Trim$(res)
    
    ' Reducir múltiples espacios internos a uno solo
    Do While InStr(res, "  ") > 0
        res = Replace(res, "  ", " ")
    Loop
    
    CleanAllSpaces = res
End Function

' --- 2. LÓGICA DE ELIMINACIÓN DE ACENTOS Y DIACRÍTICOS ---

Public Function RemoveAccentsFromString(ByVal txt As String, Optional ByVal preserveEnie As Boolean = True) As String
    Dim res As String
    If Len(txt) = 0 Then
        RemoveAccentsFromString = ""
        Exit Function
    End If
    
    res = txt
    
    Dim fromChars As Variant
    Dim toChars As Variant
    Dim i As Long
    
    ' Caracteres acentuados usando ChrW$ para garantizar independencia total del encoding del archivo
    fromChars = Array( _
        ChrW$(225), ChrW$(224), ChrW$(228), ChrW$(226), ChrW$(227), ChrW$(257), "á", "à", "ä", "â", "ã", _
        ChrW$(233), ChrW$(232), ChrW$(235), ChrW$(234), ChrW$(275), "é", "è", "ë", "ê", _
        ChrW$(237), ChrW$(236), ChrW$(239), ChrW$(238), ChrW$(299), "í", "ì", "ï", "î", _
        ChrW$(243), ChrW$(242), ChrW$(246), ChrW$(244), ChrW$(245), ChrW$(333), "ó", "ò", "ö", "ô", "õ", _
        ChrW$(250), ChrW$(249), ChrW$(252), ChrW$(251), ChrW$(363), "ú", "ù", "ü", "û", _
        ChrW$(193), ChrW$(192), ChrW$(196), ChrW$(194), ChrW$(195), ChrW$(256), "Á", "À", "Ä", "Â", "Ã", _
        ChrW$(201), ChrW$(200), ChrW$(203), ChrW$(202), ChrW$(274), "É", "È", "Ë", "Ê", _
        ChrW$(205), ChrW$(204), ChrW$(207), ChrW$(206), ChrW$(298), "Í", "Ì", "Ï", "Î", _
        ChrW$(211), ChrW$(210), ChrW$(214), ChrW$(212), ChrW$(213), ChrW$(332), "Ó", "Ò", "Ö", "Ô", "Õ", _
        ChrW$(218), ChrW$(217), ChrW$(220), ChrW$(219), ChrW$(362), "Ú", "Ù", "Ü", "Û" _
    )
    toChars = Array( _
        "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", _
        "e", "e", "e", "e", "e", "e", "e", "e", "e", _
        "i", "i", "i", "i", "i", "i", "i", "i", "i", _
        "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", _
        "u", "u", "u", "u", "u", "u", "u", "u", "u", _
        "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", "A", _
        "E", "E", "E", "E", "E", "E", "E", "E", "E", _
        "I", "I", "I", "I", "I", "I", "I", "I", "I", _
        "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", "O", _
        "U", "U", "U", "U", "U", "U", "U", "U", "U" _
    )
    
    For i = LBound(fromChars) To UBound(fromChars)
        If InStr(res, fromChars(i)) > 0 Then
            res = Replace(res, fromChars(i), toChars(i))
        End If
    Next i
    
    ' Tratamiento opcional para la letra Ñ
    If Not preserveEnie Then
        res = Replace(res, ChrW$(241), "n")
        res = Replace(res, ChrW$(209), "N")
        res = Replace(res, "ñ", "n")
        res = Replace(res, "Ñ", "N")
    End If
    
    RemoveAccentsFromString = res
End Function

' --- 3. LÓGICA DE ELIMINACIÓN DE CARACTERES RAROS / CONTROL ---

Public Function RemoveRareCharsFromString(ByVal txt As String, Optional ByVal removeLineBreaks As Boolean = True) As String
    Dim res As String
    Dim i As Long
    Dim ch As Long
    Dim sb As String
    
    If Len(txt) = 0 Then
        RemoveRareCharsFromString = ""
        Exit Function
    End If
    
    res = txt
    
    ' Quitar caracteres invisibles comunes (Zero-width spaces, BOM, etc.)
    res = Replace(res, ChrW$(&H200B), "") ' Zero width space
    res = Replace(res, ChrW$(&H200C), "") ' Zero width non-joiner
    res = Replace(res, ChrW$(&H200D), "") ' Zero width joiner
    res = Replace(res, ChrW$(&HFEFF), "") ' Byte Order Mark (BOM)
    res = Replace(res, ChrW$(&H7F), "")   ' DEL char
    
    ' Procesar caracteres de control ASCII (0 al 31)
    sb = ""
    For i = 1 To Len(res)
        ch = AscW(Mid$(res, i, 1))
        
        If ch < 32 Then
            ' Códigos de control
            If ch = 10 Or ch = 13 Then
                ' Saltos de línea (CR / LF)
                If removeLineBreaks Then
                    sb = sb & " "
                Else
                    sb = sb & Mid$(res, i, 1)
                End If
            ElseIf ch = 9 Then
                ' Tabulador
                sb = sb & " "
            Else
                ' Ignorar cualquier otro carácter de control (0 al 8, 11, 12, 14 al 31)
            End If
        Else
            sb = sb & Mid$(res, i, 1)
        End If
    Next i
    
    RemoveRareCharsFromString = sb
End Function

' --- 4. FUNCIONES UDF PARA FÓRMULAS EN HOJA DE CÁLCULO ---

Public Function UDF_LIMPIAR_ESPACIOS(ByVal valor As Variant) As Variant
    If IsError(valor) Then
        UDF_LIMPIAR_ESPACIOS = valor
    ElseIf IsEmpty(valor) Then
        UDF_LIMPIAR_ESPACIOS = ""
    Else
        UDF_LIMPIAR_ESPACIOS = CleanAllSpaces(CStr(valor))
    End If
End Function

Public Function UDF_QUITAR_ACENTOS(ByVal valor As Variant, Optional ByVal conservarEnie As Boolean = True) As Variant
    If IsError(valor) Then
        UDF_QUITAR_ACENTOS = valor
    ElseIf IsEmpty(valor) Then
        UDF_QUITAR_ACENTOS = ""
    Else
        UDF_QUITAR_ACENTOS = RemoveAccentsFromString(CStr(valor), conservarEnie)
    End If
End Function

Public Function UDF_QUITAR_RAROS(ByVal valor As Variant, Optional ByVal quitarSaltos As Boolean = True) As Variant
    If IsError(valor) Then
        UDF_QUITAR_RAROS = valor
    ElseIf IsEmpty(valor) Then
        UDF_QUITAR_RAROS = ""
    Else
        UDF_QUITAR_RAROS = RemoveRareCharsFromString(CStr(valor), quitarSaltos)
    End If
End Function

Public Function UDF_LIMPIEZA_TOTAL(ByVal valor As Variant, Optional ByVal conservarEnie As Boolean = True) As Variant
    Dim temp As String
    If IsError(valor) Then
        UDF_LIMPIEZA_TOTAL = valor
    ElseIf IsEmpty(valor) Then
        UDF_LIMPIEZA_TOTAL = ""
    Else
        temp = RemoveRareCharsFromString(CStr(valor), True)
        temp = RemoveAccentsFromString(temp, conservarEnie)
        temp = CleanAllSpaces(temp)
        UDF_LIMPIEZA_TOTAL = temp
    End If
End Function
