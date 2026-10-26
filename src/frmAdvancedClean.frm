VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdvancedClean 
   Caption         =   "Morfo - Limpieza y Transformación Avanzada"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdvancedClean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Controles dinámicos
Private lblSampleOrig As Object
Private txtOriginal As Object
Private lblSamplePreview As Object
Private txtPreview As Object

Private chkTrimSpaces As Object
Private chkMultipleSpaces As Object
Private chkNbspSpaces As Object

Private chkRemoveAccents As Object
Private chkConvertEnie As Object

Private chkRemoveRareChars As Object
Private chkRemoveLineBreaks As Object

Private optCaseNone As Object
Private optCaseUpper As Object
Private optCaseLower As Object
Private optCaseProper As Object
Private optCaseSentence As Object

Private btnApply As Object
Private btnClose As Object
Private btnOpenEditor As Object

Private isInitializing As Boolean

Private Sub UserForm_Initialize()
    isInitializing = True
    
    Me.Width = 490
    Me.Height = 440
    
    ' --- Marco de Vista Previa ---
    Dim grpPreview As Object
    Set grpPreview = Me.Controls.Add("Forms.Frame.1", "grpPreview", True)
    With grpPreview
        .Caption = " Vista Previa (Muestra de la Celda Activa) "
        .Left = 12
        .Top = 10
        .Width = 450
        .Height = 105
    End With
    
    Set lblSampleOrig = grpPreview.Controls.Add("Forms.Label.1", "lblSampleOrig", True)
    With lblSampleOrig
        .Caption = "Texto Original:"
        .Left = 10
        .Top = 18
        .Width = 85
        .Height = 14
    End With
    
    Set txtOriginal = grpPreview.Controls.Add("Forms.TextBox.1", "txtOriginal", True)
    With txtOriginal
        .Left = 95
        .Top = 16
        .Width = 340
        .Height = 32
        .MultiLine = True
        .Locked = True
        .BackColor = RGB(245, 245, 245)
    End With
    
    Set lblSamplePreview = grpPreview.Controls.Add("Forms.Label.1", "lblSamplePreview", True)
    With lblSamplePreview
        .Caption = "Resultado:"
        .Left = 10
        .Top = 56
        .Width = 85
        .Height = 14
        .Font.Bold = True
    End With
    
    Set txtPreview = grpPreview.Controls.Add("Forms.TextBox.1", "txtPreview", True)
    With txtPreview
        .Left = 95
        .Top = 54
        .Width = 340
        .Height = 32
        .MultiLine = True
        .Locked = True
        .BackColor = RGB(235, 247, 238)
        .ForeColor = RGB(15, 115, 45)
        .Font.Bold = True
    End With
    
    ' --- Marco de Limpieza de Espacios ---
    Dim grpSpaces As Object
    Set grpSpaces = Me.Controls.Add("Forms.Frame.1", "grpSpaces", True)
    With grpSpaces
        .Caption = " Limpieza de Espacios "
        .Left = 12
        .Top = 122
        .Width = 220
        .Height = 85
    End With
    
    Set chkTrimSpaces = grpSpaces.Controls.Add("Forms.CheckBox.1", "chkTrimSpaces", True)
    With chkTrimSpaces
        .Caption = "Recortar extremos (Trim)"
        .Left = 10
        .Top = 16
        .Width = 200
        .Value = True
    End With
    
    Set chkMultipleSpaces = grpSpaces.Controls.Add("Forms.CheckBox.1", "chkMultipleSpaces", True)
    With chkMultipleSpaces
        .Caption = "Convertir dobles espacios en 1"
        .Left = 10
        .Top = 36
        .Width = 200
        .Value = True
    End With
    
    Set chkNbspSpaces = grpSpaces.Controls.Add("Forms.CheckBox.1", "chkNbspSpaces", True)
    With chkNbspSpaces
        .Caption = "Limpiar espacios web/duros (160)"
        .Left = 10
        .Top = 56
        .Width = 200
        .Value = True
    End With
    
    ' --- Marco de Acentos y Caracteres Especiales ---
    Dim grpAccents As Object
    Set grpAccents = Me.Controls.Add("Forms.Frame.1", "grpAccents", True)
    With grpAccents
        .Caption = " Acentos y Tildes "
        .Left = 242
        .Top = 122
        .Width = 220
        .Height = 85
    End With
    
    Set chkRemoveAccents = grpAccents.Controls.Add("Forms.CheckBox.1", "chkRemoveAccents", True)
    With chkRemoveAccents
        .Caption = "Quitar tildes (á->a, é->e...)"
        .Left = 10
        .Top = 16
        .Width = 200
        .Value = True
    End With
    
    Set chkConvertEnie = grpAccents.Controls.Add("Forms.CheckBox.1", "chkConvertEnie", True)
    With chkConvertEnie
        .Caption = "Convertir Ñ en N (opcional)"
        .Left = 10
        .Top = 38
        .Width = 200
        .Value = False
    End With
    
    ' --- Marco de Caracteres Raros y Control ---
    Dim grpRare As Object
    Set grpRare = Me.Controls.Add("Forms.Frame.1", "grpRare", True)
    With grpRare
        .Caption = " Caracteres Raros y de Control "
        .Left = 12
        .Top = 214
        .Width = 220
        .Height = 80
    End With
    
    Set chkRemoveRareChars = grpRare.Controls.Add("Forms.CheckBox.1", "chkRemoveRareChars", True)
    With chkRemoveRareChars
        .Caption = "Eliminar códigos de control / invisibles"
        .Left = 10
        .Top = 16
        .Width = 200
        .Value = True
    End With
    
    Set chkRemoveLineBreaks = grpRare.Controls.Add("Forms.CheckBox.1", "chkRemoveLineBreaks", True)
    With chkRemoveLineBreaks
        .Caption = "Reemplazar saltos de línea por espacio"
        .Left = 10
        .Top = 38
        .Width = 200
        .Value = True
    End With
    
    ' --- Marco de Mayúsculas / Minúsculas ---
    Dim grpCase As Object
    Set grpCase = Me.Controls.Add("Forms.Frame.1", "grpCase", True)
    With grpCase
        .Caption = " Formato de Mayúsculas / Minúsculas "
        .Left = 242
        .Top = 214
        .Width = 220
        .Height = 80
    End With
    
    Set optCaseNone = grpCase.Controls.Add("Forms.OptionButton.1", "optCaseNone", True)
    With optCaseNone
        .Caption = "Mantener actual"
        .Left = 10
        .Top = 16
        .Width = 100
        .GroupName = "CaseGroup"
        .Value = True
    End With
    
    Set optCaseUpper = grpCase.Controls.Add("Forms.OptionButton.1", "optCaseUpper", True)
    With optCaseUpper
        .Caption = "MAYÚSCULAS"
        .Left = 115
        .Top = 16
        .Width = 95
        .GroupName = "CaseGroup"
    End With
    
    Set optCaseLower = grpCase.Controls.Add("Forms.OptionButton.1", "optCaseLower", True)
    With optCaseLower
        .Caption = "minúsculas"
        .Left = 10
        .Top = 38
        .Width = 100
        .GroupName = "CaseGroup"
    End With
    
    Set optCaseProper = grpCase.Controls.Add("Forms.OptionButton.1", "optCaseProper", True)
    With optCaseProper
        .Caption = "Nombre Propio"
        .Left = 115
        .Top = 38
        .Width = 95
        .GroupName = "CaseGroup"
    End With
    
    Set optCaseSentence = grpCase.Controls.Add("Forms.OptionButton.1", "optCaseSentence", True)
    With optCaseSentence
        .Caption = "Tipo Oración (1ª mayús)"
        .Left = 10
        .Top = 60
        .Width = 190
        .GroupName = "CaseGroup"
    End With
    
    ' --- Botones de Acción ---
    Set btnOpenEditor = Me.Controls.Add("Forms.CommandButton.1", "btnOpenEditor", True)
    With btnOpenEditor
        .Caption = "📝 Editor de Párrafo..."
        .Left = 12
        .Top = 310
        .Width = 168
        .Height = 32
    End With
    
    Set btnApply = Me.Controls.Add("Forms.CommandButton.1", "btnApply", True)
    With btnApply
        .Caption = "Aplicar a Selección / Filas"
        .Left = 190
        .Top = 310
        .Width = 160
        .Height = 32
        .Default = True
        .Font.Bold = True
    End With
    
    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    With btnClose
        .Caption = "Cerrar"
        .Left = 360
        .Top = 310
        .Width = 100
        .Height = 32
        .Cancel = True
    End With
    
    ' Cargar muestra de la celda activa
    LoadActiveCellSample
    
    isInitializing = False
    UpdatePreview
End Sub

Private Sub LoadActiveCellSample()
    Dim sample As String
    If TypeName(Selection) = "Range" Then
        If Not IsEmpty(ActiveCell.Value2) Then
            sample = CStr(ActiveCell.Value2)
        Else
            sample = "   Ejemplo: JUAN   PÉREZ  (áéíóú, ñ)   "
        End If
    Else
        sample = "   Ejemplo: JUAN   PÉREZ  (áéíóú, ñ)   "
    End If
    txtOriginal.Text = sample
End Sub

Public Sub UpdatePreview()
    If isInitializing Then Exit Sub
    Dim txt As String
    txt = txtOriginal.Text
    
    ' 1. Caracteres raros
    If chkRemoveRareChars.Value Then
        txt = modDataClean.RemoveRareCharsFromString(txt, chkRemoveLineBreaks.Value)
    ElseIf chkRemoveLineBreaks.Value Then
        txt = Replace(txt, vbCrLf, " ")
        txt = Replace(txt, vbLf, " ")
        txt = Replace(txt, vbCr, " ")
    End If
    
    ' 2. Acentos
    If chkRemoveAccents.Value Then
        txt = modDataClean.RemoveAccentsFromString(txt, Not chkConvertEnie.Value)
    ElseIf chkConvertEnie.Value Then
        txt = Replace(txt, "ñ", "n")
        txt = Replace(txt, "Ñ", "N")
    End If
    
    ' 3. Espacios
    If chkNbspSpaces.Value Then
        txt = Replace(txt, Chr$(160), " ")
        txt = Replace(txt, ChrW$(&HA0), " ")
    End If
    If chkMultipleSpaces.Value Then
        Do While InStr(txt, "  ") > 0
            txt = Replace(txt, "  ", " ")
        Loop
    End If
    If chkTrimSpaces.Value Then
        txt = Trim$(txt)
    End If
    
    ' 4. Formato de mayúsculas/minúsculas
    If optCaseUpper.Value Then
        txt = modTextTransform.TextToUpper(txt)
    ElseIf optCaseLower.Value Then
        txt = modTextTransform.TextToLower(txt)
    ElseIf optCaseProper.Value Then
        txt = modTextTransform.TextToProper(txt)
    ElseIf optCaseSentence.Value Then
        txt = modTextTransform.TextToSentence(txt)
    End If
    
    txtPreview.Text = txt
End Sub

' Eventos para refrescar la vista previa en vivo
Private Sub chkTrimSpaces_Click()
    UpdatePreview
End Sub

Private Sub chkMultipleSpaces_Click()
    UpdatePreview
End Sub

Private Sub chkNbspSpaces_Click()
    UpdatePreview
End Sub

Private Sub chkRemoveAccents_Click()
    UpdatePreview
End Sub

Private Sub chkConvertEnie_Click()
    UpdatePreview
End Sub

Private Sub chkRemoveRareChars_Click()
    UpdatePreview
End Sub

Private Sub chkRemoveLineBreaks_Click()
    UpdatePreview
End Sub

Private Sub optCaseNone_Click()
    UpdatePreview
End Sub

Private Sub optCaseUpper_Click()
    UpdatePreview
End Sub

Private Sub optCaseLower_Click()
    UpdatePreview
End Sub

Private Sub optCaseProper_Click()
    UpdatePreview
End Sub

Private Sub optCaseSentence_Click()
    UpdatePreview
End Sub

Private Sub btnOpenEditor_Click()
    Unload Me
    modUserFormLauncher.AbrirEditorTexto
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnApply_Click()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Seleccione un rango o fila válido antes de aplicar.", vbExclamation, "Atención"
        Exit Sub
    End If
    
    Dim doTrim As Boolean, doMultipleSpaces As Boolean, doNbsp As Boolean
    Dim doAccents As Boolean, doEnie As Boolean
    Dim doRare As Boolean, doLineBreaks As Boolean
    Dim caseMode As Integer
    
    doTrim = chkTrimSpaces.Value
    doMultipleSpaces = chkMultipleSpaces.Value
    doNbsp = chkNbspSpaces.Value
    doAccents = chkRemoveAccents.Value
    doEnie = chkConvertEnie.Value
    doRare = chkRemoveRareChars.Value
    doLineBreaks = chkRemoveLineBreaks.Value
    
    If optCaseUpper.Value Then
        caseMode = 1
    ElseIf optCaseLower.Value Then
        caseMode = 2
    ElseIf optCaseProper.Value Then
        caseMode = 3
    ElseIf optCaseSentence.Value Then
        caseMode = 4
    Else
        caseMode = 0
    End If
    
    ApplyCustomTransformation Selection, doTrim, doMultipleSpaces, doNbsp, doAccents, doEnie, doRare, doLineBreaks, caseMode
    
    MsgBox "Transformación aplicada exitosamente a la selección.", vbInformation, "Completado"
    LoadActiveCellSample
    UpdatePreview
End Sub

Private Sub ApplyCustomTransformation(ByVal rng As Range, ByVal doTrim As Boolean, ByVal doMulti As Boolean, ByVal doNbsp As Boolean, _
                                     ByVal doAccents As Boolean, ByVal doEnie As Boolean, ByVal doRare As Boolean, ByVal doLineBreaks As Boolean, _
                                     ByVal caseMode As Integer)
    Dim ws As Worksheet
    Dim validRange As Range
    Dim area As Range
    Dim dataArr As Variant
    Dim r As Long, c As Long
    Dim numRows As Long, numCols As Long
    Dim cellVal As Variant
    Dim sVal As String
    
    Dim origScreen As Boolean, origCalc As XlCalculation, origEvents As Boolean
    
    Set ws = rng.Worksheet
    If rng.CountLarge > 500000 Or rng.Rows.CountLarge = ws.Rows.CountLarge Or rng.Columns.CountLarge = ws.Columns.CountLarge Then
        If ws.UsedRange Is Nothing Then Exit Sub
        Set validRange = Intersect(rng, ws.UsedRange)
        If validRange Is Nothing Then Exit Sub
    Else
        Set validRange = rng
    End If
    
    origScreen = Application.ScreenUpdating
    origCalc = Application.Calculation
    origEvents = Application.EnableEvents
    
    On Error GoTo CleanUpCustom
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    For Each area In validRange.Areas
        numRows = area.Rows.Count
        numCols = area.Columns.Count
        
        If numRows = 1 And numCols = 1 Then
            If Not area.HasFormula And Not IsEmpty(area.Value2) Then
                cellVal = area.Value2
                If VarType(cellVal) = vbString Then
                    area.Value2 = TransformSingleCustom(CStr(cellVal), doTrim, doMulti, doNbsp, doAccents, doEnie, doRare, doLineBreaks, caseMode)
                End If
            End If
        Else
            dataArr = area.Value2
            If IsArray(dataArr) Then
                For r = 1 To numRows
                    For c = 1 To numCols
                        cellVal = dataArr(r, c)
                        If Not IsEmpty(cellVal) Then
                            If VarType(cellVal) = vbString Then
                                dataArr(r, c) = TransformSingleCustom(CStr(cellVal), doTrim, doMulti, doNbsp, doAccents, doEnie, doRare, doLineBreaks, caseMode)
                            End If
                        End If
                    Next c
                Next r
                area.Value2 = dataArr
            End If
        End If
    Next area

CleanUpCustom:
    Application.ScreenUpdating = origScreen
    Application.Calculation = origCalc
    Application.EnableEvents = origEvents
End Sub

Private Function TransformSingleCustom(ByVal txt As String, ByVal doTrim As Boolean, ByVal doMulti As Boolean, ByVal doNbsp As Boolean, _
                                      ByVal doAccents As Boolean, ByVal doEnie As Boolean, ByVal doRare As Boolean, ByVal doLineBreaks As Boolean, _
                                      ByVal caseMode As Integer) As String
    Dim res As String
    res = txt
    
    If doRare Then
        res = modDataClean.RemoveRareCharsFromString(res, doLineBreaks)
    ElseIf doLineBreaks Then
        res = Replace(res, vbCrLf, " ")
        res = Replace(res, vbLf, " ")
        res = Replace(res, vbCr, " ")
    End If
    
    If doAccents Then
        res = modDataClean.RemoveAccentsFromString(res, Not doEnie)
    ElseIf doEnie Then
        res = Replace(res, "ñ", "n")
        res = Replace(res, "Ñ", "N")
    End If
    
    If doNbsp Then
        res = Replace(res, Chr$(160), " ")
        res = Replace(res, ChrW$(&HA0), " ")
    End If
    If doMulti Then
        Do While InStr(res, "  ") > 0
            res = Replace(res, "  ", " ")
        Loop
    End If
    If doTrim Then
        res = Trim$(res)
    End If
    
    Select Case caseMode
        Case 1: res = modTextTransform.TextToUpper(res)
        Case 2: res = modTextTransform.TextToLower(res)
        Case 3: res = modTextTransform.TextToProper(res)
        Case 4: res = modTextTransform.TextToSentence(res)
    End Select
    
    TransformSingleCustom = res
End Function
