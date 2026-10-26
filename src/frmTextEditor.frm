VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTextEditor 
   Caption         =   "Morfo - Editor y Transformador de Texto Seleccionado"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTextEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ===============================================================================
' FORMULARIO: frmTextEditor
' PROPÓSITO: Permite editar y transformar fragmentos específicos de texto
'            seleccionados dentro de una celda con párrafos largos.
' ===============================================================================

Private lblHeader As Object
Private txtContent As Object
Private lblSelectionInfo As Object

Private grpActions As Object
Private btnUpper As Object
Private btnLower As Object
Private btnProper As Object
Private btnSentence As Object
Private btnAccents As Object
Private btnSpaces As Object

Private lblShortcutTip As Object
Private btnApply As Object
Private btnCancel As Object

Private targetCell As Range
Private lastCycleState As Integer ' 0 = none, 1 = upper, 2 = lower, 3 = proper
Private savedSelStart As Long
Private savedSelLength As Long

Private Sub UserForm_Initialize()
    On Error Resume Next
    
    Me.Width = 540
    Me.Height = 440
    savedSelStart = 0
    savedSelLength = 0
    
    If TypeName(Selection) = "Range" Then
        Set targetCell = ActiveCell
    End If
    
    ' 1. Encabezado informativo
    Set lblHeader = Me.Controls.Add("Forms.Label.1", "lblHeader", True)
    With lblHeader
        If Not targetCell Is Nothing Then
            .Caption = "Celda activa: [" & targetCell.Address(False, False) & "]  -  Selecciona con el ratón el fragmento que deseas transformar:"
        Else
            .Caption = "Selecciona con el ratón el fragmento de texto que deseas transformar:"
        End If
        .Left = 14
        .Top = 10
        .Width = 495
        .Height = 16
        .Font.Bold = True
        .ForeColor = RGB(30, 30, 30)
    End With
    
    ' 2. Área de Texto Multilínea (Editor)
    Set txtContent = Me.Controls.Add("Forms.TextBox.1", "txtContent", True)
    With txtContent
        .Left = 14
        .Top = 32
        .Width = 495
        .Height = 210
        .MultiLine = True
        .ScrollBars = 2 ' fmScrollBarsVertical
        .EnterKeyBehavior = True
        .HideSelection = False ' Mantiene visible la selección aunque otro control tenga foco
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(20, 20, 20)
        
        If Not targetCell Is Nothing Then
            If Not IsEmpty(targetCell.Value2) Then
                .Text = CStr(targetCell.Value)
            Else
                .Text = ""
            End If
        End If
    End With
    
    ' 3. Información de Selección
    Set lblSelectionInfo = Me.Controls.Add("Forms.Label.1", "lblSelectionInfo", True)
    With lblSelectionInfo
        .Left = 14
        .Top = 247
        .Width = 495
        .Height = 14
        .ForeColor = RGB(100, 100, 100)
        .Caption = "Total: " & Len(txtContent.Text) & " caracteres | Ningún texto seleccionado (se transformará todo si pulsas un botón)"
    End With
    
    ' 4. Marco visual para acciones
    Set grpActions = Me.Controls.Add("Forms.Frame.1", "grpActions", True)
    With grpActions
        .Caption = " Transformar Selección de Texto "
        .Left = 14
        .Top = 266
        .Width = 495
        .Height = 65
    End With
    
    ' Botones agregados a Me.Controls para que sus eventos Click funcionen de forma nativa
    Set btnUpper = Me.Controls.Add("Forms.CommandButton.1", "btnUpper", True)
    With btnUpper
        .Caption = "MAYÚSCULAS"
        .Left = 24
        .Top = 286
        .Width = 84
        .Height = 32
        .Font.Bold = True
        .TakeFocusOnClick = False ' No deselecciona el texto del TextBox al hacer clic
    End With
    
    Set btnLower = Me.Controls.Add("Forms.CommandButton.1", "btnLower", True)
    With btnLower
        .Caption = "minúsculas"
        .Left = 112
        .Top = 286
        .Width = 74
        .Height = 32
        .TakeFocusOnClick = False
    End With
    
    Set btnProper = Me.Controls.Add("Forms.CommandButton.1", "btnProper", True)
    With btnProper
        .Caption = "Nombre Propio"
        .Left = 190
        .Top = 286
        .Width = 92
        .Height = 32
        .TakeFocusOnClick = False
    End With
    
    Set btnSentence = Me.Controls.Add("Forms.CommandButton.1", "btnSentence", True)
    With btnSentence
        .Caption = "Tipo Oración"
        .Left = 286
        .Top = 286
        .Width = 80
        .Height = 32
        .TakeFocusOnClick = False
    End With
    
    Set btnAccents = Me.Controls.Add("Forms.CommandButton.1", "btnAccents", True)
    With btnAccents
        .Caption = "Quitar Acentos"
        .Left = 370
        .Top = 286
        .Width = 86
        .Height = 32
        .TakeFocusOnClick = False
    End With
    
    Set btnSpaces = Me.Controls.Add("Forms.CommandButton.1", "btnSpaces", True)
    With btnSpaces
        .Caption = "Espacios"
        .Left = 460
        .Top = 286
        .Width = 44
        .Height = 32
        .TakeFocusOnClick = False
    End With
    
    ' 5. Consejo y atajos
    Set lblShortcutTip = Me.Controls.Add("Forms.Label.1", "lblShortcutTip", True)
    With lblShortcutTip
        .Left = 14
        .Top = 340
        .Width = 495
        .Height = 15
        .ForeColor = RGB(90, 110, 130)
        .Caption = "Atajos: Puedes presionar Shift + F3 dentro del texto para alternar mayúsculas/minúsculas."
    End With
    
    ' 6. Botones de Aplicar y Cancelar
    Set btnApply = Me.Controls.Add("Forms.CommandButton.1", "btnApply", True)
    With btnApply
        .Caption = "Guardar en Celda"
        .Left = 245
        .Top = 362
        .Width = 160
        .Height = 34
        .Font.Bold = True
    End With
    
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    With btnCancel
        .Caption = "Cancelar"
        .Left = 414
        .Top = 362
        .Width = 95
        .Height = 34
    End With
End Sub

Private Sub TrackSelection()
    On Error Resume Next
    If Not txtContent Is Nothing Then
        If txtContent.SelLength > 0 Then
            savedSelStart = txtContent.SelStart
            savedSelLength = txtContent.SelLength
        Else
            savedSelStart = 0
            savedSelLength = 0
        End If
        UpdateSelectionLabel
    End If
End Sub

Private Sub UpdateSelectionLabel()
    On Error Resume Next
    Dim currentLen As Long
    If txtContent.SelLength > 0 Then
        currentLen = txtContent.SelLength
    ElseIf savedSelLength > 0 Then
        currentLen = savedSelLength
    Else
        currentLen = 0
    End If
    
    If currentLen > 0 Then
        lblSelectionInfo.Caption = "Seleccionados: " & currentLen & " caracteres de " & Len(txtContent.Text) & " | ¡La acción se aplicará ÚNICAMENTE al fragmento seleccionado!"
        lblSelectionInfo.ForeColor = RGB(15, 120, 45)
    Else
        lblSelectionInfo.Caption = "Total: " & Len(txtContent.Text) & " caracteres | Ningún texto seleccionado (se transformará todo si pulsas un botón)"
        lblSelectionInfo.ForeColor = RGB(100, 100, 100)
    End If
End Sub

Private Sub txtContent_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TrackSelection
End Sub

Private Sub txtContent_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    TrackSelection
End Sub

Private Sub txtContent_Change()
    TrackSelection
End Sub

Private Sub txtContent_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    ' Shift + F3: Ciclar Mayúsculas / minúsculas / Nombre Propio igual que en Word
    If Shift = 1 And KeyCode = 114 Then ' 114 = vbKeyF3
        KeyCode = 0
        CycleCaseSelection
        Exit Sub
    End If
    
    ' Ctrl + Enter: Guardar y cerrar
    If Shift = 2 And KeyCode = 13 Then ' 13 = vbKeyReturn
        KeyCode = 0
        btnApply_Click
        Exit Sub
    End If
    
    ' Esc: Cancelar
    If KeyCode = 27 Then ' 27 = vbKeyEscape
        KeyCode = 0
        Unload Me
        Exit Sub
    End If
End Sub

Public Sub TransformFragment(ByVal transType As Integer)
    On Error Resume Next
    Dim fullText As String
    Dim sStart As Long, sLen As Long
    Dim fragment As String
    Dim transFragment As String
    
    fullText = txtContent.Text
    
    If txtContent.SelLength > 0 Then
        sStart = txtContent.SelStart
        sLen = txtContent.SelLength
    ElseIf savedSelLength > 0 Then
        sStart = savedSelStart
        sLen = savedSelLength
    Else
        sStart = 0
        sLen = 0
    End If
    
    If sLen <= 0 Then
        ' No hay fragmento específico seleccionado: transformar TODO el texto
        transFragment = ApplyTypeToString(fullText, transType)
        txtContent.Text = transFragment
        sStart = 0
        sLen = Len(transFragment)
    Else
        ' Transformar ÚNICAMENTE el fragmento seleccionado
        fragment = Mid$(fullText, sStart + 1, sLen)
        transFragment = ApplyTypeToString(fragment, transType)
        
        fullText = Left$(fullText, sStart) & transFragment & Mid$(fullText, sStart + sLen + 1)
        txtContent.Text = fullText
        sLen = Len(transFragment)
    End If
    
    ' Mantener el foco y la selección activa sobre el texto transformado
    txtContent.SetFocus
    txtContent.SelStart = sStart
    txtContent.SelLength = sLen
    savedSelStart = sStart
    savedSelLength = sLen
    
    UpdateSelectionLabel
End Sub

Private Function ApplyTypeToString(ByVal txt As String, ByVal tType As Integer) As String
    Select Case tType
        Case 1 ' Mayúsculas
            ApplyTypeToString = modTextTransform.TextToUpper(txt)
        Case 2 ' Minúsculas
            ApplyTypeToString = modTextTransform.TextToLower(txt)
        Case 3 ' Nombre Propio
            ApplyTypeToString = modTextTransform.TextToProper(txt)
        Case 4 ' Tipo Oración
            ApplyTypeToString = modTextTransform.TextToSentence(txt)
        Case 5 ' Quitar Acentos
            ApplyTypeToString = modDataClean.RemoveAccentsFromString(txt, True)
        Case 6 ' Limpiar Espacios
            ApplyTypeToString = modDataClean.CleanAllSpaces(txt)
        Case 7 ' Limpieza Total
            ApplyTypeToString = modArrayProcessor.TransformString(txt, CLEAN_ALL)
        Case Else
            ApplyTypeToString = txt
    End Select
End Function

Private Sub CycleCaseSelection()
    If lastCycleState = 1 Then
        TransformFragment 2 ' A minúsculas
        lastCycleState = 2
    ElseIf lastCycleState = 2 Then
        TransformFragment 3 ' A Nombre Propio
        lastCycleState = 3
    Else
        TransformFragment 1 ' A MAYÚSCULAS
        lastCycleState = 1
    End If
End Sub

Private Sub btnUpper_Click()
    lastCycleState = 1
    TransformFragment 1
End Sub

Private Sub btnLower_Click()
    lastCycleState = 2
    TransformFragment 2
End Sub

Private Sub btnProper_Click()
    lastCycleState = 3
    TransformFragment 3
End Sub

Private Sub btnSentence_Click()
    TransformFragment 4
End Sub

Private Sub btnAccents_Click()
    TransformFragment 5
End Sub

Private Sub btnSpaces_Click()
    TransformFragment 6
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnApply_Click()
    On Error Resume Next
    If targetCell Is Nothing Then
        If TypeName(Selection) = "Range" Then
            Set targetCell = ActiveCell
        End If
    End If
    
    If Not targetCell Is Nothing Then
        targetCell.Value = txtContent.Text
    End If
    
    Unload Me
End Sub
