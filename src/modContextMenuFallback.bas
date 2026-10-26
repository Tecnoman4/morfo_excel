Attribute VB_Name = "modContextMenuFallback"
Option Explicit

' ===============================================================================
' MÓDULO: modContextMenuFallback
' PROPÓSITO: Creación y limpieza de opciones en el menú contextual de clic derecho
'            utilizando CommandBars("Cell") para máxima compatibilidad.
' ===============================================================================

Private Const MENU_TAG As String = "MorfoAddinMenu"

Public Sub Auto_Open()
    On Error Resume Next
    AgregarMenuContextual
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    EliminarMenuContextual
End Sub

Public Sub AgregarMenuContextual()
    If Not Application.Visible Then Exit Sub
    Dim cbCell As CommandBar
    Dim menuPopup As CommandBarPopup
    Dim btn As CommandBarButton
    
    EliminarMenuContextual
    
    On Error Resume Next
    ' Hay dos CommandBars "Cell": 1 para vista normal y 2 para vista previa de salto de página
    Dim cb As Variant
    For Each cb In Array("Cell", "Row", "Column")
        Set cbCell = Application.CommandBars(cb)
        If Not cbCell Is Nothing Then
            Set menuPopup = cbCell.Controls.Add(Type:=msoControlPopup, Temporary:=True)
            With menuPopup
                .Caption = "Morfo: Limpieza y Texto"
                .Tag = MENU_TAG
                .BeginGroup = True
            End With
            
            ' Opción Mayúsculas
            Set btn = menuPopup.Controls.Add(Type:=msoControlButton)
            With btn
                .Caption = "🔤 MAYÚSCULAS"
                .OnAction = "EjecutarMayusculas"
                .FaceId = 487
            End With
            
            ' Opción Minúsculas
            Set btn = menuPopup.Controls.Add(Type:=msoControlButton)
            With btn
                .Caption = "🔡 minúsculas"
                .OnAction = "EjecutarMinusculas"
                .FaceId = 488
            End With
            
            ' Opción Nombre Propio
            Set btn = menuPopup.Controls.Add(Type:=msoControlButton)
            With btn
                .Caption = "🔠 Nombre Propio"
                .OnAction = "EjecutarNombrePropio"
                .FaceId = 489
            End With
            
            ' Opción Tipo Oración
            Set btn = menuPopup.Controls.Add(Type:=msoControlButton)
            With btn
                .Caption = "📝 Tipo Oración"
                .OnAction = "EjecutarTipoOracion"
                .FaceId = 486
            End With
            
            ' Separador
            Set btn = menuPopup.Controls.Add(Type:=msoControlButton)
            With btn
                .Caption = "✂️ Quitar Espacios Innecesarios"
                .OnAction = "EjecutarLimpiarEspacios"
                .FaceId = 19
                .BeginGroup = True
            End With
            
            ' Opción Quitar Acentos
            Set btn = menuPopup.Controls.Add(Type:=msoControlButton)
            With btn
                .Caption = "✨ Quitar Acentos / Tildes"
                .OnAction = "EjecutarQuitarAcentos"
                .FaceId = 328
            End With
            
            ' Opción Quitar Caracteres Raros
            Set btn = menuPopup.Controls.Add(Type:=msoControlButton)
            With btn
                .Caption = "🛡️ Quitar Caracteres Raros de BD"
                .OnAction = "EjecutarQuitarRaros"
                .FaceId = 2143
            End With
            
            ' Opción Limpieza Total
            Set btn = menuPopup.Controls.Add(Type:=msoControlButton)
            With btn
                .Caption = "⚡ Limpieza Total 1-Clic"
                .OnAction = "EjecutarLimpiezaTotal"
                .FaceId = 459
                .BeginGroup = True
            End With
            
            ' Opción Avanzada con Formulario
            Set btn = menuPopup.Controls.Add(Type:=msoControlButton)
            With btn
                .Caption = "⚙️ Opciones Avanzadas (Vista Previa)..."
                .OnAction = "AbrirFormularioAvanzado"
                .FaceId = 548
            End With
        End If
    Next cb
    On Error GoTo 0
End Sub

Public Sub EliminarMenuContextual()
    If Not Application.Visible Then Exit Sub
    Dim cbName As Variant
    Dim ctrl As CommandBarControl
    
    On Error Resume Next
    For Each cbName In Array("Cell", "Row", "Column")
        Do
            Set ctrl = Application.CommandBars(cbName).FindControl(Tag:=MENU_TAG)
            If ctrl Is Nothing Then Exit Do
            ctrl.Delete
        Loop
    Next cbName
    On Error GoTo 0
End Sub
