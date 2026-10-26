VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    modContextMenuFallback.AgregarMenuContextual
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modContextMenuFallback.EliminarMenuContextual
End Sub

Private Sub Workbook_AddinInstall()
    On Error Resume Next
    modContextMenuFallback.AgregarMenuContextual
    MsgBox "Complemento Morfo instalado correctamente." & vbCrLf & _
           "Encontrará la nueva pestaña 'Morfo' en la Cinta de Opciones y en el menú de clic derecho.", vbInformation, "Morfo Activado"
End Sub

Private Sub Workbook_AddinUninstall()
    On Error Resume Next
    modContextMenuFallback.EliminarMenuContextual
End Sub
