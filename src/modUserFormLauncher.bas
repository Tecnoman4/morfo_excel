Attribute VB_Name = "modUserFormLauncher"
Option Explicit

' ===============================================================================
' MÓDULO: modUserFormLauncher
' PROPÓSITO: Procedimientos para invocar el formulario flotante de usuario
' ===============================================================================

Public Sub AbrirFormularioAvanzado()
    On Error Resume Next
    frmAdvancedClean.Show vbModeless
    If Err.Number <> 0 Then
        ' Si la versión de Excel no soporta vbModeless en algún contexto, abrir modal
        Err.Clear
        frmAdvancedClean.Show vbModal
    End If
End Sub
