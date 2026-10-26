Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

' ===============================================================================
' MÓDULO: modRibbonCallbacks
' PROPÓSITO: Callbacks oficiales RibbonX para la Cinta de Opciones y Menú Contextual
' FIRMA ESTRICTA: Sub Nombre(control As IRibbonControl)
' ===============================================================================

Public Sub CallbackMayusculas(control As IRibbonControl)
    On Error GoTo ErrHandler
    modTextTransform.EjecutarMayusculas
    Exit Sub
ErrHandler:
    MsgBox "No se pudo aplicar Mayúsculas: " & Err.Description, vbExclamation, "Morfo"
End Sub

Public Sub CallbackMinusculas(control As IRibbonControl)
    On Error GoTo ErrHandler
    modTextTransform.EjecutarMinusculas
    Exit Sub
ErrHandler:
    MsgBox "No se pudo aplicar minúsculas: " & Err.Description, vbExclamation, "Morfo"
End Sub

Public Sub CallbackNombrePropio(control As IRibbonControl)
    On Error GoTo ErrHandler
    modTextTransform.EjecutarNombrePropio
    Exit Sub
ErrHandler:
    MsgBox "No se pudo aplicar Nombre Propio: " & Err.Description, vbExclamation, "Morfo"
End Sub

Public Sub CallbackTipoOracion(control As IRibbonControl)
    On Error GoTo ErrHandler
    modTextTransform.EjecutarTipoOracion
    Exit Sub
ErrHandler:
    MsgBox "No se pudo aplicar Tipo Oración: " & Err.Description, vbExclamation, "Morfo"
End Sub

Public Sub CallbackLimpiarEspacios(control As IRibbonControl)
    On Error GoTo ErrHandler
    modDataClean.EjecutarLimpiarEspacios
    Exit Sub
ErrHandler:
    MsgBox "No se pudo limpiar espacios: " & Err.Description, vbExclamation, "Morfo"
End Sub

Public Sub CallbackQuitarAcentos(control As IRibbonControl)
    On Error GoTo ErrHandler
    modDataClean.EjecutarQuitarAcentos
    Exit Sub
ErrHandler:
    MsgBox "No se pudo quitar acentos: " & Err.Description, vbExclamation, "Morfo"
End Sub

Public Sub CallbackQuitarRaros(control As IRibbonControl)
    On Error GoTo ErrHandler
    modDataClean.EjecutarQuitarRaros
    Exit Sub
ErrHandler:
    MsgBox "No se pudo quitar caracteres raros: " & Err.Description, vbExclamation, "Morfo"
End Sub

Public Sub CallbackLimpiezaTotal(control As IRibbonControl)
    On Error GoTo ErrHandler
    modDataClean.EjecutarLimpiezaTotal
    Exit Sub
ErrHandler:
    MsgBox "No se pudo realizar la limpieza total: " & Err.Description, vbExclamation, "Morfo"
End Sub

Public Sub CallbackFormularioAvanzado(control As IRibbonControl)
    On Error GoTo ErrHandler
    modUserFormLauncher.AbrirFormularioAvanzado
    Exit Sub
ErrHandler:
    MsgBox "No se pudo abrir el formulario avanzado: " & Err.Description, vbExclamation, "Morfo"
End Sub
