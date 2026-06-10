Option Explicit

' ============================================
' MÓDULO: CONFIGURACION
' Descripción: Funciones para configurar empresas y períodos
' Autor: Sistema Contable Venezuela
' ============================================

' Variable global para almacenar empresa activa
Public EmpresaActiva As String
Public PeriodoActivo As String

' ============================================
' FUNCIÓN: CrearEmpresa
' Propósito: Crear una nueva empresa en el sistema
' ============================================
Sub CrearEmpresa()
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim fila As Long
    Dim ruc As String
    Dim nombre As String
    Dim telefono As String
    Dim direccion As String
    
    ' Asignar la hoja de Empresas
    Set ws = ThisWorkbook.Sheets("Empresas")
    
    ' Obtener entrada del usuario
    nombre = InputBox("Ingresa el nombre de la empresa:", "Nueva Empresa")
    If nombre = "" Then Exit Sub
    
    ruc = InputBox("Ingresa el RUC de la empresa:", "RUC")
    If ruc = "" Then Exit Sub
    
    telefono = InputBox("Ingresa el teléfono:", "Teléfono")
    direccion = InputBox("Ingresa la dirección:", "Dirección")
    
    ' Encontrar primera fila vacía
    fila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If fila <= 1 Then fila = 2
    
    ' Registrar empresa
    ws.Cells(fila, 1).Value = nombre
    ws.Cells(fila, 2).Value = ruc
    ws.Cells(fila, 3).Value = telefono
    ws.Cells(fila, 4).Value = direccion
    ws.Cells(fila, 5).Value = Format(Now(), "DD/MM/YYYY")
    ws.Cells(fila, 6).Value = "Activa"
    
    ' Confirmar
    MsgBox "Empresa '" & nombre & "' creada exitosamente", vbInformation
    EmpresaActiva = nombre
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================
' FUNCIÓN: ConfigurarPeriodo
' Propósito: Establecer período fiscal activo
' ============================================
Sub ConfigurarPeriodo()
    On Error GoTo ManejadorError
    
    Dim mes As String
    Dim ano As String
    
    mes = InputBox("Ingresa el mes (01-12):", "Período")
    If mes = "" Or Not IsNumeric(mes) Then Exit Sub
    
    If CLng(mes) < 1 Or CLng(mes) > 12 Then
        MsgBox "Mes inválido. Debe ser 01-12", vbExclamation
        Exit Sub
    End If
    
    ano = InputBox("Ingresa el año (YYYY):", "Período")
    If ano = "" Or Not IsNumeric(ano) Then Exit Sub
    
    PeriodoActivo = Format(CLng(mes), "00") & "/" & ano
    
    MsgBox "Período configurado: " & PeriodoActivo, vbInformation
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================
' FUNCIÓN: ObtenerProximoNumero
' Propósito: Obtener siguiente número de asiento
' ============================================
Function ObtenerProximoNumero() As Long
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim ultimoNumero As Long
    
    Set ws = ThisWorkbook.Sheets("Asientos")
    
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaFila < 2 Then
        ObtenerProximoNumero = 1
    Else
        If IsNumeric(ws.Cells(ultimaFila, 1).Value) Then
            ultimoNumero = CLng(ws.Cells(ultimaFila, 1).Value)
            ObtenerProximoNumero = ultimoNumero + 1
        Else
            ObtenerProximoNumero = 1
        End If
    End If
    
    Exit Function
ManejadorError:
    ObtenerProximoNumero = 1
End Function

' ============================================
' FUNCIÓN: IsNumeric
' Propósito: Validar si un valor es numérico
' ============================================
Function IsNumeric(ByVal sText As String) As Boolean
    On Error Resume Next
    IsNumeric = Not IsNull(CDbl(sText))
    On Error GoTo 0
End Function
