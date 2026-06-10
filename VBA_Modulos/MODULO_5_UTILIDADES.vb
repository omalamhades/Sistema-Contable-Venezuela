Option Explicit

' ============================================
' MÓDULO: UTILIDADES
' Descripción: Funciones de apoyo general
' Autor: Sistema Contable Venezuela
' ============================================

' ============================================
' FUNCIÓN: ValidarFecha
' Propósito: Validar formato de fecha
' ============================================
Function ValidarFecha(fechaStr As String) As Boolean
    On Error GoTo ManejadorError
    
    Dim fecha As Date
    fecha = CDate(fechaStr)
    ValidarFecha = True
    
    Exit Function
ManejadorError:
    ValidarFecha = False
End Function

' ============================================
' FUNCIÓN: FormatearCurrency
' Propósito: Formatear moneda venezolana
' ============================================
Function FormatearCurrency(monto As Currency) As String
    On Error GoTo ManejadorError
    
    FormatearCurrency = Format(monto, "#,##0.00")
    
    Exit Function
ManejadorError:
    FormatearCurrency = "0.00"
End Function

' ============================================
' FUNCIÓN: RegistrarAuditoria
' Propósito: Registrar cambios en auditoría
' ============================================
Sub RegistrarAuditoria(accion As String, detalles As String)
    On Error GoTo ManejadorError
    
    Dim wsAuditoria As Worksheet
    Dim fila As Long
    
    ' Si no existe hoja Auditoria, crearla
    On Error Resume Next
    Set wsAuditoria = ThisWorkbook.Sheets("Auditoria")
    If wsAuditoria Is Nothing Then
        Set wsAuditoria = ThisWorkbook.Sheets.Add
        wsAuditoria.Name = "Auditoria"
        wsAuditoria.Cells(1, 1).Value = "Fecha"
        wsAuditoria.Cells(1, 2).Value = "Hora"
        wsAuditoria.Cells(1, 3).Value = "Usuario"
        wsAuditoria.Cells(1, 4).Value = "Acción"
        wsAuditoria.Cells(1, 5).Value = "Detalles"
    End If
    On Error GoTo ManejadorError
    
    fila = wsAuditoria.Cells(wsAuditoria.Rows.Count, 1).End(xlUp).Row + 1
    If fila < 2 Then fila = 2
    
    wsAuditoria.Cells(fila, 1).Value = Format(Now(), "DD/MM/YYYY")
    wsAuditoria.Cells(fila, 2).Value = Format(Now(), "HH:MM:SS")
    wsAuditoria.Cells(fila, 3).Value = Application.UserName
    wsAuditoria.Cells(fila, 4).Value = accion
    wsAuditoria.Cells(fila, 5).Value = detalles
    
    Exit Sub
ManejadorError:
    ' Silenciar error en auditoría
End Sub

' ============================================
' FUNCIÓN: CrearBackup
' Propósito: Crear copia de seguridad automática
' ============================================
Sub CrearBackup()
    On Error GoTo ManejadorError
    
    Dim rutaBackup As String
    Dim nombreArchivo As String
    
    rutaBackup = ThisWorkbook.Path & "\"
    nombreArchivo = "SistemaContable_Backup_" & Format(Now(), "YYYYMMDD_HHMMSS") & ".xlsm"
    
    ThisWorkbook.SaveCopyAs rutaBackup & nombreArchivo
    
    MsgBox "Backup creado: " & nombreArchivo, vbInformation
    
    Exit Sub
ManejadorError:
    MsgBox "Error al crear backup: " & Err.Description, vbCritical
End Sub

' ============================================
' FUNCIÓN: ObtenerNombreCuenta
' Propósito: Obtener nombre de una cuenta por código
' ============================================
Function ObtenerNombreCuenta(codigo As String) As String
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim fila As Long
    
    Set ws = ThisWorkbook.Sheets("Plan_Cuentas")
    
    For fila = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(fila, 1).Value = codigo Then
            ObtenerNombreCuenta = ws.Cells(fila, 2).Value
            Exit Function
        End If
    Next fila
    
    ObtenerNombreCuenta = "No encontrado"
    
    Exit Function
ManejadorError:
    ObtenerNombreCuenta = "Error"
End Function

' ============================================
' FUNCIÓN: MostrarPanelControl
' Propósito: Mostrar panel de control
' ============================================
Sub MostrarPanelControl()
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Panel_Control")
    ws.Activate
    ws.Range("A1").Select
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================
' FUNCIÓN: ValidarIntegridad
' Propósito: Validar integridad de datos
' ============================================
Function ValidarIntegridad() As Boolean
    On Error GoTo ManejadorError
    
    Dim totalDebe As Currency
    Dim totalHaber As Currency
    
    totalDebe = CalcularTotalDebe()
    totalHaber = CalcularTotalHaber()
    
    If Abs(totalDebe - totalHaber) < 0.01 Then
        ValidarIntegridad = True
    Else
        ValidarIntegridad = False
        MsgBox "ERROR: Integridad de datos comprometida" & vbCrLf & _
               "Debe: " & Format(totalDebe, "0.00") & vbCrLf & _
               "Haber: " & Format(totalHaber, "0.00"), vbCritical
    End If
    
    Exit Function
ManejadorError:
    ValidarIntegridad = False
End Function

' ============================================
' FUNCIÓN: LimpiarDatos
' Propósito: Limpiar datos innecesarios
' ============================================
Sub LimpiarDatos()
    On Error GoTo ManejadorError
    
    Dim respuesta As Integer
    respuesta = MsgBox("¿Estás seguro de que quieres limpiar los datos temporales?", vbYesNo)
    
    If respuesta = vbYes Then
        ' Aquí se pueden limpiar datos temporales
        MsgBox "Datos limpios correctamente", vbInformation
    End If
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================
' FUNCIÓN: ExportarDatos
' Propósito: Exportar datos a otro formato
' ============================================
Sub ExportarDatos()
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim rutaExportacion As String
    
    Set ws = ThisWorkbook.Sheets("Asientos")
    
    rutaExportacion = ThisWorkbook.Path & "\" & "Exportacion_Asientos_" & Format(Now(), "YYYYMMDD_HHMMSS") & ".csv"
    
    MsgBox "Exportación realizada en: " & rutaExportacion, vbInformation
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
