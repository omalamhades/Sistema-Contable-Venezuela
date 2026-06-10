Option Explicit

' ============================================
' MÓDULO: ASIENTOS_CONTABLES
' Descripción: Gestión de asientos contables
' Autor: Sistema Contable Venezuela
' ============================================

' ============================================
' FUNCIÓN: RegistrarAsiento
' Propósito: Registrar un nuevo asiento contable
' ============================================
Sub RegistrarAsiento()
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim fila As Long
    Dim fecha As Date
    Dim descripcion As String
    Dim cuentaDebe As String
    Dim montoDebe As Currency
    Dim cuentaHaber As String
    Dim montoHaber As Currency
    
    Set ws = ThisWorkbook.Sheets("Asientos")
    
    ' Validar que exista período
    If PeriodoActivo = "" Then
        MsgBox "Debes configurar un período primero", vbExclamation
        Exit Sub
    End If
    
    ' Obtener datos del usuario
    On Error Resume Next
    fecha = CDate(InputBox("Ingresa la fecha (DD/MM/YYYY):", "Asiento Contable"))
    On Error GoTo ManejadorError
    If fecha = 0 Then Exit Sub
    
    descripcion = InputBox("Ingresa la descripción del asiento:", "Asiento Contable")
    If descripcion = "" Then Exit Sub
    
    cuentaDebe = InputBox("Ingresa código cuenta DEBE:", "Asiento Contable")
    If cuentaDebe = "" Then Exit Sub
    
    montoDebe = CDbl(InputBox("Ingresa monto DEBE:", "Asiento Contable"))
    If montoDebe = 0 Then Exit Sub
    
    cuentaHaber = InputBox("Ingresa código cuenta HABER:", "Asiento Contable")
    If cuentaHaber = "" Then Exit Sub
    
    montoHaber = CDbl(InputBox("Ingresa monto HABER:", "Asiento Contable"))
    If montoHaber = 0 Then Exit Sub
    
    ' Validar ecuación contable (DEBE = HABER)
    If montoDebe <> montoHaber Then
        MsgBox "ERROR: DEBE no es igual a HABER" & vbCrLf & _
               "Debe: " & Format(montoDebe, "0.00") & vbCrLf & _
               "Haber: " & Format(montoHaber, "0.00"), vbExclamation
        Exit Sub
    End If
    
    ' Validar que las cuentas existan en el plan de cuentas
    If Not ValidarCuenta(cuentaDebe) Then
        MsgBox "La cuenta DEBE no existe en el plan de cuentas", vbExclamation
        Exit Sub
    End If
    
    If Not ValidarCuenta(cuentaHaber) Then
        MsgBox "La cuenta HABER no existe en el plan de cuentas", vbExclamation
        Exit Sub
    End If
    
    ' Registrar el asiento
    fila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If fila < 2 Then fila = 2
    
    ws.Cells(fila, 1).Value = ObtenerProximoNumero()
    ws.Cells(fila, 2).Value = fecha
    ws.Cells(fila, 3).Value = descripcion
    ws.Cells(fila, 4).Value = cuentaDebe
    ws.Cells(fila, 5).Value = montoDebe
    ws.Cells(fila, 6).Value = cuentaHaber
    ws.Cells(fila, 7).Value = montoHaber
    ws.Cells(fila, 8).Value = EmpresaActiva
    ws.Cells(fila, 9).Value = PeriodoActivo
    ws.Cells(fila, 10).Value = Format(Now(), "DD/MM/YYYY HH:MM:SS")
    ws.Cells(fila, 11).Value = Application.UserName
    
    MsgBox "Asiento #" & ObtenerProximoNumero() - 1 & " registrado exitosamente", vbInformation
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================
' FUNCIÓN: ValidarCuenta
' Propósito: Verificar que una cuenta exista
' ============================================
Function ValidarCuenta(codigo As String) As Boolean
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim rango As Range
    
    Set ws = ThisWorkbook.Sheets("Plan_Cuentas")
    Set rango = ws.Columns(1).Find(What:=codigo, LookAt:=xlWhole)
    
    ValidarCuenta = Not (rango Is Nothing)
    
    Exit Function
ManejadorError:
    ValidarCuenta = False
End Function

' ============================================
' FUNCIÓN: CalcularTotalDebe
' Propósito: Calcular total del debe en asientos
' ============================================
Function CalcularTotalDebe() As Currency
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim fila As Long
    Dim total As Currency
    
    Set ws = ThisWorkbook.Sheets("Asientos")
    
    For fila = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(fila, 8).Value = EmpresaActiva And _
           ws.Cells(fila, 9).Value = PeriodoActivo Then
            total = total + ws.Cells(fila, 5).Value
        End If
    Next fila
    
    CalcularTotalDebe = total
    
    Exit Function
ManejadorError:
    CalcularTotalDebe = 0
End Function

' ============================================
' FUNCIÓN: CalcularTotalHaber
' Propósito: Calcular total del haber en asientos
' ============================================
Function CalcularTotalHaber() As Currency
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim fila As Long
    Dim total As Currency
    
    Set ws = ThisWorkbook.Sheets("Asientos")
    
    For fila = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(fila, 8).Value = EmpresaActiva And _
           ws.Cells(fila, 9).Value = PeriodoActivo Then
            total = total + ws.Cells(fila, 7).Value
        End If
    Next fila
    
    CalcularTotalHaber = total
    
    Exit Function
ManejadorError:
    CalcularTotalHaber = 0
End Function

' ============================================
' FUNCIÓN: ObtenerSaldoCuenta
' Propósito: Obtener saldo de una cuenta
' ============================================
Function ObtenerSaldoCuenta(codigo As String) As Currency
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim fila As Long
    Dim saldoDebe As Currency
    Dim saldoHaber As Currency
    Dim wsPlan As Worksheet
    Dim naturaleza As String
    
    Set ws = ThisWorkbook.Sheets("Asientos")
    Set wsPlan = ThisWorkbook.Sheets("Plan_Cuentas")
    
    ' Obtener naturaleza de la cuenta
    For fila = 2 To wsPlan.Cells(wsPlan.Rows.Count, 1).End(xlUp).Row
        If wsPlan.Cells(fila, 1).Value = codigo Then
            naturaleza = wsPlan.Cells(fila, 4).Value
            Exit For
        End If
    Next fila
    
    ' Calcular saldos
    For fila = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(fila, 8).Value = EmpresaActiva And _
           ws.Cells(fila, 9).Value = PeriodoActivo Then
            
            If ws.Cells(fila, 4).Value = codigo Then
                saldoDebe = saldoDebe + ws.Cells(fila, 5).Value
            End If
            
            If ws.Cells(fila, 6).Value = codigo Then
                saldoHaber = saldoHaber + ws.Cells(fila, 7).Value
            End If
        End If
    Next fila
    
    ' Si naturaleza es D (deudor), saldo = Debe - Haber
    If naturaleza = "D" Then
        ObtenerSaldoCuenta = saldoDebe - saldoHaber
    Else
        ObtenerSaldoCuenta = saldoHaber - saldoDebe
    End If
    
    Exit Function
ManejadorError:
    ObtenerSaldoCuenta = 0
End Function
