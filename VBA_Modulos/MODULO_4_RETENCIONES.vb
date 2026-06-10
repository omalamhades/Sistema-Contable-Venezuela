Option Explicit

' ============================================
' MÓDULO: RETENCIONES
' Descripción: Gestión de retenciones IVA e ISLR
' Autor: Sistema Contable Venezuela
' ============================================

' Constantes de retención
Const TASA_IVA As Double = 0.16
Const TASA_ISLR_MINIMA As Double = 0.03

' ============================================
' FUNCIÓN: CalcularRetencionIVA
' Propósito: Calcular retención de IVA
' ============================================
Function CalcularRetencionIVA(monto As Currency) As Currency
    On Error GoTo ManejadorError
    
    CalcularRetencionIVA = monto * TASA_IVA
    
    Exit Function
ManejadorError:
    CalcularRetencionIVA = 0
End Function

' ============================================
' FUNCIÓN: CalcularRetencionISLR
' Propósito: Calcular retención de ISLR
' ============================================
Function CalcularRetencionISLR(monto As Currency, porcentaje As Double) As Currency
    On Error GoTo ManejadorError
    
    If porcentaje < 0 Or porcentaje > 1 Then
        MsgBox "Porcentaje inválido", vbExclamation
        Exit Function
    End If
    
    CalcularRetencionISLR = monto * porcentaje
    
    Exit Function
ManejadorError:
    CalcularRetencionISLR = 0
End Function

' ============================================
' FUNCIÓN: RegistrarRetencionIVA
' Propósito: Registrar una retención de IVA
' ============================================
Sub RegistrarRetencionIVA()
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim fila As Long
    Dim fecha As Date
    Dim numeroFactura As String
    Dim montoBase As Currency
    Dim montoRetenido As Currency
    
    Set ws = ThisWorkbook.Sheets("Retencion_IVA")
    
    fecha = CDate(InputBox("Ingresa la fecha (DD/MM/YYYY):", "Retención IVA"))
    numeroFactura = InputBox("Ingresa número de factura:", "Retención IVA")
    montoBase = CDbl(InputBox("Ingresa monto base:", "Retención IVA"))
    
    montoRetenido = CalcularRetencionIVA(montoBase)
    
    fila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If fila < 2 Then fila = 2
    
    ws.Cells(fila, 1).Value = fecha
    ws.Cells(fila, 2).Value = numeroFactura
    ws.Cells(fila, 3).Value = montoBase
    ws.Cells(fila, 4).Value = TASA_IVA
    ws.Cells(fila, 5).Value = montoRetenido
    ws.Cells(fila, 6).Value = EmpresaActiva
    ws.Cells(fila, 7).Value = PeriodoActivo
    
    MsgBox "Retención IVA registrada: " & Format(montoRetenido, "0.00"), vbInformation
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================
' FUNCIÓN: RegistrarRetencionISLR
' Propósito: Registrar una retención de ISLR
' ============================================
Sub RegistrarRetencionISLR()
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim fila As Long
    Dim fecha As Date
    Dim concepto As String
    Dim montoBase As Currency
    Dim porcentaje As Double
    Dim montoRetenido As Currency
    
    Set ws = ThisWorkbook.Sheets("Retencion_ISLR")
    
    fecha = CDate(InputBox("Ingresa la fecha (DD/MM/YYYY):", "Retención ISLR"))
    concepto = InputBox("Ingresa concepto de ingreso:", "Retención ISLR")
    montoBase = CDbl(InputBox("Ingresa monto base:", "Retención ISLR"))
    porcentaje = CDbl(InputBox("Ingresa porcentaje (3, 4, 6, 8, 10):", "Retención ISLR")) / 100
    
    montoRetenido = CalcularRetencionISLR(montoBase, porcentaje)
    
    fila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If fila < 2 Then fila = 2
    
    ws.Cells(fila, 1).Value = fecha
    ws.Cells(fila, 2).Value = concepto
    ws.Cells(fila, 3).Value = montoBase
    ws.Cells(fila, 4).Value = porcentaje
    ws.Cells(fila, 5).Value = montoRetenido
    ws.Cells(fila, 6).Value = EmpresaActiva
    ws.Cells(fila, 7).Value = PeriodoActivo
    
    MsgBox "Retención ISLR registrada: " & Format(montoRetenido, "0.00"), vbInformation
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================
' FUNCIÓN: GenerarTotalRetencionIVA
' Propósito: Calcular total de retenciones IVA del mes
' ============================================
Function GenerarTotalRetencionIVA(mes As Integer, ano As Integer) As Currency
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim fila As Long
    Dim total As Currency
    Dim periodoFiltro As String
    
    Set ws = ThisWorkbook.Sheets("Retencion_IVA")
    periodoFiltro = Format(mes, "00") & "/" & ano
    
    For fila = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(fila, 7).Value = periodoFiltro And _
           ws.Cells(fila, 6).Value = EmpresaActiva Then
            total = total + ws.Cells(fila, 5).Value
        End If
    Next fila
    
    GenerarTotalRetencionIVA = total
    
    Exit Function
ManejadorError:
    GenerarTotalRetencionIVA = 0
End Function

' ============================================
' FUNCIÓN: GenerarTotalRetencionISLR
' Propósito: Calcular total de retenciones ISLR del mes
' ============================================
Function GenerarTotalRetencionISLR(mes As Integer, ano As Integer) As Currency
    On Error GoTo ManejadorError
    
    Dim ws As Worksheet
    Dim fila As Long
    Dim total As Currency
    Dim periodoFiltro As String
    
    Set ws = ThisWorkbook.Sheets("Retencion_ISLR")
    periodoFiltro = Format(mes, "00") & "/" & ano
    
    For fila = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(fila, 7).Value = periodoFiltro And _
           ws.Cells(fila, 6).Value = EmpresaActiva Then
            total = total + ws.Cells(fila, 5).Value
        End If
    Next fila
    
    GenerarTotalRetencionISLR = total
    
    Exit Function
ManejadorError:
    GenerarTotalRetencionISLR = 0
End Function
