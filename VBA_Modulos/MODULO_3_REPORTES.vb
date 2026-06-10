Option Explicit

' ============================================
' MÓDULO: REPORTES
' Descripción: Generación automática de reportes contables
' Autor: Sistema Contable Venezuela
' ============================================

' ============================================
' FUNCIÓN: GenerarDiario
' Propósito: Generar reporte diario de asientos
' ============================================
Sub GenerarDiario()
    On Error GoTo ManejadorError
    
    Dim wsAsientos As Worksheet
    Dim wsDiario As Worksheet
    Dim fila As Long
    Dim filaDestino As Long
    
    Set wsAsientos = ThisWorkbook.Sheets("Asientos")
    Set wsDiario = ThisWorkbook.Sheets("Diario")
    
    ' Limpiar reporte anterior
    wsDiario.Cells.Clear
    
    ' Escribir encabezados
    wsDiario.Cells(1, 1).Value = "DIARIO GENERAL"
    wsDiario.Cells(1, 1).Font.Bold = True
    wsDiario.Cells(1, 1).Font.Size = 14
    
    wsDiario.Cells(2, 1).Value = "Empresa: " & EmpresaActiva & " | Período: " & PeriodoActivo
    wsDiario.Cells(2, 1).Font.Bold = True
    wsDiario.Cells(3, 1).Value = "Fecha: " & Format(Now(), "DD/MM/YYYY")
    
    wsDiario.Cells(5, 1).Value = "Número"
    wsDiario.Cells(5, 2).Value = "Fecha"
    wsDiario.Cells(5, 3).Value = "Descripción"
    wsDiario.Cells(5, 4).Value = "Cuenta"
    wsDiario.Cells(5, 5).Value = "Debe"
    wsDiario.Cells(5, 6).Value = "Haber"
    
    filaDestino = 6
    
    ' Copiar asientos del período
    For fila = 2 To wsAsientos.Cells(wsAsientos.Rows.Count, 1).End(xlUp).Row
        If wsAsientos.Cells(fila, 8).Value = EmpresaActiva And _
           wsAsientos.Cells(fila, 9).Value = PeriodoActivo Then
            
            wsDiario.Cells(filaDestino, 1).Value = wsAsientos.Cells(fila, 1).Value
            wsDiario.Cells(filaDestino, 2).Value = wsAsientos.Cells(fila, 2).Value
            wsDiario.Cells(filaDestino, 3).Value = wsAsientos.Cells(fila, 3).Value
            wsDiario.Cells(filaDestino, 4).Value = wsAsientos.Cells(fila, 4).Value
            wsDiario.Cells(filaDestino, 5).Value = wsAsientos.Cells(fila, 5).Value
            
            filaDestino = filaDestino + 1
            
            wsDiario.Cells(filaDestino, 4).Value = wsAsientos.Cells(fila, 6).Value
            wsDiario.Cells(filaDestino, 6).Value = wsAsientos.Cells(fila, 7).Value
            
            filaDestino = filaDestino + 1
        End If
    Next fila
    
    ' Añadir totales
    filaDestino = filaDestino + 1
    wsDiario.Cells(filaDestino, 4).Value = "TOTALES:"
    wsDiario.Cells(filaDestino, 5).Value = CalcularTotalDebe()
    wsDiario.Cells(filaDestino, 6).Value = CalcularTotalHaber()
    
    MsgBox "Reporte Diario generado exitosamente", vbInformation
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================
' FUNCIÓN: GenerarMayor
' Propósito: Generar mayor analítico por cuenta
' ============================================
Sub GenerarMayor()
    On Error GoTo ManejadorError
    
    Dim wsAsientos As Worksheet
    Dim wsMayor As Worksheet
    Dim wsPlanCuentas As Worksheet
    Dim fila As Long
    Dim filaDestino As Long
    Dim codigoCuenta As String
    Dim filaPlan As Long
    Dim totalDebe As Currency
    Dim totalHaber As Currency
    
    Set wsAsientos = ThisWorkbook.Sheets("Asientos")
    Set wsMayor = ThisWorkbook.Sheets("Mayor")
    Set wsPlanCuentas = ThisWorkbook.Sheets("Plan_Cuentas")
    
    wsMayor.Cells.Clear
    
    wsMayor.Cells(1, 1).Value = "MAYOR ANALÍTICO"
    wsMayor.Cells(2, 1).Value = "Empresa: " & EmpresaActiva & " | Período: " & PeriodoActivo
    wsMayor.Cells(3, 1).Value = "Fecha: " & Format(Now(), "DD/MM/YYYY")
    
    wsMayor.Cells(5, 1).Value = "Código"
    wsMayor.Cells(5, 2).Value = "Nombre"
    wsMayor.Cells(5, 3).Value = "Debe"
    wsMayor.Cells(5, 4).Value = "Haber"
    wsMayor.Cells(5, 5).Value = "Saldo"
    
    filaDestino = 6
    
    ' Recorrer cada cuenta del plan
    For filaPlan = 2 To wsPlanCuentas.Cells(wsPlanCuentas.Rows.Count, 1).End(xlUp).Row
        codigoCuenta = wsPlanCuentas.Cells(filaPlan, 1).Value
        totalDebe = 0
        totalHaber = 0
        
        ' Calcular totales por cuenta
        For fila = 2 To wsAsientos.Cells(wsAsientos.Rows.Count, 1).End(xlUp).Row
            If wsAsientos.Cells(fila, 8).Value = EmpresaActiva And _
               wsAsientos.Cells(fila, 9).Value = PeriodoActivo Then
                
                If wsAsientos.Cells(fila, 4).Value = codigoCuenta Then
                    totalDebe = totalDebe + wsAsientos.Cells(fila, 5).Value
                End If
                
                If wsAsientos.Cells(fila, 6).Value = codigoCuenta Then
                    totalHaber = totalHaber + wsAsientos.Cells(fila, 7).Value
                End If
            End If
        Next fila
        
        ' Solo mostrar cuentas con movimiento
        If totalDebe > 0 Or totalHaber > 0 Then
            wsMayor.Cells(filaDestino, 1).Value = codigoCuenta
            wsMayor.Cells(filaDestino, 2).Value = wsPlanCuentas.Cells(filaPlan, 2).Value
            wsMayor.Cells(filaDestino, 3).Value = totalDebe
            wsMayor.Cells(filaDestino, 4).Value = totalHaber
            wsMayor.Cells(filaDestino, 5).Value = totalDebe - totalHaber
            
            filaDestino = filaDestino + 1
        End If
    Next filaPlan
    
    MsgBox "Reporte Mayor Analítico generado exitosamente", vbInformation
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ============================================
' FUNCIÓN: GenerarBalance
' Propósito: Generar balance de comprobación
' ============================================
Sub GenerarBalance()
    On Error GoTo ManejadorError
    
    Dim wsAsientos As Worksheet
    Dim wsBalance As Worksheet
    Dim wsPlanCuentas As Worksheet
    Dim fila As Long
    Dim filaDestino As Long
    Dim totalDebe As Currency
    Dim totalHaber As Currency
    Dim diferencia As Currency
    
    Set wsAsientos = ThisWorkbook.Sheets("Asientos")
    Set wsBalance = ThisWorkbook.Sheets("Balance")
    Set wsPlanCuentas = ThisWorkbook.Sheets("Plan_Cuentas")
    
    wsBalance.Cells.Clear
    
    wsBalance.Cells(1, 1).Value = "BALANCE DE COMPROBACIÓN"
    wsBalance.Cells(2, 1).Value = "Empresa: " & EmpresaActiva & " | Período: " & PeriodoActivo
    wsBalance.Cells(3, 1).Value = "Fecha: " & Format(Now(), "DD/MM/YYYY")
    
    wsBalance.Cells(5, 1).Value = "Cuenta"
    wsBalance.Cells(5, 2).Value = "Debe"
    wsBalance.Cells(5, 3).Value = "Haber"
    
    filaDestino = 6
    totalDebe = 0
    totalHaber = 0
    
    ' Calcular totales por cada cuenta
    For fila = 2 To wsPlanCuentas.Cells(wsPlanCuentas.Rows.Count, 1).End(xlUp).Row
        Dim sumaDebe As Currency
        Dim sumaHaber As Currency
        Dim codigoCuenta As String
        Dim filaAsiento As Long
        
        codigoCuenta = wsPlanCuentas.Cells(fila, 1).Value
        sumaDebe = 0
        sumaHaber = 0
        
        For filaAsiento = 2 To wsAsientos.Cells(wsAsientos.Rows.Count, 1).End(xlUp).Row
            If wsAsientos.Cells(filaAsiento, 8).Value = EmpresaActiva And _
               wsAsientos.Cells(filaAsiento, 9).Value = PeriodoActivo Then
                
                If wsAsientos.Cells(filaAsiento, 4).Value = codigoCuenta Then
                    sumaDebe = sumaDebe + wsAsientos.Cells(filaAsiento, 5).Value
                End If
                
                If wsAsientos.Cells(filaAsiento, 6).Value = codigoCuenta Then
                    sumaHaber = sumaHaber + wsAsientos.Cells(filaAsiento, 7).Value
                End If
            End If
        Next filaAsiento
        
        If sumaDebe > 0 Or sumaHaber > 0 Then
            wsBalance.Cells(filaDestino, 1).Value = codigoCuenta & " - " & wsPlanCuentas.Cells(fila, 2).Value
            wsBalance.Cells(filaDestino, 2).Value = sumaDebe
            wsBalance.Cells(filaDestino, 3).Value = sumaHaber
            
            totalDebe = totalDebe + sumaDebe
            totalHaber = totalHaber + sumaHaber
            
            filaDestino = filaDestino + 1
        End If
    Next fila
    
    ' Totales
    filaDestino = filaDestino + 1
    wsBalance.Cells(filaDestino, 1).Value = "TOTALES"
    wsBalance.Cells(filaDestino, 2).Value = totalDebe
    wsBalance.Cells(filaDestino, 3).Value = totalHaber
    
    ' Validar ecuación
    diferencia = Abs(totalDebe - totalHaber)
    If diferencia < 0.01 Then
        MsgBox "Balance de Comprobación CORRECTO" & vbCrLf & _
               "Total Debe = Total Haber = " & Format(totalDebe, "0.00"), vbInformation
    Else
        MsgBox "ERROR: Balance NO coincide" & vbCrLf & _
               "Debe: " & Format(totalDebe, "0.00") & vbCrLf & _
               "Haber: " & Format(totalHaber, "0.00"), vbExclamation
    End If
    
    Exit Sub
ManejadorError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
