# 🔧 MANUAL TÉCNICO - CÓDIGO VBA

## Para Programadores: Documentación Completa del Código

### Tabla de Contenidos

1. [Estructura General](#estructura-general)
2. [Módulos VBA](#módulos-vba)
3. [Funciones Principales](#funciones-principales)
4. [Variables Globales](#variables-globales)
5. [Manejo de Errores](#manejo-de-errores)

---

## Estructura General

### Organización de Módulos

```
ThisWorkbook
├── Módulo 1: CONFIGURACION
├── Módulo 2: ASIENTOS_CONTABLES
├── Módulo 3: REPORTES
├── Módulo 4: RETENCIONES
├── Módulo 5: LIBROS_AUXILIARES
└── Módulo 6: UTILIDADES
```

### Convenciones de Código

- **Variables:** camelCase (empresaActiva, periodoActivo)
- **Funciones:** PascalCase (CrearEmpresa, GenerarDiario)
- **Constantes:** UPPER_CASE (FILA_INICIO, COLUMNA_DATOS)
- **Comentarios:** En español, claros y concisos

---

## Módulos VBA

### 1. MÓDULO: CONFIGURACION

**Propósito:** Gestionar empresas, períodos y configuración general.

### 2. MÓDULO: ASIENTOS_CONTABLES

**Propósito:** Gestionar registro y validación de asientos.

### 3. MÓDULO: REPORTES

**Propósito:** Generar reportes contables automáticamente.

### 4. MÓDULO: RETENCIONES

**Propósito:** Gestionar retenciones IVA e ISLR.

### 5. MÓDULO: LIBROS_AUXILIARES

**Propósito:** Gestionar libros de compras y ventas.

### 6. MÓDULO: UTILIDADES

**Propósito:** Funciones de apoyo general.

---

## Variables Globales

```vba
' Configuración Actual
Public EmpresaActiva As String      ' Empresa seleccionada
Public PeriodoActivo As String      ' Período en formato MM/YYYY
Public UsuarioActivo As String      ' Usuario actual

' Constantes de Rango
Const FILA_INICIO As Long = 2       ' Primera fila de datos
Const COLUMNA_EMPRESA As Long = 8   ' Columna empresa en asientos
Const COLUMNA_PERIODO As Long = 9   ' Columna período en asientos

' Constantes de Retención
Const TASA_IVA As Double = 0.16     ' 16% IVA
Const TASA_ISLR_BASE As Double = 0.03  ' 3% ISLR base
```

---

## Manejo de Errores

### Patrón General

```vba
Sub MiSubrutina()
    On Error GoTo ManejadorError
    
    ' ... código aquí ...
    
    Exit Sub
ManejadorError:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    ' Log del error
End Sub
```

---

**Documentación actualizada: 10/06/2026**