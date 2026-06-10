# 📖 MANUAL DE USUARIO - SISTEMA CONTABLE VENEZUELA

## Tabla de Contenidos

1. [Instalación](#instalación)
2. [Primeros Pasos](#primeros-pasos)
3. [Configuración de Empresa](#configuración-de-empresa)
4. [Plan de Cuentas](#plan-de-cuentas)
5. [Registro de Asientos](#registro-de-asientos)
6. [Generación de Reportes](#generación-de-reportes)
7. [Gestión de Retenciones](#gestión-de-retenciones)
8. [Cierre de Período](#cierre-de-período)

---

## Instalación

### Requisitos
- Excel 2016 o superior
- Windows 10/11 o Mac OS
- Macros habilitadas en Excel

### Pasos

1. **Descargar** el archivo `SistemaContable_Venezuela.xlsm`

2. **Habilitar Macros en Excel:**
   - **Windows:** Archivo → Opciones → Centro de confianza → Configuración de macros → "Habilitar todas las macros"
   - **Mac:** Herramientas → Seguridad → "Permitir macros automatizadas"

3. **Abrir el archivo**
   - Si aparece aviso de seguridad, click en "Habilitar contenido"
   - Esperar a que cargue completamente

---

## Primeros Pasos

### Acceso al Sistema

1. Ir a la hoja **"Panel_Control"**
2. Esta es tu pantalla principal
3. Aquí encontrarás todos los botones principales

### Botones Principales

| Botón | Función |
|-------|----------|
| NUEVA EMPRESA | Crear empresa |
| NUEVO ASIENTO | Registrar asiento |
| GENERAR REPORTES | Crear reportes |
| RETENCIONES | Gestionar retenciones |
| CIERRE ANUAL | Cerrar período |

---

## Configuración de Empresa

### Crear Primera Empresa

1. Click en botón **"NUEVA EMPRESA"**
2. Se abrirá ventana de diálogo
3. Completar datos:
   - **Nombre:** Nombre legal de la empresa
   - **RUC:** Registro único de contribuyente
   - **Dirección:** Domicilio
   - **Teléfono:** Contacto

4. Click **"Aceptar"**
5. Empresa aparecerá en lista

### Cambiar Empresa Activa

1. En Panel_Control, seleccionar empresa de lista desplegable
2. Click **"Activar Empresa"**
3. Sistema confirmará cambio

### Cambiar Período Fiscal

1. Click en **"PERÍODO FISCAL"**
2. Seleccionar mes y año
3. Click **"Aceptar"**
4. Periodo se activará automáticamente

---

## Plan de Cuentas

### ¿Qué es?

Catálogo de todas las cuentas que usará la empresa para registrar movimientos contables.

### Estructura de Cuentas

**Jerarquía:**
```
1000 - 1999 → ACTIVO
  1000-1099 → Activo Circulante
    1000 → Bancos
    1100 → Cuentas por Cobrar
  1100-1199 → Activo Fijo
    1200 → Maquinaria
    1300 → Equipos

2000 - 2999 → PASIVO
  2000-2099 → Pasivo Circulante
    2000 → Proveedores
    2100 → Cuentas por Pagar
  2100-2199 → Pasivo Fijo
    2200 → Préstamos a LP

3000 - 3999 → PATRIMONIO
  3000 → Capital Social
  3100 → Reservas
  3200 → Resultados Acumulados

4000 - 4999 → INGRESOS
  4000 → Ventas
  4100 → Otros Ingresos

5000 - 5999 → GASTOS
  5000 → Compras
  5100 → Gastos Operacionales
```

### Crear Cuenta

1. Ir a hoja **"Plan_Cuentas"**
2. Click en primera fila vacía
3. Completar columnas:
   - **Código:** Número único (Ej: 1500)
   - **Nombre:** Descripción (Ej: Inventario)
   - **Tipo:** Categoría (Activo, Pasivo, Patrimonio, Ingresos, Gastos)
   - **Naturaleza:** D (Deudor) o C (Acreedor)
   - **Activo:** Sí/No

4. Presionar ENTER
5. Cuenta se guardará automáticamente

### Importar Plan de Cuentas

1. Click botón **"IMPORTAR CUENTAS"**
2. Seleccionar archivo Excel
3. Sistema importará automáticamente
4. Revisar que se hayan cargado todas

---

## Registro de Asientos

### ¿Qué es un Asiento?

Movimiento contable que registra un hecho económico. Debe cumplir:
**DEBE = HABER** (Ecuación Contable)

### Crear Asiento Manual

1. Click botón **"NUEVO ASIENTO"**
2. Se abrirá formulario
3. Completar datos:
   - **Fecha:** Día del movimiento (DD/MM/YYYY)
   - **Descripción:** Qué ocurrió (Ej: "Compra de inventario a proveedor XYZ")

4. Lado DEBE:
   - Seleccionar cuenta
   - Ingresar monto

5. Lado HABER:
   - Seleccionar cuenta
   - Ingresar monto

6. El sistema validará que DEBE = HABER

7. Click **"Registrar Asiento"**
8. Se confirmará número de asiento

### Ejemplo de Asiento

**Situación:** Compra mercadería a crédito por Bs. 500.000

| Lado | Cuenta | Monto |
|------|--------|-------|
| DEBE | 1500 Inventario | 500.000 |
| HABER | 2000 Proveedores | 500.000 |

### Asientos Pregrabados

Para asientos repetitivos:

1. Click **"ASIENTOS PREGRABADOS"**
2. Click **"NUEVO PREGRABADO"**
3. Crear asiento base
4. Dar nombre significativo (Ej: "Pago Nómina Mensual")
5. Guardar

Para usar después:
1. Click **"ASIENTOS PREGRABADOS"**
2. Seleccionar pregrabado
3. Modificar montos si es necesario
4. Click "Aplicar"

---

## Generación de Reportes

### ¿Qué Reportes Puedo Generar?

1. **Diario** - Listado cronológico de asientos
2. **Mayor** - Movimientos por cuenta
3. **Balance de Comprobación** - Validación DEBE=HABER
4. **Estado de Resultados** - Ingresos menos gastos
5. **Estado de Situación Financiera** - Activo, Pasivo, Patrimonio

### Generar Reporte

1. Click botón **"GENERAR REPORTES"**
2. Seleccionar tipo de reporte
3. Seleccionar período
4. Click **"Generar"**
5. Sistema crea automáticamente
6. Reporte aparece en hoja correspondiente
7. Listo para imprimir

### Imprimir Reporte

1. Abrir hoja del reporte (Ej: "Diario")
2. Presionar CTRL+P
3. Seleccionar impresora
4. Click **"Imprimir"**

---

## Gestión de Retenciones

### Retención de IVA

**¿Qué es?**
En Venezuela, cuando compras a proveedores que no son pequeños contribuyentes, debes retener el IVA y depositarlo.

**Registrar Retención:**

1. Ir a hoja **"Retencion_IVA"**
2. Click **"NUEVA RETENCIÓN"**
3. Completar:
   - Fecha
   - Número de factura
   - Monto base
   - Porcentaje de retención (16%)
4. Sistema calcula automáticamente
5. Click **"Guardar"**

**Generar Comprobante:**

1. Click **"GENERAR COMPROBANTE"**
2. Seleccionar mes
3. Sistema genera comprobante
4. Imprimir y guardar

**Exportar a SENIAT:**

1. Click **"EXPORTAR TXT SENIAT"**
2. Sistema genera archivo .txt
3. Descargar
4. Subir a portal SENIAT

### Retención de ISLR

**¿Qué es?**
Retención de Impuesto Sobre la Renta cuando eres prestador de servicios.

**Registrar Retención:**

1. Ir a hoja **"Retencion_ISLR"**
2. Click **"NUEVA RETENCIÓN ISLR"**
3. Completar:
   - Fecha
   - Concepto de ingreso
   - Monto
   - Porcentaje (3%, 4%, 6%, etc.)
4. Click **"Guardar"**

**Generar XML:**

1. Click **"GENERAR XML SENIAT"**
2. Sistema genera automáticamente
3. Descargar archivo .xml
4. Subir a portal SENIAT

---

## Cierre de Período

### ¿Qué es el Cierre?

Proceso anual que:
- Valida todos los asientos
- Genera reportes finales
- Cierra cuentas de ingresos/gastos
- Transfiere resultados a patrimonio
- Prepara nuevo período

### Realizar Cierre

1. Verificar que todos los asientos estén registrados
2. Click botón **"CIERRE ANUAL"**
3. Sistema mostrará resumen
4. Click **"Confirmar Cierre"**
5. Sistema generará:
   - Reporte final
   - Backup automático
   - Nuevo período listo

---

## Preguntas Frecuentes

### P: ¿Puedo editar un asiento?
**R:** No. Por seguridad contable, los asientos son de solo lectura. Para corregir, crear asiento de ajuste inverso.

### P: ¿Qué pasa si DEBE ≠ HABER?
**R:** El sistema rechaza el asiento. Revisar montos y cuentas.

### P: ¿Puedo tener varias empresas?
**R:** Sí, ilimitadas. Cambiar en Panel_Control.

### P: ¿Dónde se guardan los datos?
**R:** En el mismo archivo Excel. Hacer backup regular.

### P: ¿Puedo borrar datos?
**R:** Se recomienda NO borrar. Crear asientos de ajuste en su lugar.

### P: ¿Cómo exporto datos?
**R:** Click "Exportar" en Panel_Control. Genera Excel o PDF.

---

## Soporte y Ayuda

Si tienes problemas:

1. Revisar este manual
2. Revisar MANUAL_TECNICO_VBA.md
3. Verificar que macros estén habilitadas
4. Reiniciar Excel
5. Contactar soporte técnico

---

**¡Espero que disfrutes usando el Sistema Contable!** 🎉