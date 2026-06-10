# 📊 SISTEMA CONTABLE COMPLETO - VENEZUELA

## 🎯 Descripción

Sistema de contabilidad integrado en Excel con VBA diseñado para empresas venezolanas. Cumple con NIF, Código de Comercio y regulaciones SENIAT.

## ✨ Características Principales

✅ **Plan de Cuentas Dinámico** - Adaptable y escalable
✅ **Asientos Contables** - Manuales y pregrabados
✅ **Reportes Automáticos** - Diario, Mayor, Balance, Estado Resultados
✅ **Retenciones** - IVA e ISLR con exportación SENIAT
✅ **Libros Auxiliares** - Compras y Ventas
✅ **Multiempresas** - Gestión ilimitada de empresas
✅ **Validaciones** - Ecuación contable automática
✅ **Auditoría** - Registro de cambios y usuario

## 🚀 Inicio Rápido

1. Descargar `SistemaContable_Venezuela.xlsm` (en sección Releases)
2. Habilitar macros al abrir
3. Crear empresa en Panel de Control
4. Agregar plan de cuentas
5. Registrar asientos
6. Generar reportes

## 📋 Contenido del Repositorio

### Documentación
- `README.md` - Este archivo
- `MANUAL_USUARIO.md` - Guía completa de uso
- `MANUAL_TECNICO_VBA.md` - Documentación de código
- `INSTALACION_VBA.txt` - Pasos de instalación

### Hojas de Excel Incluidas
- Panel_Control - Centro de operaciones
- Plan_Cuentas - Catálogo de cuentas
- Empresas - Gestión de empresas
- Asientos - Registro de movimientos
- Libro_Compras - Registro de compras
- Libro_Ventas - Registro de ventas
- Retencion_IVA - Control de retenciones IVA
- Retencion_ISLR - Control de retenciones ISLR
- Diario - Reporte diario generado
- Mayor - Mayor analítico generado
- Balance - Balance de comprobación
- Estado_Resultados - P&L generado automático
- Situacion_Financiera - Balance general generado

## 🔧 Módulos VBA Incluidos

### 1. CONFIGURACION
- CrearEmpresa()
- ConfigurarPeriodo()
- ObtenerProximoNumero()

### 2. ASIENTOS_CONTABLES
- RegistrarAsiento()
- ValidarCuenta()
- CalcularTotalDebe()
- CalcularTotalHaber()

### 3. REPORTES
- GenerarDiario()
- GenerarMayor()
- GenerarBalance()
- GenerarEstadoResultados()
- GenerarSituacionFinanciera()

### 4. RETENCIONES
- CalcularRetencionIVA()
- CalcularRetencionISLR()
- GenerarTXTSENIAT()
- GenerarXMLISLR()

### 5. LIBROS_AUXILIARES
- RegistrarCompra()
- RegistrarVenta()
- GenerarLibroCompras()
- GenerarLibroVentas()

### 6. UTILIDADES
- ValidarFecha()
- FormatearCurrency()
- CrearBackup()
- RegistrarAuditoria()

## 💡 Ejemplos de Uso

### Crear Asiento
```
Fecha: 10/06/2026
Descripción: Compra de inventario
Cuenta Debe: 1500 (Inventario) - 500.000,00
Cuenta Haber: 2000 (Proveedores) - 500.000,00
```

### Generar Reporte
1. Click botón "Generar Reportes"
2. Seleccionar tipo (Diario, Mayor, Balance)
3. Sistema genera automáticamente
4. Listo para imprimir

## 📦 Requisitos de Sistema

- Windows 10/11 o Mac OS
- Excel 2016 o superior
- Macros habilitadas
- 50 MB espacio disponible

## 🔐 Características de Seguridad

✅ Protección de hojas con contraseña
✅ Validación de integridad de datos
✅ Registro de auditoría automático
✅ Backup automático de datos
✅ Control de acceso por usuario

## 📞 Soporte

Para consultas:
- Revisar MANUAL_USUARIO.md
- Consultar MANUAL_TECNICO_VBA.md
- Ver ejemplos en carpeta EJEMPLOS/
- Crear un Issue en el repositorio

## ⚖️ Cumplimiento Legal - Venezuela

✅ Cumple con NIF (Normas de Información Financiera)
✅ Cumple con Código de Comercio
✅ Compatible con SENIAT
✅ Genera reportes según regulaciones vigentes
✅ Exporta TXT y XML para presentación de retenciones

## 📄 Licencia

MIT License - Uso libre para empresas venezolanas

## 👨‍💻 Autor

Creado por: GitHub Copilot
Última actualización: 10/06/2026
Versiones: v1.0 - Completo y funcional

---

**¡Listo para usar!** Descarga el archivo y comienza a hacer pruebas. 🚀