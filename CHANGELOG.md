# Changelog - Procesador Bancario
## v0.2.13

- Se agrega soporte para CSV nuevo de Mercado Pago.
- Se mejora la descripción de movimientos de Mercado Pago.
- Se corrige la columna Cuenta para Mercado Pago.
- Se mejora el ancho de columnas del Excel generado.

## v0.2.12
- Pantalla visual de actualización.
- Mensajes de estado al buscar, descargar, instalar y abrir la app.
- Modo offline: si no hay internet, intenta abrir la última versión instalada.
- Evita descargar si la versión local ya está actualizada.
- Soporte para canal estable / beta.
- Backup antes de actualizar.
- Validación básica del ZIP descargado.
- Logs locales para diagnosticar errores.
- Opción de reinstalar la última versión.

## v0.2.11
- Informe final condicional.
- Colores de estado mejorados.
- Vista clara de OK, advertencias y errores.
- Copiar informe al portapapeles.
- Abrir Excel generado desde informe final.
- Archivos con error se omiten y quedan informados.

## Cambios acumulados recientes
- Correcciones de Patagonia para cuenta en .xls.
- Fallback binario para cuenta Patagonia.
- Salida conservadora: no inventar categoria, etiquetas ni saldos.
- Mercado Pago Argentina agregado.
- Ordenamiento por fecha real.
- Totales por mes, banco y empresa.
- Panel lateral de navegación por empresa.
- Carpeta de salida y nombre de Excel mejorados.
- Preparación para GitHub Releases y launcher actualizador.


## Seguridad para repo público
- Empresas reales removidas del código fuente.
- Nueva carga de empresas desde `empresas_config.json` privado.
- Agregado `config/empresas_config.example.json` sin datos reales.
- `.gitignore` actualizado para evitar subir configuración privada.
