# Changelog - XlsBank

## v0.2.15

### Limpieza del repositorio
- Se elimina `README.txt` desactualizado.
- Se mueve documentación secundaria a `docs/`.
- Se crea carpeta `scripts/` para instaladores y utilidades.
- Se crea carpeta `scripts/legacy/` para scripts históricos.
- Se eliminan artefactos locales de build de la raíz del proyecto.
- Se refuerza `.gitignore` para evitar subir ZIPs, builds, archivos Excel, CSV, PDF, logs y configuración privada.

## v0.2.14

### Preparación comercial
- Se agrega `LICENSE.txt` con licencia propietaria para preparar XlsBank como producto comercial.
- Se documenta que el uso, copia, modificación, redistribución y comercialización requieren autorización expresa del autor.
- Se agrega sección de licencia en `README.md`.
- Se agrega pantalla “Acerca de XlsBank” dentro de la app.
- Se agrega texto de copyright y software propietario dentro de la documentación y la app.
- Se actualiza el branding visible de la app a XlsBank.

### Seguridad y propiedad intelectual
- Se refuerza la separación entre código fuente, ejecutables, releases y datos privados locales.
- Se aclara que archivos reales de bancos, empresas, CUITs, PDFs, Excels, CSVs, tokens o configuraciones privadas no deben subirse al repositorio.
- Se mantiene el modelo de configuración privada mediante `empresas_config.json`.

### Compatibilidad
- No se cambia el nombre del ejecutable principal.
- No se cambia el nombre del ZIP de release.
- No se modifica el launcher.
- Se mantiene compatibilidad con `Procesador_Bancario.exe` y `Procesador_Bancario_Windows.zip`.

## v0.2.13

### Mercado Pago
- Se agrega soporte para CSV nuevo de Mercado Pago.
- Se mejora la descripción de movimientos de Mercado Pago.
- Se corrige la columna Cuenta para Mercado Pago.
- Se mejora el ancho de columnas del Excel generado.

### Correcciones
- Se corrige la ruta de assets para evitar bloqueo de la pantalla de carga en modo desarrollo.

## v0.2.12

### Launcher y actualización automática
- Pantalla visual de actualización.
- Mensajes de estado al buscar, descargar, instalar y abrir la app.
- Modo offline: si no hay internet, intenta abrir la última versión instalada.
- Evita descargar si la versión local ya está actualizada.
- Soporte para canal estable / beta.
- Backup antes de actualizar.
- Validación básica del ZIP descargado.
- Logs locales para diagnosticar errores.
- Opción de reinstalar la última versión.

### Seguridad y distribución
- Preparación para GitHub Releases y launcher actualizador.
- Refuerzo de `.gitignore`.
- Separación de configuración privada mediante `empresas_config.json`.
- Configuración privada fuera del repositorio público.

## v0.2.11

### Base estable
- Informe final condicional.
- Colores de estado mejorados.
- Vista clara de OK, advertencias y errores.
- Copiar informe al portapapeles.
- Abrir Excel generado desde informe final.
- Archivos con error se omiten y quedan informados.

## Cambios acumulados recientes

- Correcciones de Patagonia para cuenta en `.xls`.
- Fallback binario para cuenta Patagonia.
- Salida conservadora: no inventar categoría, etiquetas ni saldos.
- Mercado Pago Argentina agregado.
- Ordenamiento por fecha real.
- Totales por mes, banco y empresa.
- Panel lateral de navegación por empresa.
- Carpeta de salida y nombre de Excel mejorados.

## Seguridad para repo público

- Empresas reales removidas del código fuente.
- Nueva carga de empresas desde `empresas_config.json` privado.
- Agregado `config/empresas_config.example.json` sin datos reales.
- `.gitignore` actualizado para evitar subir configuración privada.