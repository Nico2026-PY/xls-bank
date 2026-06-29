# XlsBank - Procesador Bancario

Versión: `v0.2.15`

Aplicación de escritorio para Windows hecha en Python/Tkinter. Procesa extractos bancarios en Excel y CSV, detecta banco, empresa y cuenta, genera un Excel consolidado y muestra un informe final del proceso.

XlsBank está pensado para centralizar movimientos bancarios de distintas empresas y bancos en un único archivo Excel ordenado, validado y fácil de revisar.

## Qué hace

- Permite seleccionar varios archivos bancarios Excel o CSV.
- Detecta bancos soportados por nombre de archivo o contenido.
- Detecta empresas desde una configuración privada externa, sin dejar nombres reales en el código.
- Genera un Excel final con `RESUMEN_GENERAL`, hojas por empresa y bloques por banco/cuenta.
- Genera hojas de totales por mes, banco y empresa.
- Muestra vista previa antes de generar el Excel.
- Informa archivos OK, con advertencias o con error.
- Omite archivos con error sin frenar todo el proceso.
- No modifica los archivos bancarios originales.
- Incluye launcher/actualizador mediante GitHub Releases.

## Bancos soportados

Procesamiento actual:

- BPN
- Galicia
- Banco Patagonia
- Mercado Pago Argentina

Bancos contemplados para detección por nombre, a completar con reglas específicas si el formato cambia:

- Nación / BNA
- Macro
- Santander
- BBVA / Francés
- ICBC

## Configuración privada de empresas

La app carga las empresas desde un archivo local privado llamado:

```txt
empresas_config.json
```

Ubicación recomendada en cada PC:

```txt
%APPDATA%\Procesador_Bancario\empresas_config.json
```

También se puede indicar una ruta específica con variable de entorno:

```bat
setx BANCOS_EMPRESAS_CONFIG "C:\RutaPrivada\empresas_config.json"
```

Formato recomendado:

```json
{
  "empresas": {
    "EMPRESA_1": {
      "claves": ["empresa 1", "razon social ejemplo", "alias archivo"],
      "cuit": "opcional-no-subir-cuit-real-al-repo"
    },
    "EMPRESA_2": ["empresa 2", "otro alias"]
  }
}
```

La versión actual usa `nombre` y `claves` para detectar empresa. El campo `cuit` queda disponible para futuras reglas privadas, pero no es necesario para procesar movimientos.

Si no existe `empresas_config.json`, la app sigue funcionando y usa como fallback la primera palabra del nombre del archivo, mostrando una advertencia.

## Estructura del proyecto

```txt
.github/workflows/                      Workflows de GitHub Actions
app_src/app.py                          App principal
assets/                                 Íconos, logos y splash
config/empresas_config.example.json     Ejemplo público sin datos reales
docs/                                   Documentación secundaria
scripts/                                Scripts útiles de instalación
scripts/legacy/                         Scripts históricos
launcher.py                             Actualizador por GitHub Releases

README.md                               Documentación principal
CHANGELOG.md                            Historial de cambios
LICENSE.txt                             Licencia propietaria
VERSION.txt                             Versión actual
requirements.txt                        Dependencias Python

build_app.bat                           Compila la app principal
build_launcher.bat                      Compila el launcher
crear_zip_release.bat                   Crea el ZIP para GitHub Release
```

## Requisitos

```txt
pandas
openpyxl
xlrd
numpy
pillow
pyinstaller
```

Instalar dependencias:

```bat
python -m pip install -r requirements.txt
```

## Ejecutar en desarrollo

```bat
python app_src\app.py
```

## Compilar app principal

```bat
build_app.bat
```

Salida esperada:

```txt
dist\Procesador_Bancario\Procesador_Bancario.exe
```

## Crear ZIP para GitHub Release

```bat
crear_zip_release.bat
```

Salida esperada:

```txt
release\Procesador_Bancario_Windows.zip
```

Ese archivo se sube como asset de la Release en GitHub.

## Actualización automática

Cada PC instala una sola vez el launcher:

```txt
Abrir_Procesador_Bancario.exe
```

El launcher revisa GitHub Releases, descarga la última versión de:

```txt
Procesador_Bancario_Windows.zip
```

actualiza la carpeta local y abre la app.

Por compatibilidad, se mantienen estos nombres:

```txt
Procesador_Bancario.exe
Procesador_Bancario_Windows.zip
Abrir_Procesador_Bancario.exe
```

Aunque el branding visible de la app sea XlsBank.

Antes de compilar el launcher, revisar en `launcher.py`:

```py
GITHUB_OWNER = "Nico2026-PY"
GITHUB_REPO = "xls-bank"
```

## Scripts útiles

```txt
scripts/instalar_launcher_en_esta_pc.bat
scripts/instalar_config_privada_en_esta_pc.bat
```

Scripts históricos:

```txt
scripts/legacy/
```

## Publicar una nueva versión

Flujo recomendado:

1. Trabajar los cambios en una rama `dev`.
2. Probar la app en desarrollo.
3. Actualizar `VERSION.txt`, `CHANGELOG.md` y `APP_VERSION`.
4. Pasar cambios a `main`.
5. Ejecutar:

```bat
build_app.bat
crear_zip_release.bat
```

6. Crear una nueva Release en GitHub con el tag correspondiente.
7. Subir:

```txt
release\Procesador_Bancario_Windows.zip
```

8. Las PCs se actualizan solas al abrir el launcher.

## Seguridad de datos

El repositorio público no debe contener:

- Nombres reales de empresas.
- CUITs reales.
- Extractos bancarios reales.
- PDFs, Excel o CSV de clientes.
- Saldos, movimientos reales o reportes generados.
- Tokens de GitHub o claves privadas.
- Archivos `empresas_config.json` reales.

El archivo privado de empresas está ignorado por `.gitignore` y no debe subirse al repo.

## Nota importante si el repo ya estuvo público

Si antes se subieron datos reales al repo, borrarlos del archivo actual no siempre alcanza, porque pueden quedar en el historial de Git.

En ese caso conviene crear un repositorio limpio o reescribir historial antes de seguir usando el repo público.

## Historial reciente

Ver historial completo en:

```txt
CHANGELOG.md
```

Resumen:

- `v0.2.15`: limpieza del repositorio y orden de documentación/scripts.
- `v0.2.14`: preparación comercial, licencia propietaria y branding XlsBank.
- `v0.2.13`: soporte para CSV nuevo de Mercado Pago.
- `v0.2.12`: mejoras de launcher, seguridad y distribución.
- `v0.2.11`: base estable con informe final condicional.

## Licencia

XlsBank es software propietario.

Copyright © 2026 Nicolás Mellado. Todos los derechos reservados.

El uso, copia, modificación, redistribución, publicación o comercialización del software, código fuente, ejecutables, launcher, scripts, documentación, assets o cualquier archivo asociado requiere autorización expresa del autor.

Este repositorio no otorga permiso para reutilizar el código en otros proyectos, revender el software, crear versiones derivadas ni distribuir copias a terceros.

Ver condiciones completas en [`LICENSE.txt`](LICENSE.txt).