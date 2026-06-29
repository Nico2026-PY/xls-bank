# Procesador Bancario por Empresa

Versión: `v0.2.11`

Aplicación de escritorio para Windows hecha en Python/Tkinter. Procesa extractos bancarios en Excel, detecta banco/empresa/cuenta, genera un Excel consolidado y muestra un informe final del proceso.

## Qué hace

- Permite seleccionar varios archivos Excel bancarios.
- Detecta bancos soportados por nombre de archivo o contenido.
- Detecta empresas desde una configuración privada externa, sin dejar nombres reales en el código.
- Genera un Excel final con `RESUMEN_GENERAL`, hojas por empresa, bloques por banco/cuenta y hojas de control.
- Muestra vista previa antes de generar.
- Informa archivos OK, con advertencias o con error.
- Omite archivos con error sin frenar todo el proceso.
- No modifica los archivos bancarios originales.

## Bancos soportados

Procesamiento completo actual:

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

## Archivos importantes

```txt
app_src/app.py                         App principal
launcher.py                            Actualizador por GitHub Releases
requirements.txt                       Dependencias Python
VERSION.txt                            Versión actual
build_app.bat                          Compila la app principal
build_launcher.bat                     Compila el launcher
crear_zip_release.bat                  Crea el ZIP para GitHub Release
instalar_launcher_en_esta_pc.bat       Instala el launcher en una PC
instalar_config_privada_en_esta_pc.bat Instala config privada local
config/empresas_config.example.json    Ejemplo público sin datos reales
assets/                                Ícono y logo/splash
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

Cada PC instala una sola vez:

```txt
Abrir_Procesador_Bancario.exe
```

Ese launcher revisa GitHub Releases, descarga la última versión de `Procesador_Bancario_Windows.zip`, actualiza la carpeta local y abre la app.

Antes de compilar el launcher, editar en `launcher.py`:

```py
GITHUB_OWNER = "Nico2026-PY"
GITHUB_REPO = "xls-bank"
```

## Publicar una nueva versión

1. Cambiar `VERSION.txt`, por ejemplo de `v0.2.11` a `v0.2.12`.
2. Modificar `app_src/app.py`.
3. Ejecutar `build_app.bat`.
4. Ejecutar `crear_zip_release.bat`.
5. Crear una nueva Release en GitHub con el mismo tag.
6. Subir `release/Procesador_Bancario_Windows.zip`.
7. Las PCs se actualizan solas al abrir el launcher.

## Nota importante si el repo ya estuvo público

Si antes se subieron datos reales al repo, borrarlos del archivo actual no siempre alcanza, porque pueden quedar en el historial de Git. En ese caso conviene crear un repositorio limpio o reescribir historial antes de seguir usando el repo público.

## Historial v0.2.11

- Informe final condicional.
- Colores por estado en informe.
- Botón para copiar informe.
- Archivos con error omitidos e informados.
- Reglas conservadoras para no inventar categoría, etiquetas ni saldos.
- Carga privada de empresas desde `empresas_config.json` para repositorios públicos.

## Licencia

XlsBank es software propietario.

Copyright © 2026 Nicolás Mellado. Todos los derechos reservados.

El uso, copia, modificación, redistribución, publicación o comercialización
del software, código fuente, ejecutables, launcher, scripts, documentación,
assets o cualquier archivo asociado requiere autorización expresa del autor.

Este repositorio no otorga permiso para reutilizar el código en otros proyectos,
revender el software, crear versiones derivadas ni distribuir copias a terceros.

Ver condiciones completas en [`LICENSE.txt`](LICENSE.txt).