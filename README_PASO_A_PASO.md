# Procesador Bancario - Actualización profesional por GitHub Releases

Esta carpeta ya trae tu app en:

```txt
app_src/app.py
```

La versión base preparada es:

```txt
v0.2.11
```

## 1. Qué archivos se suben a GitHub

Subí al repositorio estos archivos/carpetas:

```txt
app_src/
assets/
.github/
requirements.txt
launcher.py
build_app.bat
build_launcher.bat
crear_zip_release.bat
instalar_launcher_en_esta_pc.bat
VERSION.txt
.gitignore
README_PASO_A_PASO.md
```

No subas extractos bancarios, PDFs, Excel reales, CSV reales ni datos de clientes.

## 2. Configurar el launcher

Abrí `launcher.py` y cambiá:

```py
GITHUB_OWNER = "Nico2026-PY"
GITHUB_REPO = "xls-bank"
```

Ejemplo:

```py
GITHUB_OWNER = "Nico2026-PY"
GITHUB_REPO = "xls-bank"
```

## 3. Assets de tu app

Tu app busca el GIF en:

```txt
assets/logo_girando_desde_cero_sin_fondo.gif
```

Ponelo dentro de la carpeta `assets/`.

Opcional para el ícono del exe:

```txt
assets/icono_nuevo_app.ico
```

Si el GIF no está, tu app igual abre, pero muestra una advertencia de pantalla de carga.

## 4. Compilar la app principal

Ejecutá:

```bat
build_app.bat
```

Esto genera:

```txt
dist/Procesador_Bancario/Procesador_Bancario.exe
```

Se usa `--onedir`, que es lo más recomendable para apps con `pandas`, `openpyxl`, `PIL` y assets.

## 5. Crear el ZIP para GitHub Release

Ejecutá:

```bat
crear_zip_release.bat
```

Esto genera:

```txt
release/Procesador_Bancario_Windows.zip
```

Ese archivo es el que subís como asset de la Release.

## 6. Crear la Release en GitHub

En GitHub:

```txt
Repository → Releases → Draft a new release
```

Usá:

```txt
Tag: v0.2.11
Title: Procesador Bancario v0.2.11
Asset: Procesador_Bancario_Windows.zip
```

Publicás la Release.

## 7. Compilar el launcher

Ejecutá:

```bat
build_launcher.bat
```

Esto genera:

```txt
dist/Abrir_Procesador_Bancario.exe
```

Ese es el único `.exe` que instalás una vez en cada computadora.

## 8. Instalar launcher en una PC

En cada compu, copiás esta carpeta o solo el launcher compilado y ejecutás:

```bat
instalar_launcher_en_esta_pc.bat
```

Crea un acceso directo en el escritorio llamado:

```txt
Procesador Bancario
```

Cuando el usuario abre ese acceso directo:

```txt
1. Busca la última Release en GitHub.
2. Descarga Procesador_Bancario_Windows.zip.
3. Instala la versión en %LOCALAPPDATA%.
4. Abre Procesador_Bancario.exe.
```

## 9. Cómo actualizar en el futuro

Cada vez que hagas cambios:

```txt
1. Modificás app_src/app.py.
2. Cambiás VERSION.txt, por ejemplo a v0.2.12.
3. Ejecutás build_app.bat.
4. Ejecutás crear_zip_release.bat.
5. Creás una nueva Release en GitHub con tag v0.2.12.
6. Subís Procesador_Bancario_Windows.zip.
```

Las otras PCs se actualizan solas la próxima vez que abran el launcher.

## 10. Repo privado

Si el repositorio es privado, cada PC necesita un token en variable de entorno:

```bat
setx GITHUB_TOKEN "TU_TOKEN"
```

Después cerrá y abrí Windows o al menos cerrá la sesión para que tome la variable.

## 11. Dónde queda instalado en cada PC

La app queda acá:

```txt
%LOCALAPPDATA%\Procesador_Bancario\versions\v0.2.11\
```

El launcher guarda logs acá:

```txt
%LOCALAPPDATA%\Procesador_Bancario\launcher.log
```


---

## README principal actualizado

Además de este paso a paso, el proyecto ahora incluye:

```txt
README.md
README.txt
CHANGELOG.md
requirements.txt
assets/icono_nuevo_app.ico
assets/logo_girando_desde_cero_sin_fondo.gif
assets/logo_girando_3d_corregido.gif
```

El `README.md` documenta todas las mejoras acumuladas hasta `v0.2.11`, bancos soportados, formato esperado, reglas de seguridad de datos, assets, compilación y publicación por GitHub Releases.


---

## CONFIGURACIÓN PRIVADA DE EMPRESAS

Para publicar el repo sin exponer datos internos, las empresas reales ya no quedan escritas en `app.py`.

Crear este archivo en cada PC:

```txt
%APPDATA%\Procesador_Bancario\empresas_config.json
```

Usar este formato:

```json
{
  "empresas": {
    "EMPRESA_1": ["alias archivo", "razon social"],
    "EMPRESA_2": {"claves": ["otro alias"], "cuit": "opcional"}
  }
}
```

No subir `empresas_config.json` a GitHub. El repo solo debe tener `config/empresas_config.example.json`.
