# XlsBank - Guía paso a paso para publicar actualizaciones

Esta guía explica cómo trabajar, compilar y publicar nuevas versiones de XlsBank usando GitHub Releases y el launcher actualizador.

La app principal está en:

```txt
app_src/app.py
```

La versión actual del proyecto se define en:

```txt
VERSION.txt
```

También debe coincidir con la constante `APP_VERSION` dentro de `app_src/app.py`.

---

## 1. Estructura principal del proyecto

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

No subir extractos bancarios, PDFs, Excel reales, CSV reales, reportes generados, datos de clientes, CUITs, tokens ni configuraciones privadas reales.

---

## 2. Archivos importantes

```txt
app_src/app.py                          Código principal de XlsBank
launcher.py                             Launcher / actualizador
requirements.txt                        Dependencias Python
VERSION.txt                             Versión actual
CHANGELOG.md                            Historial de versiones
README.md                               Documentación principal
LICENSE.txt                             Licencia propietaria
build_app.bat                           Compila la app principal
build_launcher.bat                      Compila el launcher
crear_zip_release.bat                   Crea el ZIP para GitHub Release
scripts/instalar_launcher_en_esta_pc.bat
scripts/instalar_config_privada_en_esta_pc.bat
config/empresas_config.example.json
assets/
```

---

## 3. Configurar el launcher

Abrir `launcher.py` y revisar:

```py
GITHUB_OWNER = "Nico2026-PY"
GITHUB_REPO = "xls-bank"
```

El launcher busca la última release publicada en GitHub y descarga el archivo:

```txt
Procesador_Bancario_Windows.zip
```

Por compatibilidad, se mantienen estos nombres internos:

```txt
Procesador_Bancario.exe
Procesador_Bancario_Windows.zip
Abrir_Procesador_Bancario.exe
```

Aunque el branding visible de la app sea XlsBank.

---

## 4. Assets de la app

La app usa estos archivos:

```txt
assets/logo_girando_desde_cero_sin_fondo.gif
assets/icono_nuevo_app.ico
```

El GIF se usa en la pantalla de carga.

El `.ico` se usa como ícono de ventana y ejecutable.

No borrar esos archivos salvo que también se actualicen las rutas correspondientes en `app_src/app.py` y/o scripts de compilación.

---

## 5. Ejecutar en desarrollo

Instalar dependencias:

```bat
python -m pip install -r requirements.txt
```

Ejecutar la app:

```bat
python app_src\app.py
```

Antes de publicar una versión, probar que:

```txt
- La app abre correctamente.
- Muestra el nombre XlsBank.
- Muestra la versión correcta.
- Agrega archivos Excel y CSV.
- Procesa los bancos soportados.
- Genera el Excel final.
- Abre la ventana “Acerca de XlsBank”.
```

---

## 6. Flujo de ramas recomendado

La rama `main` debe quedar como versión estable publicada.

Las mejoras se trabajan en ramas `dev`.

Ejemplo:

```bat
git checkout main
git pull
git checkout -b dev-v0.2.15-limpieza-repo
git push -u origin dev-v0.2.15-limpieza-repo
```

Mientras la versión está en desarrollo, puede figurar como:

```txt
v0.2.15-dev
```

Cuando se prepara para release estable, se cambia a:

```txt
v0.2.15
```

---

## 7. Preparar una versión estable

Antes de pasar una rama `dev` a `main`, revisar:

```txt
VERSION.txt
app_src/app.py       APP_VERSION
CHANGELOG.md
README.md
```

Ejemplo:

```txt
v0.2.15-dev   → durante desarrollo
v0.2.15       → antes de publicar release estable
```

Después guardar:

```bat
git status
git add .
git commit -m "Prepara version estable v0.2.15"
git push
```

---

## 8. Pasar cambios a main

Cuando la rama dev esté probada:

```bat
git checkout main
git pull
git merge dev-v0.2.15-limpieza-repo
git push
```

Desde ese momento, `main` queda actualizado.

---

## 9. Compilar la app principal

Desde `main`, ejecutar:

```bat
build_app.bat
```

Salida esperada:

```txt
dist\Procesador_Bancario\Procesador_Bancario.exe
```

Probar el `.exe` antes de crear el ZIP.

---

## 10. Crear el ZIP para GitHub Release

Ejecutar:

```bat
crear_zip_release.bat
```

Salida esperada:

```txt
release\Procesador_Bancario_Windows.zip
```

Ese archivo es el asset que se sube a GitHub Release.

No cambiar el nombre del ZIP salvo que también se actualice el launcher.

---

## 11. Crear la Release en GitHub

En GitHub:

```txt
Repository → Releases → Draft a new release
```

Usar:

```txt
Tag: v0.2.15
Title: XlsBank - Procesador Bancario v0.2.15
Asset: Procesador_Bancario_Windows.zip
```

Publicar la release como `Latest`.

---

## 12. Compilar el launcher

El launcher no necesita recompilarse en cada versión de la app.

Solo se recompila si se cambia `launcher.py`.

Para compilarlo:

```bat
build_launcher.bat
```

Salida esperada:

```txt
dist\Abrir_Procesador_Bancario.exe
```

Ese es el ejecutable que se instala una vez en cada PC.

---

## 13. Instalar launcher en una PC

El script útil está en:

```txt
scripts\instalar_launcher_en_esta_pc.bat
```

El usuario final debería usar:

```txt
Abrir_Procesador_Bancario.exe
```

Cuando se abre el launcher:

```txt
1. Busca la última release en GitHub.
2. Descarga Procesador_Bancario_Windows.zip.
3. Instala la versión en %LOCALAPPDATA%.
4. Abre Procesador_Bancario.exe.
```

---

## 14. Dónde queda instalado en cada PC

La app queda en:

```txt
%LOCALAPPDATA%\Procesador_Bancario\
```

Las versiones quedan dentro de:

```txt
%LOCALAPPDATA%\Procesador_Bancario\versions\
```

El launcher guarda logs en:

```txt
%LOCALAPPDATA%\Procesador_Bancario\launcher.log
```

---

## 15. Configuración privada de empresas

Para publicar el repo sin exponer datos internos, las empresas reales no quedan escritas en `app.py`.

Crear este archivo en cada PC:

```txt
%APPDATA%\Procesador_Bancario\empresas_config.json
```

También se puede usar variable de entorno:

```bat
setx BANCOS_EMPRESAS_CONFIG "C:\RutaPrivada\empresas_config.json"
```

Formato recomendado:

```json
{
  "empresas": {
    "EMPRESA_1": {
      "claves": ["alias archivo", "razon social"],
      "cuit": "opcional"
    },
    "EMPRESA_2": ["otro alias", "otra razon social"]
  }
}
```

No subir `empresas_config.json` a GitHub.

El repo solo debe tener:

```txt
config/empresas_config.example.json
```

---

## 16. Repo privado

Si en el futuro el repositorio pasa a privado, cada PC necesitaría autorización para descargar releases privadas.

Una opción sería usar token en variable de entorno:

```bat
setx GITHUB_TOKEN "TU_TOKEN"
```

Pero para uso comercial conviene revisar una solución más segura antes de entregar a clientes.

---

## 17. Seguridad

No subir al repositorio:

```txt
- Extractos bancarios reales.
- Excel reales.
- CSV reales.
- PDFs reales.
- CUITs reales.
- Nombres reales de empresas/clientes.
- Saldos o movimientos reales.
- Tokens o claves privadas.
- empresas_config.json real.
- ZIPs de release.
- Carpetas build/, dist/ o release/.
```

---

## 18. Licencia

XlsBank es software propietario.

Ver condiciones completas en:

```txt
LICENSE.txt
```