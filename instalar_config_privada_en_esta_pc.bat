@echo off
setlocal
cd /d "%~dp0"

set DEST=%APPDATA%\Procesador_Bancario
if not exist "%DEST%" mkdir "%DEST%"

if exist "empresas_config.json" (
  copy /Y "empresas_config.json" "%DEST%\empresas_config.json" >nul
  echo Configuracion privada instalada en:
  echo %DEST%\empresas_config.json
) else (
  if exist "config\empresas_config.example.json" (
    copy /Y "config\empresas_config.example.json" "%DEST%\empresas_config.json" >nul
    echo Se copio un ejemplo en:
    echo %DEST%\empresas_config.json
    echo.
    echo Editalo con tus empresas reales antes de usar la app.
  ) else (
    echo No se encontro empresas_config.json ni config\empresas_config.example.json
  )
)

echo.
pause
