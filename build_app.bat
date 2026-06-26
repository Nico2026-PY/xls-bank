@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Compilando Procesador Bancario
 echo ============================================================

python --version >nul 2>&1
if errorlevel 1 (
  echo No se encontro Python. Instalalo y marcá "Add Python to PATH".
  pause
  exit /b 1
)

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

set ADD_ASSETS=
if exist "assets" (
  set ADD_ASSETS=--add-data "assets;assets"
)

set ICON_ARG=
if exist "assets\icono_nuevo_app.ico" (
  set ICON_ARG=--icon "assets\icono_nuevo_app.ico"
)

pyinstaller --noconfirm --clean --onedir --windowed ^
  --name "Procesador_Bancario" ^
  %ICON_ARG% ^
  %ADD_ASSETS% ^
  "app_src\app.py"

if errorlevel 1 (
  echo.
  echo Error compilando la app.
  pause
  exit /b 1
)

echo.
echo App compilada en:
echo dist\Procesador_Bancario\Procesador_Bancario.exe
echo.
pause
