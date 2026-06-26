@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Compilando launcher / actualizador
 echo ============================================================

python --version >nul 2>&1
if errorlevel 1 (
  echo No se encontro Python. Instalalo y marcá "Add Python to PATH".
  pause
  exit /b 1
)

python -m pip install --upgrade pip
python -m pip install pyinstaller

set ICON_ARG=
if exist "assets\icono_nuevo_app.ico" (
  set ICON_ARG=--icon "assets\icono_nuevo_app.ico"
)

pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "Abrir_Procesador_Bancario" ^
  %ICON_ARG% ^
  launcher.py

if errorlevel 1 (
  echo.
  echo Error compilando el launcher.
  pause
  exit /b 1
)

echo.
echo Launcher listo en:
echo dist\Abrir_Procesador_Bancario.exe
echo.
echo Ese .exe es el unico que instalas una vez en cada PC.
pause
