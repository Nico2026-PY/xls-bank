@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Compilando launcher / actualizador profesional XlsBank
echo ============================================================

python --version >nul 2>&1
if errorlevel 1 (
  echo No se encontro Python. Instalalo y marca "Add Python to PATH".
  pause
  exit /b 1
)

echo.
echo Instalando/actualizando PyInstaller...
python -m pip install --upgrade pip
python -m pip install pyinstaller

if errorlevel 1 (
  echo.
  echo Error instalando PyInstaller.
  pause
  exit /b 1
)

echo.
echo Limpiando compilaciones anteriores del launcher...

if exist "build\Abrir_Procesador_Bancario" rmdir /s /q "build\Abrir_Procesador_Bancario"
if exist "dist\Abrir_Procesador_Bancario.exe" del /q "dist\Abrir_Procesador_Bancario.exe"
if exist "Abrir_Procesador_Bancario.spec" del /q "Abrir_Procesador_Bancario.spec"

echo.
echo Compilando launcher...

if exist "assets\icono_nuevo_app.ico" (
  python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "Abrir_Procesador_Bancario" ^
    --icon "assets\icono_nuevo_app.ico" ^
    "launcher.py"
) else (
  python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "Abrir_Procesador_Bancario" ^
    "launcher.py"
)

if errorlevel 1 (
  echo.
  echo Error compilando el launcher.
  pause
  exit /b 1
)

echo.
echo ============================================================
echo Launcher compilado correctamente.
echo Ubicacion:
echo dist\Abrir_Procesador_Bancario.exe
echo ============================================================
echo.
echo Ese .exe es el unico que instalas una vez en cada PC.
pause
