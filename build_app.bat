@echo off
setlocal

echo =====================================================
echo Compilando Procesador Bancario
echo =====================================================

python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo Error instalando requirements.
    pause
    exit /b 1
)

echo.
echo Limpiando compilaciones anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist Procesador_Bancario.spec del /q Procesador_Bancario.spec

echo.
echo Compilando app principal...

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onedir ^
  --windowed ^
  --name "Procesador_Bancario" ^
  --icon "assets\icono_nuevo_app.ico" ^
  --add-data "assets\logo_girando_desde_cero_sin_fondo.gif;assets" ^
  --add-data "assets\logo_girando_3d_corregido.gif;assets" ^
  --add-data "VERSION.txt;." ^
  --hidden-import pandas ^
  --hidden-import openpyxl ^
  --hidden-import xlrd ^
  --hidden-import numpy ^
  --hidden-import PIL ^
  --hidden-import PIL.Image ^
  --hidden-import PIL.ImageTk ^
  "app_src\app.py"

if errorlevel 1 (
    echo.
    echo Error compilando la app.
    pause
    exit /b 1
)

echo.
echo =====================================================
echo App compilada correctamente.
echo Ubicacion:
echo dist\Procesador_Bancario\
echo =====================================================
pause