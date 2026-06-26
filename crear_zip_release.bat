@echo off
setlocal
cd /d "%~dp0"

set APP_DIR=dist\Procesador_Bancario
set ZIP_NAME=Procesador_Bancario_Windows.zip

if not exist "%APP_DIR%\Procesador_Bancario.exe" (
  echo Falta %APP_DIR%\Procesador_Bancario.exe
  echo Primero ejecutá build_app.bat.
  pause
  exit /b 1
)

if exist "%ZIP_NAME%" del /f /q "%ZIP_NAME%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path '%APP_DIR%\*' -DestinationPath '%ZIP_NAME%' -Force"

if errorlevel 1 (
  echo Error creando el zip.
  pause
  exit /b 1
)

if not exist "release" mkdir "release"
copy /Y "%ZIP_NAME%" "release\%ZIP_NAME%" >nul

echo.
echo Zip listo para subir a GitHub Releases:
echo %CD%\release\%ZIP_NAME%
echo.
echo Recordá crear la Release con tag: v0.2.11
pause
