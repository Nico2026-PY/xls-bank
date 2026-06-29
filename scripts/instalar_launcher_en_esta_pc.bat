@echo off
setlocal
cd /d "%~dp0"

set DEST=%LOCALAPPDATA%\Procesador_Bancario_Launcher

if not exist "dist\Abrir_Procesador_Bancario.exe" (
  echo No existe dist\Abrir_Procesador_Bancario.exe
  echo Primero ejecuta build_launcher.bat.
  pause
  exit /b 1
)

if not exist "%DEST%" mkdir "%DEST%"
copy /Y "dist\Abrir_Procesador_Bancario.exe" "%DEST%\Abrir_Procesador_Bancario.exe" >nul

powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=(New-Object -COM WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop') + '\XlsBank.lnk'); $s.TargetPath='%DEST%\Abrir_Procesador_Bancario.exe'; $s.WorkingDirectory='%DEST%'; $s.IconLocation='%DEST%\Abrir_Procesador_Bancario.exe'; $s.Save()"

echo.
echo Launcher instalado para este usuario.
echo Acceso directo creado en el Escritorio: XlsBank
echo.
echo No requiere administrador.
pause
