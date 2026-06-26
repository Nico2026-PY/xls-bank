@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Subir proyecto a GitHub por primera vez
echo Repo: https://github.com/Nico2026-PY/xls-bank.git
echo ============================================================
echo.

where git >nul 2>&1
if errorlevel 1 (
  echo No se encontro Git instalado.
  echo Instalalo desde https://git-scm.com/download/win y volve a ejecutar este BAT.
  pause
  exit /b 1
)

if not exist ".git" (
  git init
)

git branch -M main
git remote remove origin >nul 2>&1
git remote add origin https://github.com/Nico2026-PY/xls-bank.git

git add .
git commit -m "Version inicial segura v0.2.11" || echo Puede que no haya cambios para commitear.
git push -u origin main

echo.
echo Listo. Si GitHub pide usuario/token, segui el login que te muestre Git.
pause
