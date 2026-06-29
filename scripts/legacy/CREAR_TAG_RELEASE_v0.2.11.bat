@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo Crear tag v0.2.11 y subirlo a GitHub
echo Esto dispara el workflow automatico si GitHub Actions esta habilitado.
echo ============================================================
echo.

git tag v0.2.11
git push origin v0.2.11

echo.
echo Despues revisa la pestana Actions y Releases del repo.
pause
