@echo off
setlocal

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"
set VENV_DIR=%SCRIPT_DIR%.build-venv

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no encontrado en PATH.
    exit /b 1
)

if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo Creando entorno virtual de build...
    python -m venv "%VENV_DIR%"
    if %errorlevel% neq 0 exit /b 1
)

echo Instalando dependencias...
"%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip >nul
"%VENV_DIR%\Scripts\python.exe" -m pip install -r requirements-exe.txt
if %errorlevel% neq 0 exit /b 1

echo Limpiando builds previos...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo Generando ejecutable unico (onefile, GUI)...
"%VENV_DIR%\Scripts\python.exe" -m PyInstaller --clean sap2000_portable.spec
if %errorlevel% neq 0 exit /b 1

echo.
echo ============================================
echo BUILD COMPLETADO
echo ============================================
echo Salida: %SCRIPT_DIR%dist\sap2000_capture.exe
echo.
echo El .exe es auto-contenido. Solo necesitas
echo distribuirlo junto al Excel de configuracion.
echo ============================================
exit /b 0
