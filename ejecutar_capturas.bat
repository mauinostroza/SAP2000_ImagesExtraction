@echo off
REM ============================================================
REM ejecutar_capturas.bat
REM Ejecuta las capturas SAP2000 directamente desde el explorador
REM o desde una macro Excel via Shell command.
REM
REM USO DIRECTO:
REM   Doble clic en este archivo (con SAP2000 abierto)
REM
REM USO DESDE MACRO EXCEL (alternativa a xlwings RunPython):
REM   Shell "cmd /c cd /d """ & ThisWorkbook.Path & """ && ejecutar_capturas.bat"
REM ============================================================

setlocal

REM Directorio del script (mismo que este .bat)
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

REM Verificar que Python esté disponible
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no encontrado en el PATH.
    echo Instala Python 3.8+ y asegurate de que este en el PATH.
    pause
    exit /b 1
)

REM Verificar dependencias (instalacion rapida si faltan)
echo Verificando dependencias...
python -c "import comtypes, win32gui, PIL, pyautogui, xlwings, openpyxl" 2>nul
if %errorlevel% neq 0 (
    echo Instalando dependencias faltantes...
    pip install comtypes pywin32 Pillow pyautogui xlwings openpyxl --quiet
)

REM Ejecutar con el Excel de configuracion
set EXCEL_CONFIG=%SCRIPT_DIR%SAP2000_Capturas.xlsx

if not exist "%EXCEL_CONFIG%" (
    echo El archivo de configuracion no existe: %EXCEL_CONFIG%
    echo Creando plantilla...
    python crear_excel.py --ruta "%EXCEL_CONFIG%"
    echo Plantilla creada. Configurala y vuelve a ejecutar este script.
    start excel "%EXCEL_CONFIG%"
    pause
    exit /b 0
)

echo Ejecutando capturas SAP2000...
python sap_imagenes.py --config "%EXCEL_CONFIG%"

if %errorlevel% equ 0 (
    echo.
    echo Capturas completadas exitosamente.
) else (
    echo.
    echo Hubo errores. Revisa el log anterior.
)

pause
