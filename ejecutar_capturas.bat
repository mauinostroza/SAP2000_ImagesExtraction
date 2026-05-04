@echo off
REM ============================================================
REM ejecutar_capturas.bat
REM Launcher local para SAP2000 Image Capture
REM Usa SAP2000_Capturas.xlsx y ejecuta el backend CLI disponible
REM
REM USO DIRECTO:
REM   Doble clic en este archivo (con SAP2000 abierto)
REM ============================================================

setlocal

REM Directorio del script (mismo que este .bat)
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

set EXE_PATH_LOCAL=%SCRIPT_DIR%sap2000_capture.exe
set EXE_DIR=%SCRIPT_DIR%dist\sap2000_capture
set EXE_PATH_DIST=%EXE_DIR%\sap2000_capture.exe
set EXE_PATH=

if exist "%EXE_PATH_LOCAL%" set EXE_PATH=%EXE_PATH_LOCAL%
if not defined EXE_PATH if exist "%EXE_PATH_DIST%" set EXE_PATH=%EXE_PATH_DIST%

REM Configuracion principal basada en XLSX
set EXCEL_CONFIG=%SCRIPT_DIR%SAP2000_Capturas.xlsx

if not exist "%EXCEL_CONFIG%" (
    echo No existe el archivo de configuracion: %EXCEL_CONFIG%
    echo Creando plantilla...
    if defined EXE_PATH (
        "%EXE_PATH%" --crear-excel "%EXCEL_CONFIG%"
    ) else (
        python --version >nul 2>&1
        if %errorlevel% neq 0 (
            echo ERROR: no se encontro ni el EXE portable ni Python en PATH.
            pause
            exit /b 1
        )
        python crear_excel.py --ruta "%EXCEL_CONFIG%"
    )
    echo Plantilla creada: %EXCEL_CONFIG%
    echo Completa la hoja CONFIG y la hoja CAPTURAS, guarda el .xlsx y vuelve a ejecutar este launcher.
    pause
    exit /b 0
)

if defined EXE_PATH (
    echo Ejecutando capturas SAP2000 con EXE...
    "%EXE_PATH%" --config "%EXCEL_CONFIG%" %*
) else (
    python --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo ERROR: no se encontro ni el EXE portable ni Python en PATH.
        echo Genera el EXE con build_exe.bat o instala Python 3.8+ y las dependencias del proyecto.
        pause
        exit /b 1
    )

    echo Verificando dependencias Python...
    python -c "import comtypes, win32gui, PIL, pyautogui, openpyxl" 2>nul
    if %errorlevel% neq 0 (
        echo ERROR: faltan dependencias Python requeridas.
        echo Instala manualmente:
        echo   python -m pip install -r requirements-exe.txt
        echo o usa el EXE portable generado con build_exe.bat.
        pause
        exit /b 1
    )

    echo Ejecutando capturas SAP2000 con Python...
    python sap_imagenes.py --config "%EXCEL_CONFIG%" %*
)

if %errorlevel% equ 0 (
    echo.
    echo Proceso completado.
) else (
    echo.
    echo Hubo errores. Revisa el log anterior.
)

pause
