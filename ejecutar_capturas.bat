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
set GUI_EXE_PATH_LOCAL=%SCRIPT_DIR%sap2000_capture_gui.exe
set GUI_EXE_PATH_DIST=%EXE_DIR%\sap2000_capture_gui.exe
set EXCEL_CONFIG=%SCRIPT_DIR%SAP2000_Capturas.xlsx
set EXE_PATH=
set GUI_EXE_PATH=
set GUI_MODE=
set UNSAFE_OUTPUT_FLAG=
set PASSTHROUGH_ARGS=
set CONFIG_OVERRIDE_IGNORED=

:parse_args
if "%~1"=="" goto args_done
if /I "%~1"=="--gui" (
    set GUI_MODE=1
) else if /I "%~1"=="--config" (
    set CONFIG_OVERRIDE_IGNORED=1
    shift
    if "%~1"=="" (
        echo ERROR: --config requiere una ruta, pero este launcher siempre usa:
        echo   %EXCEL_CONFIG%
        pause
        exit /b 2
    )
) else if /I "%~1"=="--allow-unsafe-output" (
    set UNSAFE_OUTPUT_FLAG=--allow-unsafe-output
) else (
    set PASSTHROUGH_ARGS=%PASSTHROUGH_ARGS% "%~1"
)
shift
goto parse_args
:args_done

if exist "%EXE_PATH_LOCAL%" set EXE_PATH=%EXE_PATH_LOCAL%
if not defined EXE_PATH if exist "%EXE_PATH_DIST%" set EXE_PATH=%EXE_PATH_DIST%
if exist "%GUI_EXE_PATH_LOCAL%" set GUI_EXE_PATH=%GUI_EXE_PATH_LOCAL%
if not defined GUI_EXE_PATH if exist "%GUI_EXE_PATH_DIST%" set GUI_EXE_PATH=%GUI_EXE_PATH_DIST%

echo Directorio del launcher: %SCRIPT_DIR%
if defined EXE_PATH (
    echo EXE detectado: %EXE_PATH%
) else (
    echo EXE detectado: ninguno
)
if defined GUI_EXE_PATH (
    echo EXE GUI detectado: %GUI_EXE_PATH%
) else (
    echo EXE GUI detectado: ninguno
)
echo Excel configuracion: %EXCEL_CONFIG%
if defined CONFIG_OVERRIDE_IGNORED (
    echo Aviso: `--config` se ignora en este launcher. Se usara el Excel detectado arriba.
)

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
    if defined GUI_MODE (
        if defined GUI_EXE_PATH (
            echo Abriendo interfaz grafica SAP2000...
            "%GUI_EXE_PATH%" --config "%EXCEL_CONFIG%" %UNSAFE_OUTPUT_FLAG% %PASSTHROUGH_ARGS%
        ) else (
            echo Abriendo interfaz grafica SAP2000 con EXE CLI...
            "%EXE_PATH%" --gui --config "%EXCEL_CONFIG%" %UNSAFE_OUTPUT_FLAG% %PASSTHROUGH_ARGS%
        )
    ) else (
        echo Ejecutando capturas SAP2000 con EXE...
        "%EXE_PATH%" --config "%EXCEL_CONFIG%" %UNSAFE_OUTPUT_FLAG% %PASSTHROUGH_ARGS%
    )
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

    if defined GUI_MODE (
        echo Abriendo interfaz grafica SAP2000 con Python...
        python sap2000_gui.py --config "%EXCEL_CONFIG%" %UNSAFE_OUTPUT_FLAG% %PASSTHROUGH_ARGS%
    ) else (
        echo Ejecutando capturas SAP2000 con Python...
        python sap_imagenes.py --config "%EXCEL_CONFIG%" %UNSAFE_OUTPUT_FLAG% %PASSTHROUGH_ARGS%
    )
)

if %errorlevel% equ 0 (
    echo.
    echo Proceso completado.
) else (
    echo.
    echo Hubo errores. Codigo devuelto: %errorlevel%.
    echo Revisa el log anterior.
)

pause
