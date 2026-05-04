# SAP2000 Image Capture

Automatiza capturas de SAP2000 a partir de un archivo `SAP2000_Capturas.xlsx`, con ejecución por CLI o por EXE portable.

El flujo vigente usa el mecanismo **Capture Picture BMP** de SAP2000 como base de exportación y se controla desde el archivo Excel de configuración. No requiere macros, `xlsm` ni `xlwings` para el uso normal por `--config`.

## Archivos relevantes

```text
README.md
ejecutar_capturas.bat
sap_imagenes.py
crear_excel.py
build_exe.bat
sap2000_portable.py
sap2000_portable.spec
requirements-exe.txt
```

## Requisitos

| Requisito | Versión |
| --- | --- |
| Windows | 10 o 11 |
| SAP2000 | v23 |
| Python | 3.8+ (solo si ejecutas el `.py`) |

Dependencias Python:

```bat
python -m pip install comtypes pywin32 Pillow pyautogui openpyxl
```

## Flujo rápido

1. Crea la plantilla Excel.
2. Completa `CONFIG` y `CAPTURAS`.
3. Abre SAP2000 con el modelo cargado.
4. Ejecuta el script, el EXE o `ejecutar_capturas.bat`.

## Crear la plantilla Excel

Con Python:

```bat
python crear_excel.py
```

O con el EXE:

```bat
sap2000_capture.exe --crear-excel SAP2000_Capturas.xlsx
```

Esto genera `SAP2000_Capturas.xlsx` en la ruta indicada.

## Ejecutar capturas

Con Python:

```bat
python sap_imagenes.py --config SAP2000_Capturas.xlsx
```

Con EXE portable:

```bat
sap2000_capture.exe --config SAP2000_Capturas.xlsx
```

Si haces doble clic sobre `sap2000_capture.exe` y existe `SAP2000_Capturas.xlsx`
en la misma carpeta del ejecutable, el EXE usará esa configuración automáticamente.

Si trabajas dentro del repo tras compilar:

```bat
dist\sap2000_capture\sap2000_capture.exe --config SAP2000_Capturas.xlsx
```

Launcher `.bat`:

```bat
ejecutar_capturas.bat
```

El launcher busca `SAP2000_Capturas.xlsx` en la misma carpeta. Si no existe, intenta crearlo con el EXE y, si no está disponible, con Python. El artifact portable del workflow ahora debe incluir también este `.bat`.

## Estructura del Excel

### Hoja `CONFIG`

| Celda | Descripción |
| --- | --- |
| `B2` | Ruta del DLL de SAP2000 |
| `B3` | Nombre del proyecto |
| `B4` | Subcarpeta de salida |

La salida por defecto se guarda en `Capturas_SAP` como subcarpeta del Excel.

### Hoja `CAPTURAS`

Cada fila activa define una imagen a generar.

| Columna | Valor |
| --- | --- |
| `ACTIVO` | `SI` / `NO` |
| `NOMBRE IMAGEN` | Nombre base de la captura |
| `TIPO VISTA` | `PLANTA`, `ELEV_X`, `ELEV_Y`, `ISO_NE`, `ISO_NO`, `ISO_SE`, `ISO_SO`, `CUSTOM` |
| `AZIMUT` | Solo para `CUSTOM` |
| `ELEVACIÓN` | Solo para `CUSTOM` |
| `MODO DISPLAY` | `MODELO` o `CARGAS` |
| `PATRÓN DE CARGA` | Requerido si `MODO DISPLAY = CARGAS` |
| `TIPO VENTANA` | `COMPLETA` o `PARCIAL` |
| `RECORTE IZQ/SUP/DER/INF %` | Solo para `PARCIAL` |

Ejemplo de salida:

```text
Capturas_SAP\
  MiProyecto_Vista_General_ISO_NE_MODELO.png
  MiProyecto_Cargas_Muertas_ISO_NE_CARGAS_DEAD.png
```

## Opciones CLI útiles

Crear Excel y salir:

```bat
sap2000_capture.exe --crear-excel C:\ruta\SAP2000_Capturas.xlsx
```

Permitir una carpeta de salida absoluta o fuera de la carpeta base del Excel:

```bat
sap2000_capture.exe --config SAP2000_Capturas.xlsx --allow-unsafe-output
```

También puedes usar la variable de entorno:

```bat
set SAP2000_ALLOW_UNSAFE_OUTPUT=1
```

## EXE portable

Build local:

```bat
build_exe.bat
```

Salida:

```text
dist\sap2000_capture\
```

El EXE usa la misma CLI que `sap_imagenes.py`.

## Solución de problemas

| Problema | Causa probable | Acción |
| --- | --- | --- |
| SAP2000 no detectado | SAP2000 no está abierto o visible | Abre el modelo y deja la ventana visible |
| No se genera el `.xlsx` | No hay EXE ni Python disponible | Compila el EXE o instala Python |
| Error de dependencias | Faltan paquetes de Python | Instala los paquetes indicados arriba |
| Fila ignorada | Configuración inválida en `CAPTURAS` | Revisa vista, display, patrón y recortes |
| No escribe fuera de la carpeta del Excel | Modo seguro activo | Usa `--allow-unsafe-output` si realmente lo necesitas |

## Alcance actual

- Archivo de configuración: `xlsx`
- Ejecución soportada: `Python CLI`, `EXE portable`, `ejecutar_capturas.bat`
- Referencias obsoletas eliminadas: macros VBA, flujo `xlsm`, dependencia de `xlwings` para el uso normal
