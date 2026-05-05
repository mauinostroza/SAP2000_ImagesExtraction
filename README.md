# SAP2000 Image Capture

Automatiza capturas de SAP2000 a partir de un archivo `SAP2000_Capturas.xlsx`.

El flujo usa el mecanismo **Capture Picture BMP** de SAP2000 como base de exportación y se controla desde el archivo Excel de configuración. No requiere macros, `xlsm` ni `xlwings`.

## Archivos del proyecto

```
sap2000_capture.exe     → Ejecutable único (GUI) — generado con build_exe.bat
sap2000_gui.py          → Interfaz gráfica Tkinter
sap_imagenes.py         → Backend de captura
build_exe.bat           → Compila sap2000_capture.exe (onefile)
sap2000_portable.spec   → Spec de PyInstaller (onefile)
requirements-exe.txt    → Dependencias de build (incluye pyinstaller)
SAP2000_Capturas.xlsx   → Plantilla de configuración
```

## Requisitos

| Requisito | Versión |
|-----------|---------|
| Windows | 10 o 11 |
| SAP2000 | v23 |
| Python | 3.8+ (solo para compilar el EXE) |

## Flujo rápido

1. **Crea la plantilla Excel**: `sap2000_capture.exe --crear-excel SAP2000_Capturas.xlsx`
2. Completa las hojas `CONFIG` y `CAPTURAS`.
3. Abre SAP2000 con el modelo cargado y la ventana visible.
4. Ejecuta `sap2000_capture.exe`.

## Usar el programa

**Recomendado**: haz doble clic en `sap2000_capture.exe`.
Si existe `SAP2000_Capturas.xlsx` en la misma carpeta, el EXE lo detecta automáticamente.

```bat
sap2000_capture.exe --config SAP2000_Capturas.xlsx
```

Con Python (sin compilar):

```bat
python sap2000_gui.py --config SAP2000_Capturas.xlsx
```

### Interfaz gráfica

La GUI tiene dos botones:

1. **Conectar a SAP2000** — valida la conexión con SAP2000.
2. **Extraer fotos** — lee el Excel y genera las capturas.

### Opciones

```bat
sap2000_capture.exe --config RUTA --allow-unsafe-output
```

| Opción | Descripción |
|--------|-------------|
| `--config RUTA` | Ruta al Excel de configuración |
| `--crear-excel RUTA` | Genera la plantilla Excel y sale |
| `--allow-unsafe-output` | Permitir carpeta de salida fuera del directorio del Excel |

También puedes usar la variable de entorno `SAP2000_ALLOW_UNSAFE_OUTPUT=1`.

## Estructura del Excel

### Hoja `CONFIG`

| Celda | Descripción |
|-------|-------------|
| `B2` | Ruta del DLL de SAP2000 |
| `B3` | Nombre del proyecto |
| `B4` | Subcarpeta de salida |

La salida por defecto se guarda en `Capturas_SAP` como subcarpeta del Excel.

### Hoja `CAPTURAS`

Cada fila activa define una imagen a generar.

| Columna | Valor |
|---------|-------|
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

```
Capturas_SAP\
  MiProyecto_Vista_General_ISO_NE_MODELO.png
  MiProyecto_Cargas_Muertas_ISO_NE_CARGAS_DEAD.png
```

## EXE portable (compilar)

Build local (requiere PyInstaller):

```bat
build_exe.bat
```

Salida:

```
dist\sap2000_capture.exe
```

Un solo archivo. Solo necesitas distribuir el `.exe` junto al Excel de configuración.

## Solución de problemas

| Problema | Causa probable | Acción |
|----------|---------------|--------|
| SAP2000 no detectado | SAP2000 no está abierto o visible | Abre el modelo y deja la ventana visible |
| Error "ARCHIVO DLL NO ENCONTRADO" | Ruta del DLL incorrecta en CONFIG | Revisa la celda B2 del Excel |
| Error "ERROR AL CARGAR LA BIBLIOTECA DE TIPO/DLL" | DLL dañado o versión incorrecta | Verifica la instalación de SAP2000 |
| Error "ERROR DE PROGID COM" | SAP2000 mal instalado | Reinstala o repara SAP2000 |
| Error "SAP2000 NO ESTÁ ABIERTO O NO RESPONDE" | SAP2000 no está en ejecución | Abre SAP2000 manualmente |
| No se genera el `.xlsx` | No hay EXE ni Python | Compila el EXE o instala Python |
| Fila ignorada | Config inválida en `CAPTURAS` | Revisa vista, display, patrón y recortes |
| No escribe fuera de la carpeta del Excel | Modo seguro activo | Usa `--allow-unsafe-output` |
