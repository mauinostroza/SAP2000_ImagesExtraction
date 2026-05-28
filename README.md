# sap_capture

CLI modular para capturar vistas de SAP2000 de forma autónoma usando `PrintWindow` (Win32), sin depender de que la ventana esté al frente ni del flujo GUI+Excel anterior.

## Estructura

```
main.py
capture_plan.py
sap_bridge.py
view_controller.py
win32_capture.py
output_writer.py
requirements.txt
requirements-exe.txt
build_exe.bat
sap2000_portable.spec
sap2000_ci.spec
```

## Requisitos

| Requisito | Versión |
|-----------|---------|
| Windows | 10 o 11 |
| SAP2000 | v23, v24 o v25 |
| Python | 3.8+ |

## Instalación

```bat
pip install -r requirements.txt
python Scripts/pywin32_postinstall.py -install
```

## Flujo

```bat
python main.py --generate-plan
python main.py --list-cases
python main.py --plan capture_plan.json --output outputs/proyecto_01
```

Opciones habituales:

```bat
python main.py --plan plan.json --sap-dll "D:/SAP2000 23/SAP2000v1.dll"
python main.py --plan plan.json --render-delay 1.2
python main.py --plan plan.json -v
```

## Formato del plan

El plan puede ser JSON o Excel. Cada fila/entrada define una captura en orden.

```json
[
  {
    "filename": "vista_3d_muerta",
    "view_type": "ISO_3D",
    "display_type": "LOAD_CASE",
    "case_name": "DEAD",
    "description": "Vista isométrica carga muerta"
  },
  {
    "filename": "modo_1",
    "view_type": "ISO_3D",
    "display_type": "MODE_SHAPE",
    "case_name": "MODAL",
    "mode_number": 1
  }
]
```

Valores principales:

- `view_type`: `ISO_3D`, `PLAN_XY`, `ELEV_XZ`, `ELEV_YZ`
- `display_type`: `GEOMETRY_ONLY`, `LOAD_PATTERN`, `LOAD_CASE`, `MODE_SHAPE`, `DEFORMED`, `FRAME_FORCES`
- `case_name`: requerido para displays que dependen de un caso, patrón o combo
- `mode_number`: requerido para `MODE_SHAPE`

## Salida

```
outputs/proyecto_01/
  001_vista_3d_muerta.png
  002_modo_1.png
  capture_log.json
```

El `capture_log.json` registra estado, timestamp, duración y archivo de salida por captura.

## Build

```bat
build_exe.bat
```

Salida esperada:

```text
dist\sap_capture.exe
```

## Notas

- `PrintWindow` falla si SAP2000 está minimizado.
- El plan Excel requiere `openpyxl`.
- La conexión COM depende de `comtypes` y de una instalación válida de SAP2000.
