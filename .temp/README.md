# sap_capture — Captura autónoma de vistas SAP2000

Script Python portable que extrae imágenes de SAP2000 de forma autónoma
usando `PrintWindow` (Win32), sin depender de que la ventana esté en primer
plano ni interferir con el usuario.

---

## Estructura del proyecto

```
sap_capture/
├── main.py             ← CLI y orquestador principal
├── win32_capture.py    ← Motor de captura (FindWindow + PrintWindow)
├── sap_bridge.py       ← Conexión COM a SAP2000 (cSapModel)
├── view_controller.py  ← Control de vistas, casos y ángulos de cámara
├── capture_plan.py     ← Lector del plan desde JSON o Excel
├── output_writer.py    ← Naming, guardado PNG y log JSON
├── requirements.txt
└── capture_plan.json   ← (se genera con --generate-plan)
```

---

## Instalación (exe portable con Python embebido)

```bash
# Dentro del entorno Python del exe:
pip install -r requirements.txt

# Post-instalación pywin32 (necesario una sola vez):
python Scripts/pywin32_postinstall.py -install
```

---

## Flujo de uso

### 1. Generar plan de capturas de ejemplo
```bash
python main.py --generate-plan
```
Crea `capture_plan.json` con vistas típicas. Editar `case_name` con los
nombres reales del modelo.

### 2. Verificar casos disponibles en el modelo
Abrir SAP2000 con el modelo cargado, luego:
```bash
python main.py --list-cases
```
Imprime todos los patrones, casos y combinaciones disponibles.

### 3. Ejecutar capturas
```bash
python main.py --plan capture_plan.json --output outputs/proyecto_01
```

### 4. Opciones avanzadas
```bash
# Especificar DLL manualmente (si SAP2000 no está en ruta estándar)
python main.py --plan plan.json --sap-dll "D:/SAP2000 23/SAP2000v1.dll"

# Aumentar delay para modelos pesados
python main.py --plan plan.json --render-delay 1.2

# Logging detallado
python main.py --plan plan.json -v
```

---

## Formato del plan JSON

```json
[
  {
    "filename":     "vista_3d_muerta",
    "view_type":    "ISO_3D",
    "display_type": "LOAD_CASE",
    "case_name":    "DEAD",
    "description":  "Vista isométrica carga muerta"
  },
  {
    "filename":     "modo_1",
    "view_type":    "ISO_3D",
    "display_type": "MODE_SHAPE",
    "case_name":    "MODAL",
    "mode_number":  1
  }
]
```

### Valores válidos

**view_type:**
| Valor       | Vista                    |
|-------------|--------------------------|
| `ISO_3D`    | Isométrica 3D (default)  |
| `PLAN_XY`   | Planta (desde arriba)    |
| `ELEV_XZ`   | Elevación X-Z (frontal)  |
| `ELEV_YZ`   | Elevación Y-Z (lateral)  |

**display_type:**
| Valor            | Muestra                                    |
|------------------|--------------------------------------------|
| `GEOMETRY_ONLY`  | Solo geometría, sin cargas (default)       |
| `LOAD_PATTERN`   | Patrón de carga (requiere `case_name`)     |
| `LOAD_CASE`      | Caso de análisis (requiere `case_name`)    |
| `MODE_SHAPE`     | Forma modal (requiere `case_name` + `mode_number`) |
| `DEFORMED`       | Forma deformada (requiere `case_name`)     |
| `FRAME_FORCES`   | Fuerzas en barras (requiere análisis previo) |

---

## Salida generada

```
outputs/proyecto_01/
  001_geometria_3d.png
  002_geometria_planta.png
  003_carga_muerta_3d.png
  ...
  capture_log.json
```

El `capture_log.json` incluye por cada captura: estado (`ok`/`error`),
timestamp, tiempo de ejecución y ruta del archivo.

---

## Notas técnicas

### Por qué PrintWindow y no CopyFromScreen
- `CopyFromScreen` captura el buffer de pantalla: falla si SAP2000 está
  detrás de otra ventana o en monitor secundario.
- `PrintWindow` con `PW_RENDERFULLCONTENT=3` solicita al renderer de
  SAP2000 que dibuje directamente en un bitmap, sin importar la visibilidad.

### DPI Scaling
El script llama `SetProcessDpiAwareness(2)` al inicio. Sin esto, en monitores
con escala > 100%, `GetWindowRect` devuelve coordenadas virtuales y la
captura queda recortada o desplazada.

### SAP2000 no debe estar minimizado
`PrintWindow` falla si la ventana está minimizada (estado `SW_MINIMIZE`).
El script verifica el hwnd pero no restaura la ventana automáticamente.
Dejar SAP2000 en estado normal o maximizado antes de ejecutar.

### Versiones SAP2000 probadas
- v23 (CSI.SAP2000.API.SapObject): confirmado
- v24, v25: misma interfaz COM, compatible

### render_delay
Valor recomendado por tipo de display:
| Display             | Delay recomendado |
|---------------------|-------------------|
| GEOMETRY_ONLY       | 0.3 s             |
| LOAD_CASE / PATTERN | 0.5 s             |
| MODE_SHAPE          | 0.6 s             |
| DEFORMED            | 0.7 s             |
| FRAME_FORCES        | 0.8 s             |

Para modelos con >5000 elementos, subir a 1.0–1.5 s.
