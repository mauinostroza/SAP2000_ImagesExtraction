# SAP2000 Image Capture — Guía completa

Captura automática de imágenes de modelos SAP2000 v23, controlada desde Excel.

---

## Archivos del paquete

```
SAP_Capturas/
├── sap_imagenes.py          ← Módulo principal (toda la lógica)
├── crear_excel.py           ← Crea la plantilla Excel desde cero
├── SAP_Capturas_VBA.bas     ← Módulo VBA para botón en Excel
├── ejecutar_capturas.bat    ← Ejecutar directamente (sin macro)
└── README.md
```

---

## Requisitos

| Requisito | Versión |
|---|---|
| Python | 3.8+ (64 bits) |
| SAP2000 | v23 (puede funcionar en v22/v24 con ajustes) |
| Windows | 10 o 11 |

### Instalar dependencias Python

```bat
pip install comtypes pywin32 Pillow pyautogui xlwings openpyxl
```

---

## Configuración inicial (una sola vez)

### Paso 1 — Crear el Excel de configuración

```bat
python crear_excel.py
```

Esto crea `SAP2000_Capturas.xlsx` en la misma carpeta del script.

### Paso 2 — Agregar el módulo VBA al Excel

1. Abre `SAP2000_Capturas.xlsx` en Excel
2. Guárdalo como `.xlsm` (habilitado para macros)
3. Abre el editor VBA: `Alt + F11`
4. Haz clic derecho en "VBAProject" > **Importar archivo**
5. Selecciona `SAP_Capturas_VBA.bas`
6. Agrega un botón en la hoja CONFIG y asígnale la macro `CapturarImagenes`

### Paso 3 — Instalar xlwings (para el botón Excel)

```bat
pip install xlwings
xlwings addin install
```

---

## Uso

### Método A — Desde Excel (recomendado)

1. Abre SAP2000 con tu modelo cargado
2. Abre `SAP2000_Capturas.xlsm`
3. Completa la hoja **CAPTURAS** (ver sección siguiente)
4. Haz clic en el botón **"Capturar Imágenes"**
5. Las imágenes PNG se guardarán en la carpeta indicada en CONFIG

### Método B — Desde línea de comandos

```bat
python sap_imagenes.py --config SAP2000_Capturas.xlsx
```

### Método C — Doble clic

1. Copia `ejecutar_capturas.bat` junto al Excel
2. Doble clic en el `.bat` (con SAP2000 abierto)

---

## Configurar la hoja CAPTURAS

Cada fila de la tabla define una captura. Columnas disponibles:

| Columna | Opciones | Descripción |
|---|---|---|
| **ACTIVO** | SI / NO | Activar o desactivar esta captura |
| **NOMBRE IMAGEN** | texto libre | Nombre base del archivo (sin extensión) |
| **TIPO VISTA** | `PLANTA` `ELEV_X` `ELEV_Y` `ISO_NE` `ISO_NO` `ISO_SE` `ISO_SO` `CUSTOM` | Ángulo de cámara |
| **AZIMUT** | 0–360 | Solo para CUSTOM — ángulo horizontal |
| **ELEVACIÓN** | 0–90 | Solo para CUSTOM — ángulo vertical |
| **MODO DISPLAY** | `MODELO` / `CARGAS` | Qué mostrar en la vista |
| **PATRÓN DE CARGA** | nombre del patrón | Solo para DISPLAY=CARGAS (ej: DEAD, LIVE) |
| **TIPO VENTANA** | `COMPLETA` / `PARCIAL` | Capturar toda la ventana o un recorte |
| **RECORTE IZQ %** | 0–100 | Solo para PARCIAL |
| **RECORTE SUP %** | 0–100 | Solo para PARCIAL |
| **RECORTE DER %** | 0–100 | Solo para PARCIAL (ej: 90 = hasta 90% desde la izq) |
| **RECORTE INF %** | 0–100 | Solo para PARCIAL |

### Ejemplo de tabla CAPTURAS

| ACTIVO | NOMBRE IMAGEN | TIPO VISTA | AZIMUT | ELEV | DISPLAY | PATRÓN |
|---|---|---|---|---|---|---|
| SI | Vista_General | ISO_NE | — | — | MODELO | — |
| SI | Planta | PLANTA | — | — | MODELO | — |
| SI | Elevacion_Norte | ELEV_X | — | — | MODELO | — |
| SI | Cargas_Muertas | ISO_NE | — | — | CARGAS | DEAD |
| SI | Cargas_Vivas | ISO_NO | — | — | CARGAS | LIVE |
| SI | Sismo_X_Elev | ELEV_X | — | — | CARGAS | SISMO_X |
| NO | Vista_Custom | CUSTOM | 45 | 60 | MODELO | — |

Los archivos PNG generados tendrán nombres como:
```
MiProyecto_Vista_General_ISO_NE_MODELO.png
MiProyecto_Cargas_Muertas_ISO_NE_CARGAS_DEAD.png
```

---

## Ajustar la navegación de menús de SAP2000

El script navega los menús de SAP2000 mediante teclado para cambiar vistas y display.
La navegación puede variar entre versiones de SAP2000. Si las vistas no cambian
correctamente, ajusta los **contadores de flechas** en estas funciones de `sap_imagenes.py`:

### `_abrir_dialogo_set3dview()` — para cambiar el ángulo de vista

```python
# Líneas a ajustar (cuenta los ítems de tu menú View):
for _ in range(6):       # ← Cambiar este número (ítems hasta "Rotate 3D View")
    pyautogui.press("down")

for _ in range(5):        # ← Cambiar este número (ítems hasta "Set 3D View...")
    pyautogui.press("down")
```

**Cómo verificarlo**: Abre SAP2000, presiona `Alt+V` para abrir el menú View,
y cuenta las posiciones hacia abajo hasta llegar a "Set 3D View...".

### `_mostrar_cargas_menu()` — para mostrar cargas

```python
# Líneas a ajustar (ítems en menú Display):
for _ in range(2):          # ← Cambiar según posición de "Show Load Assigns"
    pyautogui.press("down")
```

---

## Solución de problemas

| Problema | Causa | Solución |
|---|---|---|
| "SAP2000 no detectado" | SAP2000 no está abierto | Abre SAP2000 con el modelo antes |
| Vista no cambia | Contadores de flechas incorrectos | Ajusta `_abrir_dialogo_set3dview()` |
| Display de cargas no aparece | Menú Display diferente | Ajusta `_mostrar_cargas_menu()` |
| Error xlwings | xlwings no instalado | `pip install xlwings` y `xlwings addin install` |
| Error `GetModule` | DLL de SAP2000 no encontrado | Verifica la ruta en hoja CONFIG |
| Imagen en negro | Ventana SAP2000 minimizada | Mantén SAP2000 visible |
| Imagen recortada mal | % de crop invertidos | Recuerda: Izq% y Sup% = posición inicial; Der% e Inf% = posición final |

### Verificar la ruta del DLL de SAP2000

En la hoja **CONFIG**, celda B2, verifica que la ruta sea correcta para tu instalación:
```
C:\Program Files\Computers and Structures\SAP2000 23\SAP2000v1.dll
```

---

## Personalización avanzada

### Cambiar el tiempo de espera entre vistas

En `sap_imagenes.py`, líneas cerca del inicio:

```python
PAUSA_TRAS_VISTA   = 1.2   # ← Aumentar si SAP2000 es lento en tu equipo
PAUSA_TRAS_DISPLAY = 1.5
```

### Agregar nuevos ángulos predefinidos

En `VISTA_ANGULOS` al inicio del script:

```python
VISTA_ANGULOS = {
    ...
    "ISO_TECHO": (180, 80),     # Vista casi cenital desde el norte
    "PERSPECTIVA": (200, 15),   # Perspectiva baja dramática
}
```

Y agregar la clave a `VISTAS_VALIDAS` con su descripción.
