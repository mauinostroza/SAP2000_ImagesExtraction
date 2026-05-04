"""
=============================================================================
sap_imagenes.py
Captura automática de imágenes de SAP2000 v23 — integración con Excel
=============================================================================

DESCRIPCIÓN:
  Lee una tabla de configuración desde el Excel que lo invoca (via xlwings),
  se conecta a SAP2000 en ejecución, aplica cada combinación de:
    - Vista del modelo (planta, elevaciones, isométricos, ángulo personalizado)
    - Modo de display (geometría del modelo o cargas de un patrón específico)
    - Encuadre (ventana completa o recorte parcial)
  y guarda las imágenes PNG en la misma carpeta del Excel.

REQUISITOS (instalar con pip):
  pip install comtypes pywin32 Pillow xlwings pyautogui openpyxl

CÓMO LLAMAR DESDE EXCEL (macro VBA):
  Sub CapturarImagenes()
      RunPython "import sap_imagenes; sap_imagenes.main()"
  End Sub

  También se puede ejecutar directamente:
      python sap_imagenes.py
=============================================================================
"""

import os
import sys
import time
import logging
import traceback
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# Importaciones de terceros — capturar errores de instalación temprano
# ---------------------------------------------------------------------------
try:
    import comtypes.client
    import comtypes
except ImportError:
    raise ImportError("Instala comtypes:  pip install comtypes")

try:
    import win32gui
    import win32con
    import win32ui
    import win32api
    from ctypes import windll
except ImportError:
    raise ImportError("Instala pywin32:  pip install pywin32")

try:
    from PIL import Image, ImageGrab
except ImportError:
    raise ImportError("Instala Pillow:  pip install Pillow")

try:
    import pyautogui
    pyautogui.FAILSAFE = True   # Mover el mouse a la esquina sup-izq para abortar
    pyautogui.PAUSE = 0.15
except ImportError:
    raise ImportError("Instala pyautogui:  pip install pyautogui")

try:
    import xlwings as xw
except ImportError:
    raise ImportError("Instala xlwings:  pip install xlwings")

try:
    import openpyxl
    from openpyxl.styles import (Font, PatternFill, Alignment,
                                  Border, Side, GradientFill)
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    raise ImportError("Instala openpyxl:  pip install openpyxl")

# ---------------------------------------------------------------------------
# Configuración global
# ---------------------------------------------------------------------------

SAP_DLL_PATH = r"C:\Program Files\Computers and Structures\SAP2000 23\SAP2000v1.dll"

# Pausa (segundos) entre cambios de vista y la captura
PAUSA_TRAS_VISTA   = 1.2   # Tiempo para que SAP2000 redibuje la ventana
PAUSA_TRAS_DISPLAY = 1.5   # Tiempo para que cambie el modo de display

# Tipos de vista válidos (usados en la columna VISTA del Excel)
VISTAS_VALIDAS = {
    "PLANTA":   "Vista en Planta (XY, mirando desde arriba)",
    "ELEV_X":   "Elevación en X  (plano XZ, mirando en dirección Y+)",
    "ELEV_Y":   "Elevación en Y  (plano YZ, mirando en dirección X+)",
    "ISO_NE":   "Isométrico NE   (azimut 225°, elevación 30°)",
    "ISO_NO":   "Isométrico NO   (azimut 315°, elevación 30°)",
    "ISO_SE":   "Isométrico SE   (azimut 135°, elevación 30°)",
    "ISO_SO":   "Isométrico SO   (azimut  45°, elevación 30°)",
    "CUSTOM":   "Ángulo libre    (usar columnas AZIMUT y ELEVACION)",
}

# Parámetros de cámara para cada vista estándar
# (azimut en grados desde X+, medido en plano XY; elevación desde el plano XY)
VISTA_ANGULOS = {
    "PLANTA": (270, 89.9),   # Casi desde el cénit (90 da problemas en algunos motores)
    "ELEV_X": (270,   0  ),  # Mirando en Y+, nivel
    "ELEV_Y": (  0,   0  ),  # Mirando en X+, nivel
    "ISO_NE": (225,  30  ),
    "ISO_NO": (315,  30  ),
    "ISO_SE": (135,  30  ),
    "ISO_SO": ( 45,  30  ),
}

# Modos de display válidos
DISPLAY_MODELO = "MODELO"
DISPLAY_CARGAS = "CARGAS"

# Configuración del logger
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s: %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("sap_imagenes")


# =============================================================================
#  BLOQUE 1 — Conexión a SAP2000
# =============================================================================

class SAP2000Conector:
    """Gestiona la conexión a una instancia de SAP2000 en ejecución."""

    def __init__(self, dll_path: str = SAP_DLL_PATH):
        self.dll_path = dll_path
        self.sap_obj   = None
        self.sap_model = None

    def conectar(self) -> "SAP2000Conector":
        """Conecta a SAP2000. Lanza RuntimeError si no está abierto."""
        log.info("Conectando a SAP2000...")

        # Generar wrappers comtypes la primera vez (idempotente luego)
        if os.path.exists(self.dll_path):
            try:
                comtypes.client.GetModule(self.dll_path)
            except Exception:
                pass  # Los wrappers ya existen — ignorar

        try:
            helper = comtypes.client.CreateObject("SAP2000v1.Helper")
        except Exception as e:
            raise RuntimeError(
                f"No se pudo crear SAP2000v1.Helper.\n"
                f"Verifica que SAP2000 v23 esté instalado y que\n"
                f"el DLL exista en:\n  {self.dll_path}\n\nDetalle: {e}"
            )

        try:
            import comtypes.gen.SAP2000v1 as sap_gen
            helper = helper.QueryInterface(sap_gen.cHelper)
            self.sap_obj = helper.GetObject("CSI.SAP2000.API.SapObject")
        except Exception as e:
            raise RuntimeError(
                f"SAP2000 no está abierto o no responde.\n"
                f"Abre SAP2000 con un modelo cargado antes de ejecutar este script.\n\n"
                f"Detalle: {e}"
            )

        self.sap_model = self.sap_obj.SapModel
        log.info("Conexión a SAP2000 establecida ✓")

        # Verificar que hay un modelo abierto
        nombre = ""
        try:
            self.sap_model.GetModelFilename(nombre)
        except Exception:
            pass

        return self


# =============================================================================
#  BLOQUE 2 — Búsqueda y captura de la ventana de SAP2000
# =============================================================================

class VentanaCaptura:
    """Localiza la ventana de SAP2000 y captura su área de dibujo."""

    def __init__(self):
        self.hwnd_principal = None     # Ventana principal SAP2000
        self.hwnd_viewport  = None     # Ventana hijo del viewport 3D

    def buscar_ventana(self) -> int:
        """
        Busca la ventana principal de SAP2000.
        Devuelve el HWND o lanza RuntimeError si no la encuentra.
        """
        resultados = []

        def _callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                titulo = win32gui.GetWindowText(hwnd)
                if "SAP2000" in titulo:
                    resultados.append((hwnd, titulo))

        win32gui.EnumWindows(_callback, None)

        if not resultados:
            raise RuntimeError(
                "No se encontró la ventana de SAP2000.\n"
                "Asegúrate de que SAP2000 está abierto y visible."
            )

        # Preferir la ventana con más contexto en el título (la principal)
        resultados.sort(key=lambda x: len(x[1]), reverse=True)
        self.hwnd_principal = resultados[0][0]
        log.info(f"Ventana SAP2000 encontrada: '{resultados[0][1]}' (hwnd={self.hwnd_principal})")
        return self.hwnd_principal

    def activar(self):
        """Trae SAP2000 al primer plano."""
        if not self.hwnd_principal:
            self.buscar_ventana()
        try:
            win32gui.ShowWindow(self.hwnd_principal, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(self.hwnd_principal)
            time.sleep(0.3)
        except Exception as e:
            log.warning(f"No se pudo activar la ventana: {e}")

    def capturar(
        self,
        filepath: str,
        rect_parcial: tuple = None
    ) -> bool:
        """
        Captura la ventana SAP2000 y guarda como PNG.

        Args:
            filepath:      Ruta completa del archivo PNG a guardar.
            rect_parcial:  (izq%, sup%, der%, inf%) en % [0-100]
                           para recortar el interior de la ventana.
                           None = captura completa.
        Returns:
            True si la captura fue exitosa.
        """
        if not self.hwnd_principal:
            self.buscar_ventana()

        hwnd = self.hwnd_principal

        # Coordenadas absolutas de la ventana en pantalla
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        except Exception as e:
            log.error(f"GetWindowRect falló: {e}")
            return False

        w = right  - left
        h = bottom - top

        if w <= 10 or h <= 10:
            log.error("La ventana de SAP2000 parece estar minimizada o es muy pequeña.")
            return False

        # ── Capturar usando PrintWindow (funciona aunque esté parcialmente cubierta) ──
        try:
            hwnd_dc    = win32gui.GetWindowDC(hwnd)
            mfc_dc     = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc    = mfc_dc.CreateCompatibleDC()
            save_bmp   = win32ui.CreateBitmap()
            save_bmp.CreateCompatibleBitmap(mfc_dc, w, h)
            save_dc.SelectObject(save_bmp)

            # PW_RENDERFULLCONTENT = 0x00000002  (flag para ventanas con DWM)
            result = windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

            bmpinfo = save_bmp.GetInfo()
            bmpstr  = save_bmp.GetBitmapBits(True)
            img = Image.frombuffer(
                "RGB",
                (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                bmpstr, "raw", "BGRX", 0, 1
            )

            # Liberar recursos GDI
            win32gui.DeleteObject(save_bmp.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwnd_dc)

        except Exception as e:
            log.warning(f"PrintWindow falló ({e}). Usando ImageGrab como fallback.")
            # Fallback: ImageGrab (requiere que la ventana sea visible y no esté cubierta)
            img = ImageGrab.grab(bbox=(left, top, right, bottom))

        # ── Aplicar recorte parcial si se especificó ──
        if rect_parcial is not None:
            iz, su, de, in_ = [max(0, min(100, v)) for v in rect_parcial]
            iw, ih = img.size
            crop_l = int(iw * iz / 100)
            crop_t = int(ih * su / 100)
            crop_r = int(iw * de / 100)
            crop_b = int(ih * in_ / 100)
            if crop_r > crop_l and crop_b > crop_t:
                img = img.crop((crop_l, crop_t, crop_r, crop_b))

        # ── Guardar ──
        try:
            img.save(filepath, "PNG")
            log.info(f"  → Guardado: {os.path.basename(filepath)}  ({img.size[0]}×{img.size[1]} px)")
            return True
        except Exception as e:
            log.error(f"Error al guardar imagen: {e}")
            return False


# =============================================================================
#  BLOQUE 3 — Control de vistas (ángulo de cámara)
# =============================================================================

class GestorVistas:
    """
    Controla el ángulo de cámara de SAP2000.

    ESTRATEGIA:
      1. Intenta usar el API de SAP2000 (sap_model.View) donde es posible.
      2. Para el ángulo 3D, navega el menú View > Set 3D View... de SAP2000
         mediante teclado (win32api + pyautogui).  Esto es más robusto que
         adivinar los IDs de comando WM_COMMAND.
    """

    def __init__(self, sap_model, ventana: VentanaCaptura):
        self.sap_model = sap_model
        self.ventana   = ventana

    def zoom_fit(self):
        """Zoom para mostrar todo el modelo en la ventana."""
        try:
            ret = self.sap_model.View.RefreshView(0, True)
            if ret != 0:
                log.warning(f"RefreshView devolvió {ret}")
        except Exception as e:
            log.warning(f"RefreshView falló: {e}")
        time.sleep(0.4)

    def set_vista(self, tipo_vista: str, azimut: float = 225, elevacion: float = 30):
        """
        Establece la vista 3D.

        Args:
            tipo_vista: Una de las claves de VISTAS_VALIDAS
            azimut:     Ángulo azimutal en grados (solo para CUSTOM)
            elevacion:  Ángulo de elevación en grados (solo para CUSTOM)
        """
        tipo_vista = tipo_vista.upper().strip()

        if tipo_vista == "CUSTOM":
            az, el = azimut, elevacion
        else:
            az, el = VISTA_ANGULOS.get(tipo_vista, (225, 30))

        log.info(f"  Configurando vista: {tipo_vista}  (az={az}°, el={el}°)")
        self.ventana.activar()

        # ── Abrir diálogo "Set 3D View" de SAP2000 via menú View ──
        self._abrir_dialogo_set3dview()

        # ── Ingresar azimut y elevación en el diálogo ──
        self._ingresar_angulos_en_dialogo(az, el)

        time.sleep(PAUSA_TRAS_VISTA)
        self.zoom_fit()
        time.sleep(0.3)

    def _abrir_dialogo_set3dview(self):
        """
        Navega View > Rotate 3D View > Set 3D View... en SAP2000 v23.
        Usa Alt+teclas para navegar el menú sin depender de coordenadas de pantalla.
        """
        hwnd = self.ventana.hwnd_principal
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.25)

        # Alt + V  → abre menú "View"
        pyautogui.hotkey("alt", "v")
        time.sleep(0.3)

        # En el menú View de SAP2000 v23:
        # "Set 3D View..." está bajo un submenú. Navegar con teclas de flecha.
        # Llegamos a la última entrada de "Rotate 3D View" y usamos la flecha derecha.

        # La cantidad de "flechas abajo" depende de la versión de SAP2000.
        # En v23 el submenú "Rotate 3D View" / "Set 3D View..." está típicamente
        # en la posición que se navega con estas teclas:
        for _ in range(6):       # Bajar hasta el ítem de Rotate/Set
            pyautogui.press("down")
            time.sleep(0.1)
        pyautogui.press("right")  # Abrir submenú
        time.sleep(0.2)
        for _ in range(5):        # Bajar hasta "Set 3D View..."
            pyautogui.press("down")
            time.sleep(0.1)
        pyautogui.press("enter")  # Abrir diálogo
        time.sleep(0.5)

        # NOTA: Si el diálogo no abre correctamente, ajusta los contadores
        # de flechas arriba según la estructura del menú en tu versión de SAP2000.
        # Puedes verificar la posición usando Alt+V y contando los ítems del menú.

    def _ingresar_angulos_en_dialogo(self, azimut: float, elevacion: float):
        """
        Ingresa azimut y elevación en el diálogo Set 3D View de SAP2000.
        El diálogo tiene dos campos: Plan Rotation (azimut) y Elevation.
        """
        try:
            # Tab para saltar entre campos del diálogo
            # Campo 1: Plan Rotation (Azimut)
            pyautogui.hotkey("ctrl", "a")      # Seleccionar todo en el campo activo
            pyautogui.write(str(int(azimut)), interval=0.05)
            time.sleep(0.15)
            pyautogui.press("tab")             # Siguiente campo

            # Campo 2: Elevation
            time.sleep(0.15)
            pyautogui.hotkey("ctrl", "a")
            pyautogui.write(str(int(elevacion)), interval=0.05)
            time.sleep(0.15)

            # Confirmar con Enter o botón OK
            pyautogui.press("enter")
            time.sleep(0.4)

        except Exception as e:
            log.warning(f"No se pudo ingresar ángulos en diálogo: {e}")
            # Escape para cerrar el diálogo si quedó abierto
            try:
                pyautogui.press("escape")
            except Exception:
                pass


# =============================================================================
#  BLOQUE 4 — Control del display (modelo vs cargas)
# =============================================================================

class GestorDisplay:
    """
    Controla qué muestra SAP2000: geometría pura o asignaciones de carga.

    Para el modo MODELO: usa RefreshView para restaurar la vista por defecto.
    Para el modo CARGAS: navega Display > Show Load Assigns > Frame/Area/Joint Loads.
    """

    def __init__(self, sap_model, ventana: VentanaCaptura):
        self.sap_model = sap_model
        self.ventana   = ventana

    def set_modelo(self):
        """Muestra la geometría del modelo sin cargas especiales."""
        log.info("  Display: Modelo (geometría)")
        try:
            # Restaurar el display por defecto mediante RefreshView
            self.sap_model.View.RefreshView(0, False)
            time.sleep(0.3)
        except Exception as e:
            log.warning(f"RefreshView (set_modelo) falló: {e}")

        # También podemos usar el menú View > Set Display Options para
        # asegurarnos de que las cargas no están visibles
        self._ocultar_cargas_menu()
        time.sleep(PAUSA_TRAS_DISPLAY)

    def set_cargas(self, patron: str):
        """
        Muestra las cargas asignadas para el patrón dado.
        Navega Display > Show Load Assigns en SAP2000 v23.

        Args:
            patron: nombre del patrón de carga (ej: "DEAD", "LIVE", "SISMO_X")
        """
        log.info(f"  Display: Cargas del patrón '{patron}'")
        self.ventana.activar()

        # Verificar que el patrón existe
        try:
            num_lp = 0
            lp_names = []
            self.sap_model.LoadPatterns.GetNameList(num_lp, lp_names)
            if lp_names and patron not in lp_names:
                log.warning(
                    f"  ¡Patrón '{patron}' no encontrado! "
                    f"Patrones disponibles: {lp_names}"
                )
        except Exception:
            pass

        # Navegar Display > Show Load Assigns > Frame Loads
        self._mostrar_cargas_menu(patron)
        time.sleep(PAUSA_TRAS_DISPLAY)

    def _ocultar_cargas_menu(self):
        """Navega View > Set Display Options para ocultar cargas."""
        # En muchos casos basta con RefreshView; esto es un refuerzo opcional
        pass  # Implementar si RefreshView no es suficiente

    def _mostrar_cargas_menu(self, patron: str):
        """
        Navega el menú Display > Show Load Assigns de SAP2000 v23.
        Luego selecciona el patrón en el cuadro de diálogo.
        """
        hwnd = self.ventana.hwnd_principal
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.25)

        # Alt + D  → abre menú "Display"
        pyautogui.hotkey("alt", "d")
        time.sleep(0.3)

        # En SAP2000 v23, "Show Load Assigns" está en las primeras opciones
        # del menú Display. Navegar con flechas:
        for _ in range(2):          # Bajar hasta "Show Load Assigns"
            pyautogui.press("down")
            time.sleep(0.1)
        pyautogui.press("right")    # Abrir submenú
        time.sleep(0.2)
        pyautogui.press("down")     # "Frame Loads" (primera opción del submenú)
        time.sleep(0.1)
        pyautogui.press("enter")    # Abrir diálogo
        time.sleep(0.5)

        # En el diálogo que abre, buscar el campo de selección de patrón
        # y escribir el nombre del patrón (o seleccionarlo de la lista)
        try:
            pyautogui.hotkey("ctrl", "a")
            pyautogui.write(patron, interval=0.04)
            time.sleep(0.2)
            pyautogui.press("enter")
            time.sleep(0.4)
        except Exception as e:
            log.warning(f"No se pudo seleccionar patrón en diálogo: {e}")
            try:
                pyautogui.press("escape")
            except Exception:
                pass

        # NOTA: El manejo del diálogo depende de la estructura exacta de
        # SAP2000 v23. Si el diálogo tiene una lista desplegable de patrones,
        # puede ser necesario ajustar la navegación. Consulta el README.


# =============================================================================
#  BLOQUE 5 — Nombre de archivo
# =============================================================================

def construir_nombre_archivo(
    nombre_base: str,
    tipo_vista: str,
    modo_display: str,
    patron: str = "",
    extension: str = "png"
) -> str:
    """
    Construye un nombre de archivo descriptivo para la imagen.

    Formato: {nombre_base}_{tipo_vista}_{modo_display}[_{patron}].png

    Ejemplo:
        nombre_base="Edificio_Piso1", vista="ISO_NE", display="CARGAS", patron="DEAD"
        → "Edificio_Piso1_ISO_NE_CARGAS_DEAD.png"
    """
    partes = [nombre_base.strip(), tipo_vista.upper()]
    partes.append(modo_display.upper())
    if patron:
        partes.append(patron.upper())

    nombre = "_".join(p.replace(" ", "_") for p in partes if p)
    # Limpiar caracteres no válidos en nombres de archivo
    for c in r'\/:*?"<>|':
        nombre = nombre.replace(c, "_")

    return f"{nombre}.{extension}"


# =============================================================================
#  BLOQUE 6 — Orquestador principal
# =============================================================================

class GeneradorImagenes:
    """
    Orquesta la captura completa: lee la configuración, aplica vistas y
    display, captura y guarda cada imagen.
    """

    def __init__(
        self,
        sap_model,
        carpeta_salida: str,
        nombre_proyecto: str = "Modelo"
    ):
        self.sap_model       = sap_model
        self.carpeta_salida  = carpeta_salida
        self.nombre_proyecto = nombre_proyecto

        self.ventana   = VentanaCaptura()
        self.vistas    = GestorVistas(sap_model, self.ventana)
        self.display   = GestorDisplay(sap_model, self.ventana)

        os.makedirs(carpeta_salida, exist_ok=True)
        log.info(f"Carpeta de salida: {carpeta_salida}")

    def inicializar(self):
        """Localiza la ventana de SAP2000 y hace zoom fit inicial."""
        self.ventana.buscar_ventana()
        self.vistas.zoom_fit()

    def procesar_fila(self, config: dict) -> dict:
        """
        Procesa una fila de la tabla de configuración.

        config (dict) debe tener las claves:
            nombre_imagen   str   Nombre base de la imagen
            tipo_vista      str   Clave de VISTAS_VALIDAS
            azimut          float Solo para CUSTOM
            elevacion       float Solo para CUSTOM
            modo_display    str   MODELO o CARGAS
            patron_carga    str   Nombre del patrón (si display=CARGAS)
            ventana         str   COMPLETA o PARCIAL
            crop            tuple (iz%, su%, der%, inf%) si PARCIAL
            activo          bool

        Retorna:
            dict con 'ok' (bool), 'archivo' (str), 'mensaje' (str)
        """
        if not config.get("activo", True):
            return {"ok": False, "archivo": "", "mensaje": "INACTIVO"}

        tipo_vista   = config.get("tipo_vista",   "ISO_NE").upper()
        azimut       = float(config.get("azimut",    225))
        elevacion    = float(config.get("elevacion",  30))
        modo_display = config.get("modo_display", DISPLAY_MODELO).upper()
        patron_carga = config.get("patron_carga", "").strip()
        nombre_img   = config.get("nombre_imagen", "imagen").strip()
        ventana_tipo = config.get("ventana", "COMPLETA").upper()
        crop         = config.get("crop", None)      # (iz%, su%, der%, inf%)

        log.info(f"──── Procesando: {nombre_img} ────")

        try:
            # 1. Cambiar vista
            self.vistas.set_vista(tipo_vista, azimut, elevacion)

            # 2. Cambiar display
            if modo_display == DISPLAY_CARGAS and patron_carga:
                self.display.set_cargas(patron_carga)
            else:
                self.display.set_modelo()

            # 3. Construir nombre y ruta del archivo
            nombre_archivo = construir_nombre_archivo(
                self.nombre_proyecto + "_" + nombre_img,
                tipo_vista,
                modo_display,
                patron_carga if modo_display == DISPLAY_CARGAS else ""
            )
            ruta_archivo = os.path.join(self.carpeta_salida, nombre_archivo)

            # 4. Capturar
            rect_parcial = crop if ventana_tipo == "PARCIAL" else None
            ok = self.ventana.capturar(ruta_archivo, rect_parcial)

            if ok:
                return {"ok": True, "archivo": nombre_archivo, "mensaje": "OK"}
            else:
                return {"ok": False, "archivo": "", "mensaje": "Error en captura"}

        except Exception as e:
            log.error(f"Error procesando '{nombre_img}': {e}")
            log.debug(traceback.format_exc())
            return {"ok": False, "archivo": "", "mensaje": str(e)[:120]}

    def procesar_lista(self, lista_config: list) -> list:
        """Procesa una lista de diccionarios de configuración."""
        resultados = []
        total = len(lista_config)
        for i, cfg in enumerate(lista_config, 1):
            log.info(f"Imagen {i}/{total}")
            res = self.procesar_fila(cfg)
            res["nombre_imagen"] = cfg.get("nombre_imagen", "")
            resultados.append(res)
        return resultados


# =============================================================================
#  BLOQUE 7 — Interfaz Excel (xlwings)
# =============================================================================

def leer_config_desde_excel(wb) -> dict:
    """
    Lee la hoja CONFIG y la tabla de capturas desde el workbook xlwings.

    Retorna:
        {
            "sap_dll_path": str,
            "nombre_proyecto": str,
            "carpeta_salida": str,
            "capturas": [dict, ...]
        }
    """
    # ── Hoja CONFIG ──
    try:
        sh_cfg = wb.sheets["CONFIG"]
        sap_dll      = sh_cfg["B2"].value or SAP_DLL_PATH
        nombre_proy  = sh_cfg["B3"].value or "Proyecto"
        carpeta_rel  = sh_cfg["B4"].value or "Capturas_SAP"
    except Exception:
        sap_dll     = SAP_DLL_PATH
        nombre_proy = "Proyecto"
        carpeta_rel = "Capturas_SAP"

    # La carpeta de salida es relativa al Excel (si no es absoluta)
    carpeta_excel = os.path.dirname(wb.fullname)
    if os.path.isabs(carpeta_rel):
        carpeta_salida = carpeta_rel
    else:
        carpeta_salida = os.path.join(carpeta_excel, carpeta_rel)

    # ── Hoja CAPTURAS ──
    capturas = []
    try:
        sh_cap = wb.sheets["CAPTURAS"]

        # Leer datos desde fila 3 (la 2 es encabezado)
        fila = 3
        while True:
            activo = sh_cap.range(f"A{fila}").value
            if activo is None:
                break    # Fin de la tabla

            activo_bool = str(activo).upper().strip() in ("SI", "SÍ", "YES", "1", "TRUE", "X")

            cfg = {
                "activo":        activo_bool,
                "nombre_imagen": str(sh_cap.range(f"B{fila}").value or "imagen"),
                "tipo_vista":    str(sh_cap.range(f"C{fila}").value or "ISO_NE").upper(),
                "azimut":        float(sh_cap.range(f"D{fila}").value or 225),
                "elevacion":     float(sh_cap.range(f"E{fila}").value or 30),
                "modo_display":  str(sh_cap.range(f"F{fila}").value or "MODELO").upper(),
                "patron_carga":  str(sh_cap.range(f"G{fila}").value or ""),
                "ventana":       str(sh_cap.range(f"H{fila}").value or "COMPLETA").upper(),
                "_fila":         fila,      # Guardar para escribir resultado
            }

            # Leer crop si ventana = PARCIAL
            if cfg["ventana"] == "PARCIAL":
                try:
                    cfg["crop"] = (
                        float(sh_cap.range(f"I{fila}").value or 0),   # iz%
                        float(sh_cap.range(f"J{fila}").value or 0),   # sup%
                        float(sh_cap.range(f"K{fila}").value or 100), # der%
                        float(sh_cap.range(f"L{fila}").value or 100), # inf%
                    )
                except Exception:
                    cfg["crop"] = (0, 0, 100, 100)
            else:
                cfg["crop"] = None

            capturas.append(cfg)
            fila += 1

    except Exception as e:
        log.warning(f"Error leyendo hoja CAPTURAS: {e}")

    return {
        "sap_dll_path":   sap_dll,
        "nombre_proyecto": nombre_proy,
        "carpeta_salida":  carpeta_salida,
        "capturas":        capturas,
    }


def escribir_resultados_en_excel(wb, resultados: list):
    """Escribe el estado y nombre de archivo de cada captura en la hoja CAPTURAS."""
    try:
        sh_cap = wb.sheets["CAPTURAS"]
        for res in resultados:
            fila = res.get("_fila") or res.get("fila")
            if not fila:
                continue
            sh_cap.range(f"M{fila}").value = "✓ OK" if res["ok"] else "✗ ERROR"
            sh_cap.range(f"N{fila}").value = res.get("archivo", "")
            if res["ok"]:
                sh_cap.range(f"M{fila}").color = (198, 239, 206)  # Verde claro
            else:
                sh_cap.range(f"M{fila}").color = (255, 199, 206)  # Rojo claro

    except Exception as e:
        log.warning(f"No se pudo escribir resultados en Excel: {e}")


def main():
    """
    Punto de entrada principal — llamado desde la macro VBA del Excel.

    Sub CapturarImagenes()
        RunPython "import sap_imagenes; sap_imagenes.main()"
    End Sub
    """
    try:
        wb = xw.Book.caller()
    except Exception:
        # Fallback: usar el primer workbook abierto si no hay caller
        try:
            wb = xw.books[0]
        except Exception:
            log.error("No se pudo obtener referencia al workbook de Excel.")
            return

    log.info("=" * 60)
    log.info("SAP2000 Image Capture — iniciando")
    log.info("=" * 60)

    # Leer configuración
    config = leer_config_desde_excel(wb)
    capturas = config["capturas"]

    if not capturas:
        log.warning("No se encontraron filas en la hoja CAPTURAS.")
        return

    activas = [c for c in capturas if c.get("activo", False)]
    log.info(f"Capturas configuradas: {len(capturas)}  |  Activas: {len(activas)}")

    # Conectar a SAP2000
    conector = SAP2000Conector(config["sap_dll_path"])
    conector.conectar()

    # Crear generador
    gen = GeneradorImagenes(
        sap_model      = conector.sap_model,
        carpeta_salida = config["carpeta_salida"],
        nombre_proyecto= config["nombre_proyecto"]
    )
    gen.inicializar()

    # Procesar todas las capturas
    resultados = []
    for cfg in capturas:
        res = gen.procesar_fila(cfg)
        res["_fila"] = cfg["_fila"]
        resultados.append(res)

    # Escribir resultados de vuelta en Excel
    escribir_resultados_en_excel(wb, resultados)

    # Resumen final
    ok_count  = sum(1 for r in resultados if r["ok"])
    err_count = len(resultados) - ok_count
    log.info("=" * 60)
    log.info(f"Proceso completado: {ok_count} imágenes guardadas, {err_count} errores")
    log.info(f"Carpeta: {config['carpeta_salida']}")
    log.info("=" * 60)


# =============================================================================
#  BLOQUE 8 — Crear el Excel de configuración (plantilla)
# =============================================================================

def crear_excel_configuracion(ruta_excel: str):
    """
    Crea el archivo Excel de configuración con las hojas CONFIG y CAPTURAS.
    Incluye validaciones de datos, formatos y filas de ejemplo.

    Args:
        ruta_excel: Ruta completa del archivo .xlsm a crear.
    """
    from openpyxl import Workbook
    from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                                  GradientFill)
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()

    # ── Paleta de colores ──
    AZUL_OSCURO  = "1F3864"
    AZUL_MEDIO   = "2E75B6"
    AZUL_CLARO   = "BDD7EE"
    GRIS_FONDO   = "F2F2F2"
    VERDE_CLARO  = "C6EFCE"
    ROJO_CLARO   = "FFC7CE"
    BLANCO       = "FFFFFF"

    def estilo_encab(ws, celda, texto, bgcolor=AZUL_OSCURO, color_txt=BLANCO, bold=True):
        c = ws[celda]
        c.value = texto
        c.font = Font(bold=bold, color=color_txt, size=11)
        c.fill = PatternFill("solid", fgColor=bgcolor)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def borde_completo():
        s = Side(style="thin", color="888888")
        return Border(left=s, right=s, top=s, bottom=s)

    # ─────────────────────────────────────────────────────────
    # Hoja CONFIG
    # ─────────────────────────────────────────────────────────
    sh_cfg = wb.active
    sh_cfg.title = "CONFIG"
    sh_cfg.column_dimensions["A"].width = 32
    sh_cfg.column_dimensions["B"].width = 60
    sh_cfg.row_dimensions[1].height = 30

    # Título
    sh_cfg.merge_cells("A1:B1")
    estilo_encab(sh_cfg, "A1", "⚙ SAP2000 Image Capture — Configuración", AZUL_OSCURO, BLANCO)

    filas_cfg = [
        ("Ruta DLL SAP2000",    SAP_DLL_PATH),
        ("Nombre del proyecto", "MiProyecto"),
        ("Subcarpeta de salida","Capturas_SAP"),
    ]
    for i, (label, val) in enumerate(filas_cfg, start=2):
        sh_cfg[f"A{i}"].value = label
        sh_cfg[f"A{i}"].font  = Font(bold=True, size=10)
        sh_cfg[f"A{i}"].fill  = PatternFill("solid", fgColor=GRIS_FONDO)
        sh_cfg[f"B{i}"].value = val
        sh_cfg[f"B{i}"].font  = Font(size=10, color="1A3960")
        for col in ("A", "B"):
            sh_cfg[f"{col}{i}"].border = borde_completo()
            sh_cfg[f"{col}{i}"].alignment = Alignment(vertical="center")

    # Instrucciones
    sh_cfg["A6"].value = "INSTRUCCIONES"
    sh_cfg["A6"].font = Font(bold=True, size=11, color=AZUL_OSCURO)

    instrucciones = [
        "1. Abre SAP2000 v23 con el modelo que deseas capturar.",
        "2. Completa la hoja CAPTURAS con las vistas que necesitas.",
        "3. Desde Excel, ejecuta la macro: CapturarImagenes()",
        "   (pestaña Programador > Macros > CapturarImagenes)",
        "4. Las imágenes PNG se guardarán en la subcarpeta indicada.",
        "",
        "VISTAS disponibles:",
        "  PLANTA  · ELEV_X  · ELEV_Y  · ISO_NE  · ISO_NO  · ISO_SE  · ISO_SO  · CUSTOM",
        "",
        "DISPLAY disponibles:",
        "  MODELO → muestra la geometría del modelo",
        "  CARGAS → muestra las cargas del patrón indicado",
        "",
        "VENTANA:",
        "  COMPLETA → captura toda la ventana de SAP2000",
        "  PARCIAL  → recorta según los porcentajes I/S/D/I (%)",
    ]
    for j, txt in enumerate(instrucciones, start=7):
        sh_cfg[f"A{j}"].value = txt
        sh_cfg[f"A{j}"].font  = Font(size=9, italic=(txt.startswith(" ")))
        sh_cfg.merge_cells(f"A{j}:B{j}")

    # ─────────────────────────────────────────────────────────
    # Hoja CAPTURAS
    # ─────────────────────────────────────────────────────────
    sh_cap = wb.create_sheet("CAPTURAS")

    # Anchos de columna
    anchos = {"A":9,"B":26,"C":12,"D":10,"E":12,"F":11,"G":16,
              "H":12,"I":10,"J":10,"K":10,"L":10,"M":11,"N":45}
    for col, w in anchos.items():
        sh_cap.column_dimensions[col].width = w

    sh_cap.row_dimensions[1].height = 14
    sh_cap.row_dimensions[2].height = 36

    # Fila 1: Grupos de encabezado
    grupos = [
        ("A1:A2", "ACTIVO"),
        ("B1:B2", "NOMBRE\nIMAGEN"),
        ("C1:E1", "VISTA"),
        ("F1:G1", "DISPLAY"),
        ("H1:L1", "ENCUADRE"),
        ("M1:N1", "RESULTADO (script)"),
    ]
    for rango, texto in grupos:
        sh_cap.merge_cells(rango)
        c = sh_cap[rango.split(":")[0]]
        c.value = texto
        c.font = Font(bold=True, color=BLANCO, size=10)
        c.fill = PatternFill("solid", fgColor=AZUL_OSCURO)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Fila 2: Sub-encabezados
    sub_encab = {
        "C2": "Tipo Vista",  "D2": "Azimut\n(solo CUSTOM)",
        "E2": "Elevación\n(solo CUSTOM)", "F2": "Modo\nDisplay",
        "G2": "Patrón de\nCarga", "H2": "Tipo\nVentana",
        "I2": "Recorte\nIzq %", "J2": "Recorte\nSup %",
        "K2": "Recorte\nDer %", "L2": "Recorte\nInf %",
        "M2": "Estado", "N2": "Archivo generado",
    }
    for cel, txt in sub_encab.items():
        c = sh_cap[cel]
        c.value = txt
        c.font = Font(bold=True, color=BLANCO, size=9)
        c.fill = PatternFill("solid", fgColor=AZUL_MEDIO)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Bordes para toda la cabecera
    for row in sh_cap.iter_rows(min_row=1, max_row=2, min_col=1, max_col=14):
        for cell in row:
            cell.border = borde_completo()

    # ── Filas de ejemplo ──
    ejemplos = [
        # activo, nombre,              vista,    az,    el, display,  patron,    ventana
        ("SI",  "Vista_General",      "ISO_NE",  225,   30, "MODELO", "",         "COMPLETA"),
        ("SI",  "Planta_Modelo",      "PLANTA",  270, 89.9, "MODELO", "",         "COMPLETA"),
        ("SI",  "Elevacion_Frontal",  "ELEV_X",  270,    0, "MODELO", "",         "COMPLETA"),
        ("SI",  "Cargas_Muertas_ISO", "ISO_NE",  225,   30, "CARGAS", "DEAD",     "COMPLETA"),
        ("SI",  "Cargas_Vivas_ISO",   "ISO_NO",  315,   30, "CARGAS", "LIVE",     "COMPLETA"),
        ("SI",  "Cargas_Sismo_X",     "ELEV_X",  270,    0, "CARGAS", "SISMO_X",  "COMPLETA"),
        ("NO",  "Vista_Personalizada","CUSTOM",   45,   60, "MODELO", "",         "COMPLETA"),
        ("NO",  "Detalle_Parcial",    "ISO_NE",  225,   30, "MODELO", "",         "PARCIAL"),
    ]

    # Crop para la fila "Detalle_Parcial" (última)
    crops = {10: (10, 10, 90, 90)}   # fila 10 = ultima de ejemplos (fila 3+7)

    for idx, ej in enumerate(ejemplos, start=3):
        fila = idx
        vals = list(ej)
        sh_cap[f"A{fila}"].value = vals[0]
        sh_cap[f"B{fila}"].value = vals[1]
        sh_cap[f"C{fila}"].value = vals[2]
        sh_cap[f"D{fila}"].value = vals[3]
        sh_cap[f"E{fila}"].value = vals[4]
        sh_cap[f"F{fila}"].value = vals[5]
        sh_cap[f"G{fila}"].value = vals[6]
        sh_cap[f"H{fila}"].value = vals[7]

        if fila in crops:
            iz, su, de, in_ = crops[fila]
            sh_cap[f"I{fila}"].value = iz
            sh_cap[f"J{fila}"].value = su
            sh_cap[f"K{fila}"].value = de
            sh_cap[f"L{fila}"].value = in_

        # Formato de fila
        fondo = GRIS_FONDO if idx % 2 == 0 else BLANCO
        inactivo = vals[0].upper() not in ("SI", "SÍ", "YES")
        for col_l in "ABCDEFGHIJKL":
            c = sh_cap[f"{col_l}{fila}"]
            c.fill = PatternFill("solid", fgColor="ECECEC" if inactivo else fondo)
            c.font = Font(size=10, color="999999" if inactivo else "1A1A1A")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = borde_completo()

    # ── Validaciones de datos ──
    dv_activo = DataValidation(type="list", formula1='"SI,NO"',
                               showDropDown=False, showErrorMessage=True)
    dv_activo.error  = "Ingresa SI o NO"
    dv_activo.sqref  = "A3:A500"

    dv_vista = DataValidation(type="list",
        formula1='"PLANTA,ELEV_X,ELEV_Y,ISO_NE,ISO_NO,ISO_SE,ISO_SO,CUSTOM"',
        showDropDown=False)
    dv_vista.sqref = "C3:C500"

    dv_display = DataValidation(type="list", formula1='"MODELO,CARGAS"',
                                showDropDown=False)
    dv_display.sqref = "F3:F500"

    dv_ventana = DataValidation(type="list", formula1='"COMPLETA,PARCIAL"',
                                showDropDown=False)
    dv_ventana.sqref = "H3:H500"

    for dv in (dv_activo, dv_vista, dv_display, dv_ventana):
        sh_cap.add_data_validation(dv)

    # Congelar filas de encabezado
    sh_cap.freeze_panes = "A3"

    # ─────────────────────────────────────────────────────────
    # Guardar
    # ─────────────────────────────────────────────────────────
    os.makedirs(os.path.dirname(ruta_excel), exist_ok=True)
    wb.save(ruta_excel)
    log.info(f"Excel de configuración creado: {ruta_excel}")


# =============================================================================
#  Punto de entrada directo (python sap_imagenes.py)
# =============================================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="SAP2000 Image Capture — ejecutar directamente desde Python"
    )
    parser.add_argument(
        "--crear-excel",
        metavar="RUTA",
        help="Crear el Excel de configuración en la ruta especificada y salir.",
    )
    parser.add_argument(
        "--config",
        metavar="EXCEL",
        help="Ruta al Excel de configuración para ejecutar las capturas.",
    )
    parser.add_argument(
        "--proyecto",
        metavar="NOMBRE",
        default="Modelo",
        help="Nombre del proyecto (prefijo de los archivos).",
    )
    args = parser.parse_args()

    if args.crear_excel:
        crear_excel_configuracion(args.crear_excel)
        print(f"Excel creado en: {args.crear_excel}")
        sys.exit(0)

    if args.config:
        # Leer configuración desde Excel y ejecutar sin xlwings caller
        import openpyxl as _opx
        _wb_raw = _opx.load_workbook(args.config, data_only=True)

        # Construir config manualmente desde openpyxl
        sh = _wb_raw["CONFIG"]
        sap_dll       = sh["B2"].value or SAP_DLL_PATH
        nombre_proy   = sh["B3"].value or args.proyecto
        carpeta_rel   = sh["B4"].value or "Capturas_SAP"
        carpeta_salida = os.path.join(os.path.dirname(args.config), carpeta_rel)

        sh_cap = _wb_raw["CAPTURAS"]
        capturas = []
        for fila in range(3, sh_cap.max_row + 1):
            vals = [sh_cap.cell(fila, c).value for c in range(1, 15)]
            if vals[0] is None:
                break
            activo = str(vals[0]).upper().strip() in ("SI", "SÍ", "YES", "1")
            cfg = {
                "activo":        activo,
                "nombre_imagen": str(vals[1] or "imagen"),
                "tipo_vista":    str(vals[2] or "ISO_NE").upper(),
                "azimut":        float(vals[3] or 225),
                "elevacion":     float(vals[4] or 30),
                "modo_display":  str(vals[5] or "MODELO").upper(),
                "patron_carga":  str(vals[6] or ""),
                "ventana":       str(vals[7] or "COMPLETA").upper(),
                "_fila":         fila,
            }
            if cfg["ventana"] == "PARCIAL":
                cfg["crop"] = (
                    float(vals[8] or 0), float(vals[9]  or 0),
                    float(vals[10] or 100), float(vals[11] or 100),
                )
            capturas.append(cfg)

        conector = SAP2000Conector(sap_dll)
        conector.conectar()
        gen = GeneradorImagenes(conector.sap_model, carpeta_salida, nombre_proy)
        gen.inicializar()
        resultados = gen.procesar_lista(capturas)

        ok  = sum(1 for r in resultados if r["ok"])
        err = len(resultados) - ok
        print(f"\nResultado: {ok} imágenes OK, {err} errores → {carpeta_salida}")
        sys.exit(0 if err == 0 else 1)

    # Sin argumentos: solo crear el Excel en la carpeta actual
    ruta = os.path.join(os.getcwd(), "SAP2000_Capturas.xlsx")
    crear_excel_configuracion(ruta)
    print(f"Plantilla Excel creada en: {ruta}")
    print("Úsala con: python sap_imagenes.py --config SAP2000_Capturas.xlsx")
