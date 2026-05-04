"""
=============================================================================
sap_imagenes.py
Captura automática de imágenes de SAP2000 v23 — integración con Excel
=============================================================================

DESCRIPCIÓN:
  Lee una tabla de configuración desde un archivo Excel por CLI
  o, de forma opcional, desde un workbook abierto vía xlwings,
  se conecta a SAP2000 en ejecución, aplica cada combinación de:
    - Vista del modelo (planta, elevaciones, isométricos, ángulo personalizado)
    - Modo de display (geometría del modelo o cargas de un patrón específico)
    - Encuadre (ventana completa o recorte parcial)
  y guarda las imágenes PNG en la misma carpeta del Excel.

REQUISITOS (instalar con pip):
  pip install comtypes pywin32 Pillow pyautogui openpyxl
  pip install xlwings   # solo si se usará integración opcional con workbook abierto

EJECUCIÓN DIRECTA:
  python sap_imagenes.py --config SAP2000_Capturas.xlsx
=============================================================================
"""

import os
import sys
import time
import logging
import traceback
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Importaciones de terceros
# ---------------------------------------------------------------------------
try:
    import openpyxl
except ImportError:
    raise ImportError("Instala openpyxl:  pip install openpyxl")

comtypes = None
win32gui = None
win32con = None
Image = None
pyautogui = None
xw = None


def _ensure_comtypes():
    global comtypes
    if comtypes is not None:
        return comtypes
    try:
        import comtypes.client as comtypes_client
        import comtypes as _comtypes
    except ImportError as exc:
        raise ImportError("Instala comtypes:  pip install comtypes") from exc

    _comtypes.client = comtypes_client
    comtypes = _comtypes
    return comtypes


def _ensure_windows_capture_libs():
    global win32gui, win32con, Image
    if all(mod is not None for mod in (win32gui, win32con, Image)):
        return

    try:
        import win32gui as _win32gui
        import win32con as _win32con
    except ImportError as exc:
        raise ImportError("Instala pywin32:  pip install pywin32") from exc

    try:
        from PIL import Image as _Image
    except ImportError as exc:
        raise ImportError("Instala Pillow:  pip install Pillow") from exc

    win32gui = _win32gui
    win32con = _win32con
    Image = _Image


def _ensure_pillow():
    global Image
    if Image is not None:
        return Image
    try:
        from PIL import Image as _Image
    except ImportError as exc:
        raise ImportError("Instala Pillow:  pip install Pillow") from exc
    Image = _Image
    return Image


def _ensure_pyautogui():
    global pyautogui
    if pyautogui is not None:
        return pyautogui
    try:
        import pyautogui as _pyautogui
    except ImportError as exc:
        raise ImportError("Instala pyautogui:  pip install pyautogui") from exc

    _pyautogui.FAILSAFE = True
    _pyautogui.PAUSE = 0.15
    pyautogui = _pyautogui
    return pyautogui


def _ensure_xlwings():
    global xw
    if xw is not None:
        return xw
    try:
        import xlwings as _xw
    except ImportError as exc:
        raise ImportError("Instala xlwings:  pip install xlwings") from exc

    xw = _xw
    return xw

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

SAFE_OUTPUT_ENVVAR = "SAP2000_ALLOW_UNSAFE_OUTPUT"


def permitir_salida_insegura(explicito: bool = False) -> bool:
    """Permite rutas de salida fuera de la carpeta base solo con opt-in claro."""
    if explicito:
        return True
    return valor_verdadero(os.environ.get(SAFE_OUTPUT_ENVVAR, ""))


def _resolver_hijo_seguro(base_dir: str, ruta_configurada: str) -> str:
    base_path = Path(base_dir).resolve()
    destino = (base_path / ruta_configurada).resolve()
    try:
        destino.relative_to(base_path)
    except ValueError as exc:
        raise ValueError(
            f"La ruta de salida '{ruta_configurada}' sale de la carpeta base '{base_path}'."
        ) from exc
    return str(destino)


def valor_verdadero(valor) -> bool:
    """Normaliza celdas SI/NO o booleanos a bool."""
    return str(valor).upper().strip() in ("SI", "SÍ", "YES", "1", "TRUE", "X")


def validar_tipo_vista(tipo_vista: str) -> str:
    tipo_vista = str(tipo_vista or "ISO_NE").upper().strip()
    if tipo_vista not in VISTAS_VALIDAS:
        raise ValueError(
            f"Vista inválida: '{tipo_vista}'. "
            f"Opciones válidas: {', '.join(VISTAS_VALIDAS.keys())}"
        )
    return tipo_vista


def normalizar_crop(rect_parcial):
    """Normaliza y valida un crop porcentual (izq, sup, der, inf)."""
    if rect_parcial is None:
        return None
    try:
        iz, su, de, inf_ = [max(0.0, min(100.0, float(v))) for v in rect_parcial]
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Crop inválido: {rect_parcial}") from exc
    if de <= iz or inf_ <= su:
        raise ValueError(
            f"Crop inválido: {rect_parcial}. "
            "Se requiere der > izq e inf > sup."
        )
    return (iz, su, de, inf_)


def cargar_fila_captura(fila: int, valores) -> dict:
    """Construye y valida una fila de configuración desde una secuencia de valores."""
    cfg = {
        "activo":        valor_verdadero(valores[0]),
        "nombre_imagen": str(valores[1] or "imagen").strip(),
        "tipo_vista":    validar_tipo_vista(valores[2]),
        "azimut":        float(valores[3] or 225),
        "elevacion":     float(valores[4] or 30),
        "modo_display":  str(valores[5] or DISPLAY_MODELO).upper().strip(),
        "patron_carga":  str(valores[6] or "").strip(),
        "ventana":       str(valores[7] or "COMPLETA").upper().strip(),
        "_fila":         fila,
    }
    if cfg["modo_display"] not in (DISPLAY_MODELO, DISPLAY_CARGAS):
        raise ValueError(
            f"Modo display inválido en fila {fila}: '{cfg['modo_display']}'. "
            f"Usa {DISPLAY_MODELO} o {DISPLAY_CARGAS}."
        )
    if cfg["ventana"] not in ("COMPLETA", "PARCIAL"):
        raise ValueError(
            f"Tipo de ventana inválido en fila {fila}: '{cfg['ventana']}'. "
            "Usa COMPLETA o PARCIAL."
        )
    if cfg["modo_display"] == DISPLAY_CARGAS and not cfg["patron_carga"]:
        raise ValueError(
            f"Fila {fila}: display=CARGAS requiere un patrón de carga."
        )
    if cfg["ventana"] == "PARCIAL":
        cfg["crop"] = normalizar_crop((
            valores[8] or 0,
            valores[9] or 0,
            valores[10] or 100,
            valores[11] or 100,
        ))
    else:
        cfg["crop"] = None
    return cfg


def resolver_carpeta_salida(
    base_dir: str,
    carpeta_configurada: str,
    allow_unsafe: bool = False,
) -> str:
    """
    Resuelve la carpeta de salida en modo seguro por defecto.

    Solo se aceptan subdirectorios dentro de ``base_dir`` salvo opt-in explícito.
    """
    carpeta_configurada = str(carpeta_configurada or "Capturas_SAP").strip()
    if not carpeta_configurada:
        carpeta_configurada = "Capturas_SAP"

    if allow_unsafe:
        ruta = Path(carpeta_configurada)
        if not ruta.is_absolute():
            ruta = Path(base_dir) / ruta
        return str(ruta.resolve())

    if os.path.isabs(carpeta_configurada):
        raise ValueError(
            "Las rutas absolutas de salida están deshabilitadas en modo seguro. "
            "Usa una subcarpeta relativa o habilita la salida insegura explícitamente."
        )

    return _resolver_hijo_seguro(base_dir, carpeta_configurada)


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
        _ensure_comtypes()

        # Generar wrappers comtypes la primera vez (idempotente luego)
        sap_gen = None
        if os.path.exists(self.dll_path):
            try:
                sap_gen = comtypes.client.GetModule(self.dll_path)
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
            if sap_gen is None and os.path.exists(self.dll_path):
                sap_gen = comtypes.client.GetModule(self.dll_path)
            if sap_gen is not None:
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
        try:
            model_filename = self.sap_model.GetModelFilename()
            if isinstance(model_filename, (list, tuple)):
                model_filename = next((v for v in model_filename if isinstance(v, str)), "")
            if not model_filename:
                log.warning("SAP2000 respondió, pero no devolvió una ruta de modelo abierta.")
        except Exception:
            pass

        return self


# =============================================================================
#  BLOQUE 2 — Búsqueda y captura de la ventana de SAP2000
# =============================================================================

class VentanaCaptura:
    """Localiza la ventana de SAP2000 y exporta imágenes del viewport."""

    def __init__(self):
        _ensure_windows_capture_libs()
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

    def _hwnd_en_raiz(self, hwnd: int) -> int:
        """Devuelve la ventana raíz para comparar foco real contra la principal."""
        if not hwnd:
            return 0
        try:
            return win32gui.GetAncestor(hwnd, win32con.GA_ROOT)
        except Exception:
            return hwnd

    def asegurar_enfoque_estricto(self):
        """
        Verifica que SAP2000 tenga el foco antes de enviar automatización.

        Si no logra recuperar el foco confirmado, aborta para no teclear sobre
        otra ventana del usuario.
        """
        if not self.hwnd_principal:
            self.buscar_ventana()

        for _ in range(3):
            self.activar()
            hwnd_foreground = win32gui.GetForegroundWindow()
            if self._hwnd_en_raiz(hwnd_foreground) == self._hwnd_en_raiz(self.hwnd_principal):
                return
            time.sleep(0.25)

        titulo = ""
        try:
            titulo = win32gui.GetWindowText(win32gui.GetForegroundWindow())
        except Exception:
            pass
        raise RuntimeError(
            "Se abortó la automatización porque SAP2000 no quedó en primer plano. "
            f"Ventana activa actual: '{titulo}'."
        )

    def capturar(
        self,
        filepath: str,
        rect_parcial: tuple = None
    ) -> bool:
        """
        Exporta la vista actual de SAP2000 y guarda en el formato solicitado.

        Args:
            filepath:      Ruta completa del archivo de salida.
            rect_parcial:  (izq%, sup%, der%, inf%) en % [0-100]
                           para recortar la imagen exportada.
                           None = exportación completa.
        Returns:
            True si la exportación fue exitosa.
        """
        if not self.hwnd_principal:
            self.buscar_ventana()

        destino = Path(filepath)
        destino.parent.mkdir(parents=True, exist_ok=True)

        bmp_temporal = None
        try:
            if destino.suffix.lower() == ".bmp" and rect_parcial is None:
                bmp_origen = destino
            else:
                with tempfile.NamedTemporaryFile(
                    prefix="sap_capture_",
                    suffix=".bmp",
                    dir=str(destino.parent),
                    delete=False,
                ) as tmp_file:
                    bmp_temporal = Path(tmp_file.name)
                bmp_origen = bmp_temporal

            self._capturar_desde_sap_a_bmp(str(bmp_origen))

            return self._postprocesar_exportacion(
                bmp_path=str(bmp_origen),
                filepath=str(destino),
                rect_parcial=rect_parcial,
            )
        except Exception as e:
            log.error(f"Error al exportar imagen: {e}")
            return False
        finally:
            if bmp_temporal is not None and bmp_temporal.exists():
                try:
                    bmp_temporal.unlink()
                except Exception:
                    pass

    def _capturar_desde_sap_a_bmp(self, bmp_path: str):
        """Ejecuta File > Capture Picture y guarda el BMP mediante el cuadro Save As."""
        pyautogui_local = _ensure_pyautogui()

        self.asegurar_enfoque_estricto()
        self._eliminar_si_existe(bmp_path)

        menu_id = self._buscar_menu_capture_picture()
        if menu_id is None:
            raise RuntimeError("No se encontró el comando 'Capture Picture' en el menú File.")

        win32gui.PostMessage(self.hwnd_principal, win32con.WM_COMMAND, menu_id, 0)

        dialogo = self._esperar_dialogo_guardado()
        self._guardar_dialogo_como(dialogo, bmp_path, pyautogui_local)

        if not self._esperar_archivo(bmp_path, timeout=20.0):
            raise RuntimeError("SAP2000 no generó el BMP dentro del tiempo esperado.")

        log.info(f"  → BMP exportado desde SAP2000: {os.path.basename(bmp_path)}")

    def _buscar_menu_capture_picture(self):
        """Busca el comando Capture Picture dentro del menú File."""
        menu_raiz = win32gui.GetMenu(self.hwnd_principal)
        if not menu_raiz:
            return None

        try:
            menu_file = win32gui.GetSubMenu(menu_raiz, 0)
        except Exception:
            menu_file = None
        if not menu_file:
            return None

        total = win32gui.GetMenuItemCount(menu_file)
        for idx in range(total):
            try:
                texto = win32gui.GetMenuString(menu_file, idx, win32con.MF_BYPOSITION)
            except Exception:
                continue
            normalizado = texto.replace("&", "").strip().lower()
            if "capture picture" in normalizado:
                menu_id = win32gui.GetMenuItemID(menu_file, idx)
                if menu_id != -1:
                    return menu_id
        return None

    def _esperar_dialogo_guardado(self, timeout: float = 8.0) -> int:
        """Espera el cuadro Save As disparado por Capture Picture."""
        deadline = time.time() + timeout
        candidato = 0

        while time.time() < deadline:
            ventanas = []

            def _callback(hwnd, _):
                if not win32gui.IsWindowVisible(hwnd):
                    return
                clase = win32gui.GetClassName(hwnd)
                if clase != "#32770":
                    return
                if self._hwnd_en_raiz(win32gui.GetWindow(hwnd, win32con.GW_OWNER)) != self._hwnd_en_raiz(self.hwnd_principal):
                    return
                ventanas.append(hwnd)

            win32gui.EnumWindows(_callback, None)
            if ventanas:
                candidato = ventanas[0]
                break
            time.sleep(0.2)

        if not candidato:
            raise RuntimeError("No apareció el cuadro de guardado de SAP2000.")

        return candidato

    def _esperar_archivo(self, ruta: str, timeout: float = 15.0) -> bool:
        """Espera a que un archivo exista y deje de crecer."""
        deadline = time.time() + timeout
        ultimo_tamano = -1
        estable_desde = 0.0

        while time.time() < deadline:
            if os.path.exists(ruta):
                try:
                    tamano = os.path.getsize(ruta)
                except OSError:
                    tamano = -1
                if tamano > 0 and tamano == ultimo_tamano:
                    if not estable_desde:
                        estable_desde = time.time()
                    elif time.time() - estable_desde >= 0.4:
                        return True
                else:
                    estable_desde = 0.0
                    ultimo_tamano = tamano
            time.sleep(0.2)
        return False

    def _guardar_dialogo_como(self, dialogo: int, ruta: str, pyautogui_local):
        """Completa el Save As usando controles nativos; si falla, usa teclado."""
        try:
            win32gui.SetForegroundWindow(dialogo)
        except Exception:
            pass
        time.sleep(0.2)

        edit = self._buscar_control_dialogo(dialogo, ("Edit",))
        if edit:
            try:
                win32gui.SetWindowText(edit, ruta)
                win32gui.SendMessage(dialogo, win32con.WM_COMMAND, win32con.IDOK, 0)
                return
            except Exception:
                pass

        pyautogui_local.hotkey("ctrl", "l")
        time.sleep(0.1)
        pyautogui_local.hotkey("ctrl", "a")
        pyautogui_local.write(ruta, interval=0.01)
        time.sleep(0.1)
        pyautogui_local.press("enter")

    def _buscar_control_dialogo(self, hwnd_padre: int, clases_objetivo: tuple[str, ...]) -> int:
        """Busca recursivamente el primer control cuya clase esté en la lista dada."""
        encontrado = 0

        def _callback(hwnd, _):
            nonlocal encontrado
            if encontrado:
                return False
            try:
                clase = win32gui.GetClassName(hwnd)
            except Exception:
                return True
            if clase in clases_objetivo:
                encontrado = hwnd
                return False
            return True

        try:
            win32gui.EnumChildWindows(hwnd_padre, _callback, None)
        except Exception:
            return 0
        return encontrado

    def _postprocesar_exportacion(self, bmp_path: str, filepath: str, rect_parcial: tuple = None) -> bool:
        """Convierte el BMP exportado y aplica recorte si corresponde."""
        image_lib = _ensure_pillow()

        with image_lib.open(bmp_path) as img_bmp:
            img = img_bmp.copy()

        if rect_parcial is not None:
            iz, su, de, in_ = normalizar_crop(rect_parcial)
            iw, ih = img.size
            crop_l = int(iw * iz / 100)
            crop_t = int(ih * su / 100)
            crop_r = int(iw * de / 100)
            crop_b = int(ih * in_ / 100)
            img = img.crop((crop_l, crop_t, crop_r, crop_b))

        extension = Path(filepath).suffix.lower()
        formato = "BMP" if extension == ".bmp" else "PNG"
        img.save(filepath, formato)
        log.info(f"  → Guardado: {os.path.basename(filepath)}  ({img.size[0]}×{img.size[1]} px)")
        return True

    def _eliminar_si_existe(self, ruta: str):
        try:
            os.remove(ruta)
        except FileNotFoundError:
            return


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
        _ensure_pyautogui()
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
        self.ventana.asegurar_enfoque_estricto()

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
        self.ventana.asegurar_enfoque_estricto()
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
            azimut_txt = f"{float(azimut):g}"
            elevacion_txt = f"{float(elevacion):g}"

            # Tab para saltar entre campos del diálogo
            # Campo 1: Plan Rotation (Azimut)
            pyautogui.hotkey("ctrl", "a")      # Seleccionar todo en el campo activo
            pyautogui.write(azimut_txt, interval=0.05)
            time.sleep(0.15)
            pyautogui.press("tab")             # Siguiente campo

            # Campo 2: Elevation
            time.sleep(0.15)
            pyautogui.hotkey("ctrl", "a")
            pyautogui.write(elevacion_txt, interval=0.05)
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
        _ensure_pyautogui()
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
        self.ventana.asegurar_enfoque_estricto()

        # Verificar que el patrón existe
        try:
            response = self.sap_model.LoadPatterns.GetNameList()
            lp_names = []
            if isinstance(response, (list, tuple)):
                for value in response:
                    if isinstance(value, (list, tuple)):
                        lp_names.extend(str(item) for item in value)
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
        self.ventana.asegurar_enfoque_estricto()
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

        tipo_vista   = validar_tipo_vista(config.get("tipo_vista", "ISO_NE"))
        azimut       = float(config.get("azimut",    225))
        elevacion    = float(config.get("elevacion",  30))
        modo_display = config.get("modo_display", DISPLAY_MODELO).upper()
        patron_carga = config.get("patron_carga", "").strip()
        nombre_img   = config.get("nombre_imagen", "imagen").strip()
        ventana_tipo = config.get("ventana", "COMPLETA").upper()
        crop         = normalizar_crop(config.get("crop", None))

        log.info(f"──── Procesando: {nombre_img} ────")

        if modo_display not in (DISPLAY_MODELO, DISPLAY_CARGAS):
            raise ValueError(f"Modo display inválido: '{modo_display}'")
        if ventana_tipo not in ("COMPLETA", "PARCIAL"):
            raise ValueError(f"Tipo de ventana inválido: '{ventana_tipo}'")
        if modo_display == DISPLAY_CARGAS and not patron_carga:
            raise ValueError("Display=CARGAS requiere un patrón de carga.")

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
#  BLOQUE 7 — Integración opcional con workbook abierto (xlwings)
# =============================================================================

def leer_config_desde_excel(wb, allow_unsafe_output: bool = False) -> dict:
    """
    Lee la hoja CONFIG y la tabla de capturas desde un workbook abierto en Excel.

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

    carpeta_excel = os.path.dirname(wb.fullname)
    carpeta_salida = resolver_carpeta_salida(
        carpeta_excel,
        carpeta_rel,
        allow_unsafe=allow_unsafe_output,
    )

    # ── Hoja CAPTURAS ──
    capturas = []
    try:
        sh_cap = wb.sheets["CAPTURAS"]

        # Leer datos desde fila 3 (la 2 es encabezado)
        fila = 3
        while True:
            valores = [sh_cap.range(f"{col}{fila}").value for col in "ABCDEFGHIJKL"]
            if all(valor in (None, "") for valor in valores):
                break
            try:
                capturas.append(cargar_fila_captura(fila, valores))
            except Exception as e:
                log.warning(f"Fila {fila} ignorada por configuración inválida: {e}")
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
    Punto de entrada opcional para ejecutar contra un workbook ya abierto.
    """
    xw_app = _ensure_xlwings()

    try:
        wb = xw_app.Book.caller()
    except Exception:
        # Fallback: usar el primer workbook abierto si no hay caller
        try:
            wb = xw_app.books[0]
        except Exception:
            log.error("No se pudo obtener referencia al workbook de Excel.")
            return

    log.info("=" * 60)
    log.info("SAP2000 Image Capture — iniciando")
    log.info("=" * 60)

    # Leer configuración
    try:
        config = leer_config_desde_excel(
            wb,
            allow_unsafe_output=permitir_salida_insegura(),
        )
    except ValueError as e:
        log.error(str(e))
        return
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
        ruta_excel: Ruta completa del archivo .xlsx a crear.
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
        "3. Ejecuta el script o el EXE portable con este archivo como configuración.",
        "   Ejemplo: sap2000_capture.exe --config SAP2000_Capturas.xlsx",
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
    carpeta_destino = os.path.dirname(os.path.abspath(ruta_excel))
    os.makedirs(carpeta_destino, exist_ok=True)
    wb.save(ruta_excel)
    log.info(f"Excel de configuración creado: {ruta_excel}")


# =============================================================================
#  CLI
# =============================================================================

def main_cli(argv=None) -> int:
    """CLI reutilizable para Python y para el ejecutable portable."""
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
    parser.add_argument(
        "--allow-unsafe-output",
        action="store_true",
        help=(
            "Permitir una carpeta de salida absoluta o fuera de la carpeta base del Excel. "
            "Por defecto solo se aceptan subcarpetas seguras."
        ),
    )
    args = parser.parse_args(argv)

    if args.crear_excel:
        crear_excel_configuracion(args.crear_excel)
        print(f"Excel creado en: {args.crear_excel}")
        return 0

    if args.config:
        try:
            # Leer configuración desde Excel y ejecutar sin xlwings caller
            import openpyxl as _opx
            _wb_raw = _opx.load_workbook(args.config, data_only=True)

            # Construir config manualmente desde openpyxl
            sh = _wb_raw["CONFIG"]
            sap_dll       = sh["B2"].value or SAP_DLL_PATH
            nombre_proy   = sh["B3"].value or args.proyecto
            carpeta_rel   = sh["B4"].value or "Capturas_SAP"
            carpeta_salida = resolver_carpeta_salida(
                os.path.dirname(args.config),
                carpeta_rel,
                allow_unsafe=permitir_salida_insegura(args.allow_unsafe_output),
            )

            sh_cap = _wb_raw["CAPTURAS"]
            capturas = []
            for fila in range(3, sh_cap.max_row + 1):
                vals = [sh_cap.cell(fila, c).value for c in range(1, 13)]
                if all(valor in (None, "") for valor in vals):
                    break
                try:
                    capturas.append(cargar_fila_captura(fila, vals))
                except Exception as e:
                    log.warning(f"Fila {fila} ignorada por configuración inválida: {e}")

            conector = SAP2000Conector(sap_dll)
            conector.conectar()
            gen = GeneradorImagenes(conector.sap_model, carpeta_salida, nombre_proy)
            gen.inicializar()
            resultados = gen.procesar_lista(capturas)

            ok  = sum(1 for r in resultados if r["ok"])
            err = len(resultados) - ok
            print(f"\nResultado: {ok} imágenes OK, {err} errores → {carpeta_salida}")
            return 0 if err == 0 else 1
        except ValueError as e:
            log.error(str(e))
            return 2

    # Sin argumentos: si existe la plantilla en la carpeta del ejecutable o actual,
    # asumir que el usuario quiere ejecutar capturas con doble clic.
    rutas_candidatas = []
    if getattr(sys, "frozen", False):
        rutas_candidatas.append(os.path.join(os.path.dirname(sys.executable), "SAP2000_Capturas.xlsx"))
    rutas_candidatas.append(os.path.join(os.getcwd(), "SAP2000_Capturas.xlsx"))

    for ruta_config in rutas_candidatas:
        if os.path.exists(ruta_config):
            print(f"Usando configuración detectada: {ruta_config}")
            return main_cli(["--config", ruta_config])

    # Sin argumentos ni plantilla: crear el Excel en la carpeta del ejecutable
    # si estamos dentro del EXE, o en la carpeta actual si corremos con Python.
    if getattr(sys, "frozen", False):
        ruta = os.path.join(os.path.dirname(sys.executable), "SAP2000_Capturas.xlsx")
    else:
        ruta = os.path.join(os.getcwd(), "SAP2000_Capturas.xlsx")
    crear_excel_configuracion(ruta)
    print(f"Plantilla Excel creada en: {ruta}")
    print("Úsala con: python sap_imagenes.py --config SAP2000_Capturas.xlsx")
    return 0


# =============================================================================
#  Punto de entrada directo (python sap_imagenes.py)
# =============================================================================

if __name__ == "__main__":
    sys.exit(main_cli())
