"""
win32_capture.py
Motor de captura de ventanas usando PrintWindow (Win32).

Ventajas sobre CopyFromScreen:
  - Captura aunque SAP2000 esté detrás de otra ventana
  - No requiere que el usuario no toque el teclado
  - Funciona aunque SAP esté parcialmente fuera del monitor

Dependencias: pywin32, Pillow
  pip install pywin32 Pillow

Uso:
  engine = Win32CaptureEngine()
  hwnd   = engine.find_sap2000()
  engine.capture(hwnd, "vista_3d.png")
"""

import ctypes
import ctypes.wintypes
import time
from pathlib import Path

import win32con
import win32gui
import win32ui
from PIL import Image

# ── DPI awareness ──────────────────────────────────────────────────────────────
# Debe llamarse UNA VEZ antes de cualquier operación de ventana.
# Sin esto, GetWindowRect devuelve coordenadas escaladas y la captura queda
# recortada en monitores con escala > 100%.
def _set_dpi_awareness():
    try:
        # Per-Monitor DPI Aware v2 (Windows 10 1703+)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            # Fallback: System DPI Aware
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

_set_dpi_awareness()


# ── Utilidades de ventana ──────────────────────────────────────────────────────

def list_windows(title_filter: str = "") -> list[tuple[int, str]]:
    """Devuelve (hwnd, title) de todas las ventanas visibles."""
    results = []
    def _callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title_filter.lower() in title.lower():
                results.append((hwnd, title))
    win32gui.EnumWindows(_callback, None)
    return results


def find_window_by_partial_title(partial: str) -> int | None:
    """Busca la primera ventana cuyo título contenga 'partial' (case-insensitive).
    Retorna el hwnd o None si no se encuentra.
    """
    matches = list_windows(partial)
    if matches:
        return matches[0][0]
    return None


def find_sap2000_hwnd() -> int | None:
    """Localiza la ventana principal de SAP2000.
    Prueba varios títulos conocidos en orden de prioridad.
    """
    candidates = ["SAP2000", "CSI.SAP2000", "SAP 2000"]
    for candidate in candidates:
        hwnd = find_window_by_partial_title(candidate)
        if hwnd:
            return hwnd
    return None


def get_client_rect_screen(hwnd: int) -> tuple[int, int, int, int]:
    """Retorna (left, top, width, height) del área cliente en coordenadas de pantalla.
    Usa GetClientRect + ClientToScreen para evitar incluir la barra de título.
    """
    # Área cliente relativa a la ventana
    left_c, top_c, right_c, bottom_c = win32gui.GetClientRect(hwnd)
    width  = right_c - left_c
    height = bottom_c - top_c

    # Convertir origen a coordenadas de pantalla
    pt = ctypes.wintypes.POINT(left_c, top_c)
    ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(pt))

    return pt.x, pt.y, width, height


# ── Motor de captura ───────────────────────────────────────────────────────────

class Win32CaptureEngine:
    """Captura ventanas Windows usando PrintWindow.

    PrintWindow copia el contenido de la ventana directamente desde el renderer
    del proceso destino, sin depender de lo que hay visible en pantalla.

    Parámetro PW_RENDERFULLCONTENT (3): captura el contenido completo incluyendo
    elementos que usen DirectComposition, lo cual es necesario para SAP2000 v23+.
    """

    PW_RENDERFULLCONTENT = 3  # Documentado en WinUser.h desde Windows 8.1

    def __init__(self, render_delay: float = 0.5):
        """
        Args:
            render_delay: segundos a esperar tras un cambio de vista antes de capturar.
                          SAP2000 necesita ~300-600 ms para redibujar.
        """
        self.render_delay = render_delay

    def wait_render(self):
        time.sleep(self.render_delay)

    def capture(
        self,
        hwnd: int,
        output_path: str | Path,
        use_client_area: bool = True,
    ) -> Path:
        """Captura la ventana 'hwnd' y guarda como PNG en 'output_path'.

        Args:
            hwnd:            Handle de la ventana a capturar.
            output_path:     Ruta del archivo PNG de salida.
            use_client_area: Si True, captura solo el área cliente (sin barra de título).
                             Si False, captura la ventana completa.

        Returns:
            Path del archivo guardado.

        Raises:
            RuntimeError: Si PrintWindow falla o la ventana es inválida.
        """
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Obtener dimensiones
        if use_client_area:
            x, y, width, height = get_client_rect_screen(hwnd)
        else:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            x, y = left, top
            width, height = right - left, bottom - top

        if width <= 0 or height <= 0:
            raise RuntimeError(f"Ventana hwnd={hwnd} tiene dimensiones inválidas ({width}×{height})")

        # Crear Device Context compatible
        hwnd_dc  = win32gui.GetWindowDC(hwnd)
        mfc_dc   = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc  = mfc_dc.CreateCompatibleDC()

        # Crear bitmap destino
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mfc_dc, width, height)
        save_dc.SelectObject(bmp)

        try:
            # PrintWindow: copia el contenido del renderer al bitmap
            # Llamamos directo a user32 para poder pasar PW_RENDERFULLCONTENT=3
            result = ctypes.windll.user32.PrintWindow(
                hwnd,
                save_dc.GetSafeHdc(),
                self.PW_RENDERFULLCONTENT,
            )
            if result == 0:
                raise RuntimeError(
                    f"PrintWindow falló para hwnd={hwnd}. "
                    "Verifica que SAP2000 no esté minimizado."
                )

            # Convertir bitmap Win32 → imagen PIL
            bmpinfo = bmp.GetInfo()
            bmpbits = bmp.GetBitmapBits(True)  # True = datos como bytes

            img = Image.frombuffer(
                "RGB",
                (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                bmpbits,
                "raw",
                "BGRX",   # Win32 usa BGR con byte de relleno X
                0,
                1,
            )

            # Guardar PNG
            img.save(str(output_path), format="PNG", optimize=True)
            return output_path

        finally:
            # Liberar recursos GDI (crítico: fugas de DC crashean el sistema)
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwnd_dc)
            win32gui.DeleteObject(bmp.GetHandle())

    def capture_after_render(
        self,
        hwnd: int,
        output_path: str | Path,
        use_client_area: bool = True,
    ) -> Path:
        """Espera el render_delay y luego captura. Atajo para el flujo habitual."""
        self.wait_render()
        return self.capture(hwnd, output_path, use_client_area)


# ── CLI de diagnóstico ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys

    print("Ventanas visibles:")
    for hwnd, title in list_windows():
        print(f"  hwnd={hwnd:8d}  '{title}'")

    hwnd = find_sap2000_hwnd()
    if hwnd is None:
        print("\nNo se encontró SAP2000 abierto.")
        sys.exit(1)

    print(f"\nSAP2000 encontrado: hwnd={hwnd}")
    engine = Win32CaptureEngine(render_delay=0.3)
    out = engine.capture(hwnd, "test_capture.png")
    print(f"Captura guardada en: {out}")
