"""
win32_capture.py
Motor de captura de ventanas usando el contenido visible en pantalla.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import logging
import os
import time
from pathlib import Path

try:
    import win32gui
    import win32process
    import win32ui
except ImportError:
    win32gui = None
    win32process = None
    win32ui = None

try:
    from PIL import Image
except ImportError:
    Image = None

logger = logging.getLogger(__name__)
SRCCOPY = 0x00CC0020


def _set_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


_set_dpi_awareness()


def _require_win32() -> None:
    if win32gui is None or win32process is None or win32ui is None:
        raise RuntimeError("pywin32 no esta disponible en este entorno.")
    if Image is None:
        raise RuntimeError("Pillow no esta disponible en este entorno.")


def _normalize_title(text: str) -> str:
    return " ".join(text.lower().split())


def _score_sap2000_window(title: str, class_name: str, pid: int, current_pid: int) -> int | None:
    normalized = _normalize_title(title)
    if not normalized:
        return None
    if pid == current_pid:
        return None
    if "sap2000 capture" in normalized:
        return None
    if "capture" in normalized and "sap2000" in normalized:
        return None
    if "sap2000" not in normalized and "sap 2000" not in normalized and "csi.sap2000" not in normalized:
        return None

    score = 0
    if normalized.startswith("sap2000"):
        score += 5
    if normalized.startswith("sap 2000"):
        score += 5
    if "csi.sap2000" in normalized:
        score += 4
    if "computers and structures" in normalized:
        score += 2
    if class_name and "tk" in class_name.lower():
        score -= 10
    return score


def _list_window_details() -> list[tuple[int, str, str, int]]:
    _require_win32()
    results: list[tuple[int, str, str, int]] = []

    def _callback(hwnd, _lparam):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        class_name = win32gui.GetClassName(hwnd)
        _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
        results.append((hwnd, title, class_name, pid))

    win32gui.EnumWindows(_callback, None)
    return results


def list_windows(title_filter: str = "") -> list[tuple[int, str]]:
    results: list[tuple[int, str]] = []
    for hwnd, title, _class_name, _pid in _list_window_details():
        if title_filter.lower() in title.lower():
            results.append((hwnd, title))
    return results


def find_window_by_partial_title(partial: str) -> int | None:
    matches = list_windows(partial)
    return matches[0][0] if matches else None


def find_sap2000_hwnd() -> int | None:
    current_pid = os.getpid()
    ranked: list[tuple[int, int, str, str, int]] = []
    for hwnd, title, class_name, pid in _list_window_details():
        score = _score_sap2000_window(title, class_name, pid, current_pid)
        if score is None:
            continue
        ranked.append((score, hwnd, title, class_name, pid))

    if ranked:
        ranked.sort(key=lambda item: item[0], reverse=True)
        score, hwnd, title, class_name, pid = ranked[0]
        logger.info(
            "Ventana SAP2000 seleccionada: hwnd=%s pid=%s class='%s' title='%s' score=%s",
            hwnd,
            pid,
            class_name,
            title,
            score,
        )
        return hwnd
    return None


def prepare_window_for_capture(hwnd: int) -> None:
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, 9)
        else:
            win32gui.ShowWindow(hwnd, 5)
        win32gui.BringWindowToTop(hwnd)
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception as exc:
            logger.debug("No se pudo activar hwnd=%s: %r", hwnd, exc)
    except Exception as exc:
        logger.warning("No se pudo preparar hwnd=%s para captura: %r", hwnd, exc)


def get_window_rect(hwnd: int) -> tuple[int, int, int, int]:
    _require_win32()
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    return left, top, right - left, bottom - top


def get_client_rect_on_window(hwnd: int) -> tuple[int, int, int, int]:
    _require_win32()
    left_c, top_c, right_c, bottom_c = win32gui.GetClientRect(hwnd)
    width = right_c - left_c
    height = bottom_c - top_c

    origin = ctypes.wintypes.POINT(left_c, top_c)
    ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(origin))

    win_left, win_top, _, _ = get_window_rect(hwnd)
    return origin.x - win_left, origin.y - win_top, width, height


class Win32CaptureEngine:
    def __init__(self, render_delay: float = 0.5):
        self.render_delay = render_delay

    def wait_render(self) -> None:
        time.sleep(self.render_delay)

    def capture(
        self,
        hwnd: int,
        output_path: str | Path,
        use_client_area: bool = True,
    ) -> Path:
        _require_win32()
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        if use_client_area:
            left, top, width, height = self._get_client_screen_rect(hwnd)
        else:
            left, top, width, height = self._get_window_screen_rect(hwnd)

        image = self._capture_screen_rect(left, top, width, height)
        image.save(str(output_path), format="PNG", optimize=True)
        return output_path

    def capture_after_render(
        self,
        hwnd: int,
        output_path: str | Path,
        use_client_area: bool = True,
    ) -> Path:
        self.wait_render()
        return self.capture(hwnd, output_path, use_client_area)

    def _get_window_screen_rect(self, hwnd: int) -> tuple[int, int, int, int]:
        left, top, width, height = get_window_rect(hwnd)
        if width <= 0 or height <= 0:
            raise RuntimeError(
                f"Ventana hwnd={hwnd} tiene dimensiones invalidas ({width}x{height})"
            )
        return left, top, width, height

    def _get_client_screen_rect(self, hwnd: int) -> tuple[int, int, int, int]:
        win_left, win_top, _win_width, _win_height = get_window_rect(hwnd)
        client_left, client_top, client_width, client_height = get_client_rect_on_window(hwnd)
        if client_width <= 0 or client_height <= 0:
            raise RuntimeError(
                f"Area cliente invalida para hwnd={hwnd} ({client_width}x{client_height})"
            )
        return win_left + client_left, win_top + client_top, client_width, client_height

    def _capture_screen_rect(self, left: int, top: int, width: int, height: int) -> Image.Image:
        screen_dc = win32gui.GetDC(0)
        src_dc = win32ui.CreateDCFromHandle(screen_dc)
        mem_dc = src_dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(src_dc, width, height)
        mem_dc.SelectObject(bmp)

        try:
            mem_dc.BitBlt((0, 0), (width, height), src_dc, (left, top), SRCCOPY)
            bmpinfo = bmp.GetInfo()
            bmpbits = bmp.GetBitmapBits(True)
            return Image.frombuffer(
                "RGB",
                (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                bmpbits,
                "raw",
                "BGRX",
                0,
                1,
            )
        finally:
            mem_dc.DeleteDC()
            src_dc.DeleteDC()
            win32gui.ReleaseDC(0, screen_dc)
            win32gui.DeleteObject(bmp.GetHandle())


if __name__ == "__main__":
    import sys

    print("Ventanas visibles:")
    for hwnd, title in list_windows():
        print(f"  hwnd={hwnd:8d}  '{title}'")

    hwnd = find_sap2000_hwnd()
    if hwnd is None:
        print("\nNo se encontro SAP2000 abierto.")
        sys.exit(1)

    print(f"\nSAP2000 encontrado: hwnd={hwnd}")
    engine = Win32CaptureEngine(render_delay=0.3)
    out = engine.capture(hwnd, "test_capture.png")
    print(f"Captura guardada en: {out}")
