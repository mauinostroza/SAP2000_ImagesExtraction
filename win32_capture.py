"""
win32_capture.py
Motor de captura de ventanas usando PrintWindow (Win32).
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import time
from pathlib import Path

import win32gui
import win32ui
from PIL import Image


def _set_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


_set_dpi_awareness()


def list_windows(title_filter: str = "") -> list[tuple[int, str]]:
    results: list[tuple[int, str]] = []

    def _callback(hwnd, _lparam):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title_filter.lower() in title.lower():
                results.append((hwnd, title))

    win32gui.EnumWindows(_callback, None)
    return results


def find_window_by_partial_title(partial: str) -> int | None:
    matches = list_windows(partial)
    return matches[0][0] if matches else None


def find_sap2000_hwnd() -> int | None:
    for candidate in ("SAP2000", "CSI.SAP2000", "SAP 2000"):
        hwnd = find_window_by_partial_title(candidate)
        if hwnd:
            return hwnd
    return None


def get_window_rect(hwnd: int) -> tuple[int, int, int, int]:
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    return left, top, right - left, bottom - top


def get_client_rect_on_window(hwnd: int) -> tuple[int, int, int, int]:
    left_c, top_c, right_c, bottom_c = win32gui.GetClientRect(hwnd)
    width = right_c - left_c
    height = bottom_c - top_c

    origin = ctypes.wintypes.POINT(left_c, top_c)
    ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(origin))

    win_left, win_top, _, _ = get_window_rect(hwnd)
    return origin.x - win_left, origin.y - win_top, width, height


class Win32CaptureEngine:
    PW_RENDERFULLCONTENT = 3

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
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        win_left, win_top, width, height = get_window_rect(hwnd)
        if width <= 0 or height <= 0:
            raise RuntimeError(
                f"Ventana hwnd={hwnd} tiene dimensiones inválidas ({width}x{height})"
            )

        hwnd_dc = win32gui.GetWindowDC(hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mfc_dc, width, height)
        save_dc.SelectObject(bmp)

        try:
            result = ctypes.windll.user32.PrintWindow(
                hwnd,
                save_dc.GetSafeHdc(),
                self.PW_RENDERFULLCONTENT,
            )
            if result == 0:
                result = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 0)
            if result == 0:
                raise RuntimeError(
                    f"PrintWindow falló para hwnd={hwnd}. "
                    "Verifica que SAP2000 no esté minimizado."
                )

            bmpinfo = bmp.GetInfo()
            bmpbits = bmp.GetBitmapBits(True)
            image = Image.frombuffer(
                "RGB",
                (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                bmpbits,
                "raw",
                "BGRX",
                0,
                1,
            )

            if use_client_area:
                client_left, client_top, client_width, client_height = get_client_rect_on_window(hwnd)
                image = image.crop(
                    (
                        client_left,
                        client_top,
                        client_left + client_width,
                        client_top + client_height,
                    )
                )

            image.save(str(output_path), format="PNG", optimize=True)
            return output_path
        finally:
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
        self.wait_render()
        return self.capture(hwnd, output_path, use_client_area)


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
