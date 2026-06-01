"""
sap_ui_automation.py
Automatizacion UI ligera para SAP2000 usando teclado sobre la ventana activa.
"""

from __future__ import annotations

import ctypes
import logging
import time
from enum import IntEnum
from typing import Callable

logger = logging.getLogger(__name__)

user32 = ctypes.windll.user32

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
VK_SHIFT = 0x10
MF_BYPOSITION = 0x400
WM_COMMAND = 0x0111
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
MK_LBUTTON = 0x0001
GA_ROOT = 2
GA_ROOTOWNER = 3

VK_MAP = {
    "alt": 0x12,
    "ctrl": 0x11,
    "enter": 0x0D,
    "tab": 0x09,
    "escape": 0x1B,
    "down": 0x28,
    "right": 0x27,
    "up": 0x26,
    "left": 0x25,
}


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class _INPUTUNION(ctypes.Union):
    _fields_ = [("ki", _KEYBDINPUT)]


class _INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("union", _INPUTUNION)]


class SapUIView(IntEnum):
    ISO_3D = 0
    PLAN_XY = 1
    ELEV_XZ = 2
    ELEV_YZ = 3


DEFAULT_VIEW_ANGLES = {
    SapUIView.ISO_3D: (225, 30),
}


class UIAutomationAborted(RuntimeError):
    pass


def _normalize_menu_text(text: str) -> str:
    return text.replace("&", "").split("\t", 1)[0].strip().lower()


def _make_lparam(low: int, high: int) -> int:
    return ((high & 0xFFFF) << 16) | (low & 0xFFFF)


def _vk_for_key(key: str) -> int:
    lowered = key.lower()
    if lowered in VK_MAP:
        return VK_MAP[lowered]
    if len(lowered) == 1:
        code = user32.VkKeyScanW(ord(lowered))
        if code == -1:
            raise ValueError(f"Tecla no soportada: {key}")
        return code & 0xFF
    raise ValueError(f"Tecla no soportada: {key}")


def _send_input(*inputs: _INPUT) -> None:
    array_type = _INPUT * len(inputs)
    sent = user32.SendInput(len(inputs), array_type(*inputs), ctypes.sizeof(_INPUT))
    if sent != len(inputs):
        raise RuntimeError(f"SendInput envio {sent}/{len(inputs)} eventos")


def _key_event(vk: int, keyup: bool = False) -> _INPUT:
    flags = KEYEVENTF_KEYUP if keyup else 0
    return _INPUT(type=INPUT_KEYBOARD, union=_INPUTUNION(ki=_KEYBDINPUT(vk, 0, flags, 0, None)))


def _unicode_event(char: str, keyup: bool = False) -> _INPUT:
    flags = KEYEVENTF_UNICODE | (KEYEVENTF_KEYUP if keyup else 0)
    return _INPUT(
        type=INPUT_KEYBOARD,
        union=_INPUTUNION(ki=_KEYBDINPUT(0, ord(char), flags, 0, None)),
    )


def _keybd_event(vk: int, keyup: bool = False) -> None:
    flags = KEYEVENTF_KEYUP if keyup else 0
    user32.keybd_event(vk, 0, flags, 0)


def _send_vk_press(vk: int) -> None:
    try:
        _send_input(_key_event(vk), _key_event(vk, keyup=True))
    except Exception:
        _keybd_event(vk)
        time.sleep(0.02)
        _keybd_event(vk, keyup=True)


def _send_hotkey(vks: list[int]) -> None:
    try:
        downs = [_key_event(vk) for vk in vks]
        ups = [_key_event(vk, keyup=True) for vk in reversed(vks)]
        _send_input(*(downs + ups))
        return
    except Exception:
        pass
    for vk in vks:
        _keybd_event(vk)
        time.sleep(0.02)
    for vk in reversed(vks):
        _keybd_event(vk, keyup=True)
        time.sleep(0.02)


def _vk_combo_for_char(char: str) -> tuple[list[int], int] | None:
    code = user32.VkKeyScanW(ord(char))
    if code == -1:
        return None
    vk = code & 0xFF
    shift_state = (code >> 8) & 0xFF
    modifiers: list[int] = []
    if shift_state & 0x01:
        modifiers.append(VK_SHIFT)
    if shift_state & 0x02:
        modifiers.append(VK_MAP["ctrl"])
    if shift_state & 0x04:
        modifiers.append(VK_MAP["alt"])
    return modifiers, vk


def press_key(key: str) -> None:
    vk = _vk_for_key(key)
    _send_vk_press(vk)


def hotkey(*keys: str) -> None:
    vks = [_vk_for_key(key) for key in keys]
    _send_hotkey(vks)


def write_text(text: str, interval: float = 0.04) -> None:
    for char in text:
        combo = _vk_combo_for_char(char)
        if combo is None:
            _send_input(_unicode_event(char), _unicode_event(char, keyup=True))
        else:
            modifiers, vk = combo
            _send_hotkey([*modifiers, vk])
        if interval > 0:
            time.sleep(interval)


class SAP2000UIController:
    def __init__(
        self,
        hwnd: int,
        enabled: bool = False,
        stop_requested: Callable[[], bool] | object | None = None,
    ):
        self.hwnd_principal = hwnd
        self.enabled = enabled
        if stop_requested is None:
            self.stop_requested = lambda: False
        elif callable(stop_requested):
            self.stop_requested = stop_requested
        elif hasattr(stop_requested, "is_set"):
            self.stop_requested = stop_requested.is_set
        else:
            raise TypeError("stop_requested debe ser callable, Event o None")
        self.sap_pid = self._get_window_pid(hwnd)
        self._menu_dumped = False
        self._sent_input = False

    def _is_escape_pressed(self) -> bool:
        return bool(user32.GetAsyncKeyState(VK_MAP["escape"]) & 0x8000)

    def _get_foreground_hwnd(self) -> int:
        return int(user32.GetForegroundWindow())

    def _get_window_pid(self, hwnd: int) -> int:
        pid = ctypes.c_ulong()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return int(pid.value)

    def _foreground_belongs_to_sap(self) -> bool:
        hwnd = self._get_foreground_hwnd()
        if hwnd == 0:
            return False
        return self._window_is_sap_context(hwnd)

    def _get_ancestor(self, hwnd: int, flag: int) -> int:
        return int(user32.GetAncestor(hwnd, flag))

    def _window_is_sap_context(self, hwnd: int) -> bool:
        if hwnd == 0:
            return False
        if self._get_window_pid(hwnd) != self.sap_pid:
            return False
        if hwnd == self.hwnd_principal:
            return True
        root = self._get_ancestor(hwnd, GA_ROOT)
        root_owner = self._get_ancestor(hwnd, GA_ROOTOWNER)
        return self.hwnd_principal in {root, root_owner}

    def _get_window_text(self, hwnd: int) -> str:
        length = int(user32.GetWindowTextLengthW(hwnd))
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, length + 1)
        return buffer.value

    def _get_class_name(self, hwnd: int) -> str:
        buffer = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, buffer, 256)
        return buffer.value

    def _iter_child_windows(self, parent_hwnd: int, max_items: int = 200):
        items: list[int] = []
        enum_proc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def callback(hwnd, _lparam):
            items.append(int(hwnd))
            return len(items) < max_items

        user32.EnumChildWindows(parent_hwnd, enum_proc(callback), 0)
        return items

    def _find_child_by_title(self, title: str) -> int | None:
        needle = _normalize_menu_text(title)
        for hwnd in self._iter_child_windows(self.hwnd_principal):
            current = _normalize_menu_text(self._get_window_text(hwnd))
            if current == needle:
                return hwnd
        return None

    def _click_child_center(self, child_hwnd: int, pause: float = 0.4) -> None:
        rect = ctypes.wintypes.RECT()
        if not user32.GetClientRect(child_hwnd, ctypes.byref(rect)):
            raise RuntimeError(f"No se pudo leer client rect de hwnd={child_hwnd}")
        x = max(1, (rect.right - rect.left) // 2)
        y = max(1, (rect.bottom - rect.top) // 2)
        lparam = _make_lparam(x, y)
        self._run_guarded(
            f"click child hwnd={child_hwnd}",
            lambda: (
                user32.PostMessageW(child_hwnd, WM_LBUTTONDOWN, MK_LBUTTON, lparam),
                user32.PostMessageW(child_hwnd, WM_LBUTTONUP, 0, lparam),
            ),
            pause=pause,
        )

    def _assert_safe_context(self, step: str) -> None:
        if not self.enabled:
            raise UIAutomationAborted("automatizacion UI no armada")
        if self.stop_requested():
            raise UIAutomationAborted(f"automatizacion UI detenida antes de {step}")
        if self._is_escape_pressed():
            raise UIAutomationAborted(f"tecla Esc detectada antes de {step}")
        if not self._foreground_belongs_to_sap():
            raise UIAutomationAborted(f"foco fuera de SAP2000 antes de {step}")

    def _cancel_if_possible(self) -> None:
        if not self._sent_input:
            logger.info("UI abortada antes de enviar teclas; no se enviara Escape de limpieza.")
            return
        try:
            if self._foreground_belongs_to_sap():
                press_key("escape")
        except Exception:
            pass

    def _get_menu_handle(self):
        return user32.GetMenu(self.hwnd_principal)

    def _iter_menu_items(self, menu_handle, prefix: tuple[str, ...] = (), depth: int = 0, max_depth: int = 3):
        if not menu_handle:
            return
        count = int(user32.GetMenuItemCount(menu_handle))
        for index in range(max(count, 0)):
            text_buf = ctypes.create_unicode_buffer(256)
            user32.GetMenuStringW(menu_handle, index, text_buf, 256, MF_BYPOSITION)
            raw_text = text_buf.value
            submenu = user32.GetSubMenu(menu_handle, index)
            item_id = int(user32.GetMenuItemID(menu_handle, index))
            path = (*prefix, raw_text)
            yield {"path": path, "text": raw_text, "id": item_id, "submenu": submenu, "depth": depth}
            if submenu and depth + 1 <= max_depth:
                yield from self._iter_menu_items(submenu, path, depth + 1, max_depth)

    def _log_menu_snapshot(self, reason: str) -> None:
        if self._menu_dumped:
            return
        self._menu_dumped = True
        menu_handle = self._get_menu_handle()
        if not menu_handle:
            logger.warning("SAP2000 no expone menu Win32 para diagnostico (%s).", reason)
            self._log_window_context(reason)
            self._log_child_windows_snapshot(reason)
            return
        logger.warning("Volcado de menu SAP2000 por %s:", reason)
        for item in self._iter_menu_items(menu_handle, max_depth=2):
            text = " > ".join(part for part in item["path"] if part)
            logger.warning("  menu depth=%s id=%s path=%s", item["depth"], item["id"], text)

    def _find_menu_command_id(self, *needles: str) -> int | None:
        menu_handle = self._get_menu_handle()
        if not menu_handle:
            return None
        normalized_needles = [needle.strip().lower() for needle in needles if needle.strip()]
        for item in self._iter_menu_items(menu_handle, max_depth=4):
            if item["submenu"]:
                continue
            haystack = " > ".join(_normalize_menu_text(part) for part in item["path"] if part)
            if all(needle in haystack for needle in normalized_needles):
                if item["id"] != -1:
                    return item["id"]
        return None

    def _invoke_menu_command(self, command_id: int, pause: float = 0.5) -> None:
        self._run_guarded(
            f"invocar comando menu {command_id}",
            lambda: user32.PostMessageW(self.hwnd_principal, WM_COMMAND, command_id, 0),
            pause=pause,
        )

    def _run_guarded(self, step: str, action, pause: float = 0.0) -> None:
        self._assert_safe_context(step)
        self._sent_input = True
        action()
        if pause > 0:
            time.sleep(pause)
        self._assert_safe_context(step)

    def _log_window_context(self, reason: str) -> None:
        hwnd = self._get_foreground_hwnd()
        if hwnd == 0:
            logger.warning("No hay ventana foreground durante %s.", reason)
            return
        logger.warning(
            "Foreground durante %s: hwnd=%s pid=%s class=%r title=%r root=%s root_owner=%s",
            reason,
            hwnd,
            self._get_window_pid(hwnd),
            self._get_class_name(hwnd),
            self._get_window_text(hwnd),
            self._get_ancestor(hwnd, GA_ROOT),
            self._get_ancestor(hwnd, GA_ROOTOWNER),
        )

    def _log_child_windows_snapshot(self, reason: str, max_items: int = 80) -> None:
        logger.warning("Volcado de ventanas hijas SAP2000 por %s:", reason)
        items = [
            (
                hwnd,
                self._get_window_pid(hwnd),
                self._get_class_name(hwnd),
                self._get_window_text(hwnd),
            )
            for hwnd in self._iter_child_windows(self.hwnd_principal, max_items=max_items)
        ]
        if not items:
            logger.warning("  sin ventanas hijas enumerables")
            return
        for hwnd, pid, class_name, title in items:
            logger.warning("  child hwnd=%s pid=%s class=%r title=%r", hwnd, pid, class_name, title)

    def _log_process_windows_snapshot(self, reason: str, max_items: int = 80) -> None:
        logger.warning("Volcado de ventanas top-level SAP2000 por %s:", reason)
        items: list[tuple[int, int, str, str]] = []
        enum_proc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def callback(hwnd, _lparam):
            hwnd = int(hwnd)
            if self._get_window_pid(hwnd) != self.sap_pid:
                return True
            items.append((hwnd, self._get_window_pid(hwnd), self._get_class_name(hwnd), self._get_window_text(hwnd)))
            return len(items) < max_items

        user32.EnumWindows(enum_proc(callback), 0)
        if not items:
            logger.warning("  sin ventanas top-level del proceso")
            return
        for hwnd, pid, class_name, title in items:
            logger.warning("  top hwnd=%s pid=%s class=%r title=%r", hwnd, pid, class_name, title)

    def activar(self) -> None:
        user32.ShowWindow(self.hwnd_principal, 5)
        user32.SetForegroundWindow(self.hwnd_principal)
        time.sleep(0.25)
        try:
            self._assert_safe_context("activar ventana")
        except UIAutomationAborted:
            self._log_window_context("activar ventana")
            raise

    def set_vista(
        self,
        view_type: SapUIView,
        azimuth: float | None = None,
        elevation: float | None = None,
    ) -> bool:
        if not self.enabled:
            logger.info("Automatizacion UI desarmada; no se enviaran teclas para vista.")
            return False
        if view_type not in DEFAULT_VIEW_ANGLES:
            logger.info("UI vista no implementada para %s; se conserva la actual.", view_type.name)
            return False

        azimut, elevacion = DEFAULT_VIEW_ANGLES[view_type]
        if azimuth is not None:
            azimut = azimuth
        if elevation is not None:
            elevacion = elevation
        logger.info("UI vista %s -> az=%s el=%s", view_type.name, azimut, elevacion)
        self._sent_input = False
        self.activar()
        try:
            self._abrir_dialogo_set3dview()
            self._ingresar_angulos_en_dialogo(azimut, elevacion)
            time.sleep(0.3)
            self._assert_safe_context("fin set_vista")
            return True
        except UIAutomationAborted as exc:
            logger.warning("Automatizacion UI abortada en vista: %s", exc)
            self._log_window_context("aborto vista")
            self._cancel_if_possible()
            return False

    def mostrar_cargas_patron(self, patron: str) -> bool:
        if not self.enabled:
            logger.info("Automatizacion UI desarmada; no se enviaran teclas para cargas.")
            return False
        try:
            self._sent_input = False
            self.activar()
            command_id = (
                self._find_menu_command_id("show load assigns")
                or self._find_menu_command_id("load assigns", "show")
            )
            if command_id is None:
                self._log_menu_snapshot("no se encontro Show Load Assigns")
                raise UIAutomationAborted("no se encontro comando de menu Show Load Assigns")
            logger.info("UI comando Show Load Assigns encontrado: id=%s", command_id)
            self._invoke_menu_command(command_id, pause=0.7)
            self._run_guarded("seleccionar patron", lambda: hotkey("ctrl", "a"), pause=0.05)
            self._run_guarded("escribir patron", lambda: write_text(patron, interval=0.04), pause=0.2)
            self._run_guarded("confirmar patron", lambda: press_key("enter"), pause=0.4)
            return True
        except UIAutomationAborted as exc:
            logger.warning("Automatizacion UI abortada en cargas: %s", exc)
            self._log_window_context("aborto cargas")
            self._cancel_if_possible()
            return False
        except Exception as exc:
            logger.warning("No se pudo seleccionar patron en dialogo: %r", exc)
            self._cancel_if_possible()
            return False

    def set_extrusion(self, enabled: bool) -> bool:
        if not enabled:
            logger.info("Vista extruida desactivada; se conserva el estado visual actual.")
            return False
        if not self.enabled:
            logger.info("Automatizacion UI desarmada; no se enviaran teclas para extrusion.")
            return False
        logger.info("UI extrusion solicitada, pero aun no esta implementada.")
        return False

    def _abrir_dialogo_set3dview(self) -> None:
        command_id = (
            self._find_menu_command_id("set 3d view")
            or self._find_menu_command_id("3d view", "set")
            or self._find_menu_command_id("rotate 3d", "set")
        )
        if command_id is None:
            self._log_menu_snapshot("no se encontro Set 3D View")
            self._probe_view_control()
            raise UIAutomationAborted("no se encontro comando de menu Set 3D View")
        logger.info("UI comando Set 3D View encontrado: id=%s", command_id)
        self._invoke_menu_command(command_id, pause=0.7)

    def _probe_view_control(self) -> None:
        child_hwnd = self._find_child_by_title("View")
        if child_hwnd is None:
            logger.warning("No se encontro control hijo 'View' para sondeo.")
            self._log_process_windows_snapshot("sin control View")
            return
        logger.warning("Sondeo seguro sobre control hijo 'View': hwnd=%s", child_hwnd)
        try:
            self._click_child_center(child_hwnd, pause=0.6)
        except UIAutomationAborted:
            raise
        except Exception as exc:
            logger.warning("No se pudo clickear control 'View': %r", exc)
            self._log_process_windows_snapshot("fallo click View")
            return
        self._log_window_context("post click View")
        self._log_process_windows_snapshot("post click View")
        self._log_child_windows_snapshot("post click View")

    def _ingresar_angulos_en_dialogo(self, azimut: float, elevacion: float) -> None:
        try:
            self._run_guarded("seleccionar azimut", lambda: hotkey("ctrl", "a"), pause=0.05)
            self._run_guarded(
                "escribir azimut",
                lambda: write_text(str(int(azimut)), interval=0.05),
                pause=0.15,
            )
            self._run_guarded("pasar a elevacion", lambda: press_key("tab"), pause=0.15)
            self._run_guarded("seleccionar elevacion", lambda: hotkey("ctrl", "a"), pause=0.05)
            self._run_guarded(
                "escribir elevacion",
                lambda: write_text(str(int(elevacion)), interval=0.05),
                pause=0.15,
            )
            self._run_guarded("confirmar dialogo 3d", lambda: press_key("enter"), pause=0.4)
        except UIAutomationAborted:
            raise
        except Exception as exc:
            logger.warning("No se pudo ingresar angulos en dialogo: %r", exc)
            self._cancel_if_possible()
            raise
