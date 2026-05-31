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
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
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


VIEW_ANGLES = {
    SapUIView.ISO_3D: (225, 30),
}


class UIAutomationAborted(RuntimeError):
    pass


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


def press_key(key: str) -> None:
    vk = _vk_for_key(key)
    _send_input(_key_event(vk), _key_event(vk, keyup=True))


def hotkey(*keys: str) -> None:
    vks = [_vk_for_key(key) for key in keys]
    downs = [_key_event(vk) for vk in vks]
    ups = [_key_event(vk, keyup=True) for vk in reversed(vks)]
    _send_input(*(downs + ups))


def write_text(text: str, interval: float = 0.04) -> None:
    for char in text:
        _send_input(_unicode_event(char), _unicode_event(char, keyup=True))
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
        return self._get_window_pid(hwnd) == self.sap_pid

    def _assert_safe_context(self, step: str) -> None:
        if not self.enabled:
            raise UIAutomationAborted("automatizacion UI no armada")
        if self.stop_requested():
            raise UIAutomationAborted(f"automatizacion UI detenida antes de {step}")
        if self._is_escape_pressed():
            raise UIAutomationAborted(f"tecla Esc detectada antes de {step}")
        if not self._foreground_belongs_to_sap():
            raise UIAutomationAborted(f"foco fuera de SAP2000 antes de {step}")

    def _run_guarded(self, step: str, action, pause: float = 0.0) -> None:
        self._assert_safe_context(step)
        action()
        if pause > 0:
            time.sleep(pause)
        self._assert_safe_context(step)

    def activar(self) -> None:
        user32.ShowWindow(self.hwnd_principal, 5)
        user32.SetForegroundWindow(self.hwnd_principal)
        time.sleep(0.25)
        self._assert_safe_context("activar ventana")

    def set_vista(self, view_type: SapUIView) -> bool:
        if not self.enabled:
            logger.info("Automatizacion UI desarmada; no se enviaran teclas para vista.")
            return False
        if view_type not in VIEW_ANGLES:
            logger.info("UI vista no implementada para %s; se conserva la actual.", view_type.name)
            return False

        azimut, elevacion = VIEW_ANGLES[view_type]
        logger.info("UI vista %s -> az=%s el=%s", view_type.name, azimut, elevacion)
        self.activar()
        try:
            self._abrir_dialogo_set3dview()
            self._ingresar_angulos_en_dialogo(azimut, elevacion)
            time.sleep(0.3)
            self._assert_safe_context("fin set_vista")
            return True
        except UIAutomationAborted as exc:
            logger.warning("Automatizacion UI abortada en vista: %s", exc)
            return False

    def mostrar_cargas_patron(self, patron: str) -> bool:
        if not self.enabled:
            logger.info("Automatizacion UI desarmada; no se enviaran teclas para cargas.")
            return False
        try:
            self.activar()
            self._run_guarded("abrir menu display", lambda: hotkey("alt", "d"), pause=0.3)
            for idx in range(2):
                self._run_guarded(f"display down {idx+1}", lambda: press_key("down"), pause=0.1)
            self._run_guarded("display submenu right", lambda: press_key("right"), pause=0.2)
            self._run_guarded("display load assigns down", lambda: press_key("down"), pause=0.1)
            self._run_guarded("abrir dialogo load assigns", lambda: press_key("enter"), pause=0.5)
            self._run_guarded("seleccionar patron", lambda: hotkey("ctrl", "a"), pause=0.05)
            self._run_guarded("escribir patron", lambda: write_text(patron, interval=0.04), pause=0.2)
            self._run_guarded("confirmar patron", lambda: press_key("enter"), pause=0.4)
            return True
        except UIAutomationAborted as exc:
            logger.warning("Automatizacion UI abortada en cargas: %s", exc)
            return False
        except Exception as exc:
            logger.warning("No se pudo seleccionar patron en dialogo: %r", exc)
            try:
                if self._foreground_belongs_to_sap():
                    press_key("escape")
            except Exception:
                pass
            return False

    def _abrir_dialogo_set3dview(self) -> None:
        self._run_guarded("abrir menu view", lambda: hotkey("alt", "v"), pause=0.3)
        for idx in range(6):
            self._run_guarded(f"view down {idx+1}", lambda: press_key("down"), pause=0.1)
        self._run_guarded("view submenu right", lambda: press_key("right"), pause=0.2)
        for idx in range(5):
            self._run_guarded(f"rotate 3d down {idx+1}", lambda: press_key("down"), pause=0.1)
        self._run_guarded("abrir dialogo 3d", lambda: press_key("enter"), pause=0.5)

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
            try:
                if self._foreground_belongs_to_sap():
                    press_key("escape")
            except Exception:
                pass
            raise
