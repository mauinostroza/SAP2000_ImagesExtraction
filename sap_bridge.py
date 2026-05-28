"""
sap_bridge.py
Conexión COM a SAP2000 ya abierto y acceso a SapModel.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


class SapBridge:
    def __init__(self, sap_dll_path: str | Path | None = None):
        self._sap_object = None
        self._model = None
        self._dll_path = Path(sap_dll_path) if sap_dll_path else self._find_dll()

    @staticmethod
    def _find_dll() -> Path | None:
        candidates = [
            r"C:\Program Files\Computers and Structures\SAP2000 23\SAP2000v1.dll",
            r"C:\Program Files\Computers and Structures\SAP2000 24\SAP2000v1.dll",
            r"C:\Program Files\Computers and Structures\SAP2000 25\SAP2000v1.dll",
            r"C:\Program Files (x86)\Computers and Structures\SAP2000 23\SAP2000v1.dll",
        ]
        for candidate in candidates:
            path = Path(candidate)
            if path.exists():
                logger.debug("DLL encontrado: %s", path)
                return path
        logger.warning(
            "SAP2000v1.dll no encontrado en rutas estándar. "
            "Pasa --sap-dll explícitamente si hace falta."
        )
        return None

    def connect(self) -> None:
        try:
            import comtypes.client
        except ImportError as exc:
            raise ImportError("comtypes no está instalado. Ejecuta: pip install comtypes") from exc

        if self._dll_path and self._dll_path.exists():
            try:
                comtypes.client.GetModule(str(self._dll_path))
                logger.debug("Wrappers comtypes generados desde DLL")
            except Exception as exc:
                logger.warning("GetModule falló (%s), intentando conexión directa", exc)

        try:
            helper = comtypes.client.CreateObject("SAP2000v1.Helper")
            try:
                import comtypes.gen.SAP2000v1 as sap_api

                helper = helper.QueryInterface(sap_api.cHelper)
            except Exception as exc:
                logger.debug("QueryInterface tipado no disponible: %s", exc)
            self._sap_object = helper.GetObject("CSI.SAP2000.API.SapObject")
        except Exception as exc:
            raise RuntimeError(
                f"No se pudo conectar a SAP2000: {exc}\n"
                "Verifica que SAP2000 esté abierto y el modelo cargado."
            ) from exc

        self._model = self._sap_object.SapModel
        logger.info("Conexión COM a SAP2000 establecida")

    def disconnect(self) -> None:
        self._model = None
        self._sap_object = None
        logger.info("Conexión COM liberada")

    @property
    def model(self):
        if self._model is None:
            raise RuntimeError("No hay conexión activa. Llama connect() primero.")
        return self._model

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, *_args):
        self.disconnect()

    @staticmethod
    def _coerce_names(result: Any) -> list[str]:
        if isinstance(result, tuple):
            for item in reversed(result):
                if isinstance(item, (list, tuple)):
                    return [str(name) for name in item]
            if len(result) >= 2 and isinstance(result[1], (list, tuple)):
                return [str(name) for name in result[1]]
        if isinstance(result, (list, tuple)):
            return [str(name) for name in result]
        return []

    def _resolve_attr(self, dotted_name: str):
        current = self.model
        for part in dotted_name.split("."):
            current = getattr(current, part)
        return current

    def _get_name_list(self, getter_name: str, operation: str) -> list[str]:
        getter = self._resolve_attr(getter_name)
        call_patterns = ((), (0, []))

        last_error: Exception | None = None
        for args in call_patterns:
            try:
                result = getter(*args)
                if result is None:
                    return []
                if isinstance(result, int):
                    if result != 0:
                        logger.warning("%s retornó %s", operation, result)
                    return []
                names = self._coerce_names(result)
                if names:
                    return names
            except TypeError as exc:
                last_error = exc
            except Exception as exc:
                last_error = exc
                break

        if last_error is not None:
            raise RuntimeError(f"No se pudo obtener {operation}: {last_error}") from last_error
        return []

    def get_load_case_names(self) -> list[str]:
        return self._get_name_list("LoadCases.GetNameList", "GetNameList (casos)")

    def get_load_pattern_names(self) -> list[str]:
        return self._get_name_list("LoadPatterns.GetNameList", "GetNameList (patrones)")

    def get_combo_names(self) -> list[str]:
        return self._get_name_list("RespCombo.GetNameList", "GetNameList (combos)")

    @staticmethod
    def check_ret(ret: int, operation: str) -> None:
        if ret != 0:
            raise RuntimeError(f"SAP2000 API error {ret} en '{operation}'")
