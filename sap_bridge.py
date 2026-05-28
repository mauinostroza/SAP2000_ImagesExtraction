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
        self._comtypes = None
        self._com_initialized = False
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
            import comtypes
            import comtypes.client
        except ImportError as exc:
            raise ImportError("comtypes no está instalado. Ejecuta: pip install comtypes") from exc

        self._comtypes = comtypes
        try:
            comtypes.CoInitialize()
            self._com_initialized = True
        except Exception as exc:
            logger.debug("CoInitialize no se pudo ejecutar en este thread: %s", exc)

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
            self._uninitialize_com()
            raise RuntimeError(
                f"No se pudo conectar a SAP2000: {exc}\n"
                "Verifica que SAP2000 esté abierto y el modelo cargado."
            ) from exc

        self._model = self._sap_object.SapModel
        logger.info("Conexión COM a SAP2000 establecida")

    def disconnect(self) -> None:
        self._model = None
        self._sap_object = None
        self._uninitialize_com()
        logger.info("Conexión COM liberada")

    def _uninitialize_com(self) -> None:
        if not self._com_initialized or self._comtypes is None:
            return
        try:
            self._comtypes.CoUninitialize()
        except Exception as exc:
            logger.debug("CoUninitialize falló: %s", exc)
        finally:
            self._com_initialized = False
            self._comtypes = None

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
        if result is None:
            return []
        if isinstance(result, int):
            return []

        tolist = getattr(result, "tolist", None)
        if callable(tolist):
            try:
                result = tolist()
            except Exception:
                pass

        if isinstance(result, str):
            name = result.strip()
            return [name] if name else []

        if isinstance(result, tuple):
            for item in reversed(result):
                if isinstance(item, (list, tuple)):
                    names = SapBridge._coerce_names(item)
                    if names:
                        return names

            names: list[str] = []
            for item in result:
                if item is None or isinstance(item, int):
                    continue
                if callable(getattr(item, "tolist", None)):
                    names.extend(SapBridge._coerce_names(item.tolist()))
                    continue
                text = str(item).strip()
                if text:
                    names.append(text)
            return names

        if isinstance(result, list):
            names: list[str] = []
            for item in result:
                if item is None or isinstance(item, int):
                    continue
                if callable(getattr(item, "tolist", None)):
                    names.extend(SapBridge._coerce_names(item.tolist()))
                    continue
                text = str(item).strip()
                if text:
                    names.append(text)
            return names

        text = str(result).strip()
        return [text] if text else []

    def _resolve_attr(self, dotted_name: str):
        current = self.model
        for part in dotted_name.split("."):
            current = getattr(current, part)
        return current

    def _get_name_list(self, getter_name: str, operation: str) -> list[str]:
        getter = self._resolve_attr(getter_name)
        call_patterns = ((), (0,), (0, []), ([],), (0, None))

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
                continue

        if last_error is not None:
            raise RuntimeError(f"No se pudo obtener {operation}: {last_error}") from last_error
        return []

    def get_load_case_names(self) -> list[str]:
        return self._get_name_list("LoadCases.GetNameList", "GetNameList (casos)")

    def get_load_pattern_names(self) -> list[str]:
        return self._get_name_list("LoadPatterns.GetNameList", "GetNameList (patrones)")

    def get_combo_names(self) -> list[str]:
        return self._get_name_list("RespCombo.GetNameList", "GetNameList (combos)")

    def get_model_catalog(self) -> dict[str, list[str]]:
        return {
            "load_patterns": self.get_load_pattern_names(),
            "load_cases": self.get_load_case_names(),
            "combos": self.get_combo_names(),
        }

    @staticmethod
    def check_ret(ret: int, operation: str) -> None:
        if ret != 0:
            raise RuntimeError(f"SAP2000 API error {ret} en '{operation}'")
