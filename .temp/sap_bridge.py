"""
sap_bridge.py
Conexión COM a SAP2000 ya abierto y objeto SapModel.

Usa comtypes (preferido sobre win32com para SAP2000 porque genera
wrappers tipados que exponen los métodos con sus firmas correctas).

Dependencias:
  pip install comtypes

Uso:
  bridge = SapBridge()
  bridge.connect()          # se adjunta al proceso SAP2000 corriendo
  model  = bridge.model     # cSapModel
  bridge.disconnect()
"""

from __future__ import annotations

import logging
from pathlib import Path

logger = logging.getLogger(__name__)


class SapBridge:
    """Encapsula la conexión COM a una instancia SAP2000 ya abierta."""

    def __init__(self, sap_dll_path: str | Path | None = None):
        """
        Args:
            sap_dll_path: Ruta al SAP2000v1.dll para registrar los wrappers comtypes.
                          Si es None, busca en la ruta de instalación estándar.
        """
        self._sap_object = None
        self._model      = None
        self._dll_path   = Path(sap_dll_path) if sap_dll_path else self._find_dll()

    # ── Localización del DLL ──────────────────────────────────────────────────

    @staticmethod
    def _find_dll() -> Path | None:
        """Busca SAP2000v1.dll en las rutas de instalación típicas."""
        candidates = [
            r"C:\Program Files\Computers and Structures\SAP2000 23\SAP2000v1.dll",
            r"C:\Program Files\Computers and Structures\SAP2000 24\SAP2000v1.dll",
            r"C:\Program Files\Computers and Structures\SAP2000 25\SAP2000v1.dll",
            r"C:\Program Files (x86)\Computers and Structures\SAP2000 23\SAP2000v1.dll",
        ]
        for c in candidates:
            p = Path(c)
            if p.exists():
                logger.debug(f"DLL encontrado: {p}")
                return p
        logger.warning("SAP2000v1.dll no encontrado en rutas estándar. "
                       "Pasa sap_dll_path explícitamente.")
        return None

    # ── Conexión ──────────────────────────────────────────────────────────────

    def connect(self) -> None:
        """Se adjunta a la instancia SAP2000 ya abierta.

        Flujo:
          1. Registra los wrappers comtypes con GetModule (solo la primera vez).
          2. Crea un Helper COM.
          3. Usa GetObject para adjuntarse al proceso existente (no abre una nueva
             instancia — el usuario ya debe tener SAP2000 abierto con el modelo).
          4. Obtiene SapModel.

        Raises:
            RuntimeError: Si SAP2000 no está corriendo o la conexión falla.
        """
        try:
            import comtypes.client
        except ImportError:
            raise ImportError("comtypes no está instalado. Ejecuta: pip install comtypes")

        # Registrar wrappers (genera caché la primera vez, rápido las siguientes)
        if self._dll_path and self._dll_path.exists():
            try:
                comtypes.client.GetModule(str(self._dll_path))
                logger.debug("Wrappers comtypes generados desde DLL")
            except Exception as e:
                logger.warning(f"GetModule falló ({e}), intentando conexión directa")

        try:
            import comtypes.gen.SAP2000v1 as sap_api
        except ImportError as e:
            raise RuntimeError(
                "No se pudo importar el wrapper comtypes de SAP2000. "
                "Pasa --sap-dll con la ruta a SAP2000v1.dll o genera el cache "
                "de comtypes en una instalación donde la DLL sea accesible."
            ) from e

        # Crear Helper y adjuntarse al proceso existente
        try:
            helper = comtypes.client.CreateObject("SAP2000v1.Helper")
            helper = helper.QueryInterface(sap_api.cHelper)
            self._sap_object = helper.GetObject("CSI.SAP2000.API.SapObject")
        except Exception as e:
            raise RuntimeError(
                f"No se pudo conectar a SAP2000: {e}\n"
                "Verifica que SAP2000 esté abierto y el modelo cargado."
            )

        self._model = self._sap_object.SapModel
        logger.info("Conexión COM a SAP2000 establecida")

    def disconnect(self) -> None:
        """Libera la referencia COM (no cierra SAP2000)."""
        self._model      = None
        self._sap_object = None
        logger.info("Conexión COM liberada")

    # ── Acceso al modelo ──────────────────────────────────────────────────────

    @property
    def model(self):
        """cSapModel — punto de entrada a toda la OAPI."""
        if self._model is None:
            raise RuntimeError("No hay conexión activa. Llama connect() primero.")
        return self._model

    # ── Context manager ───────────────────────────────────────────────────────

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, *_):
        self.disconnect()

    # ── Utilidades del modelo ─────────────────────────────────────────────────

    def get_load_case_names(self) -> list[str]:
        """Retorna todos los nombres de casos de carga definidos."""
        return self._extract_name_list(self.model.LoadCases.GetNameList())

    def get_load_pattern_names(self) -> list[str]:
        """Retorna todos los patrones de carga."""
        return self._extract_name_list(self.model.LoadPatterns.GetNameList())

    def get_combo_names(self) -> list[str]:
        """Retorna todos los nombres de combinaciones."""
        return self._extract_name_list(self.model.RespCombo.GetNameList())

    @staticmethod
    def _extract_name_list(result) -> list[str]:
        """Normaliza respuestas de comtypes/COM con arrays de nombres."""
        if result is None:
            return []

        if isinstance(result, tuple):
            # Comtypes suele devolver solo los out-params; en SAP2000 eso suele
            # ser (cantidad, lista_de_nombres) o variantes cercanas.
            for item in reversed(result):
                if isinstance(item, (list, tuple)):
                    return [str(v) for v in item if v is not None and str(v)]
            return [str(v) for v in result if v is not None and str(v)]

        if isinstance(result, (list, tuple)):
            return [str(v) for v in result if v is not None and str(v)]

        return [str(result)] if str(result) else []

    def check_ret(self, ret: int, operation: str) -> None:
        """Lanza excepción si la llamada OAPI retornó error."""
        if ret != 0:
            raise RuntimeError(f"SAP2000 API error {ret} en '{operation}'")
