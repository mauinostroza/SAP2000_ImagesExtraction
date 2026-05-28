"""
view_controller.py
Control de vistas, casos de carga y ángulos de cámara en SAP2000 via OAPI.

La API de SAP2000 no expone ángulos de cámara como parámetros directos.
El control de la vista se hace a través de:
  - SapModel.View.RefreshView()
  - SapModel.View.SetDisplayOptionsMModel()
  - SapModel.View.SetActiveDisplayCase()   ← cambia caso visible
  - SapModel.View.SetView()                ← ángulo isométrico/planta/elevación

Nota sobre SetView: en SAP2000 v23/v24 la firma es:
  SetView(WindowNumber, ViewType)
  ViewType: 0=3D, 1=XY(planta), 2=XZ, 3=YZ
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from enum import IntEnum

logger = logging.getLogger(__name__)


# ── Enumeraciones ──────────────────────────────────────────────────────────────

class ViewType(IntEnum):
    """Tipos de vista disponibles en SapModel.View.SetView."""
    ISO_3D  = 0   # Vista isométrica 3D
    PLAN_XY = 1   # Planta (vista desde arriba)
    ELEV_XZ = 2   # Elevación X-Z (vista frontal)
    ELEV_YZ = 3   # Elevación Y-Z (vista lateral)


class DisplayType(IntEnum):
    """Tipo de display que se quiere mostrar sobre el modelo."""
    GEOMETRY_ONLY  = 0   # Solo geometría, sin cargas ni resultados
    LOAD_PATTERN   = 1   # Patrón de carga
    LOAD_CASE      = 2   # Caso de análisis o combinación
    MODE_SHAPE     = 3   # Forma modal
    FRAME_FORCES   = 4   # Fuerzas en barras (requiere análisis previo)
    DEFORMED       = 5   # Forma deformada


@dataclass
class ViewConfig:
    """Configuración completa de una captura de vista.

    Attributes:
        filename:       Nombre del archivo PNG de salida (sin extensión).
        view_type:      Ángulo de cámara (ViewType).
        display_type:   Qué mostrar sobre el modelo (DisplayType).
        case_name:      Nombre del caso/patrón/combo (requerido para LOAD_*, DEFORMED).
        mode_number:    Número de modo (solo para MODE_SHAPE, base 1).
        window_number:  Número de ventana SAP2000 (default 0 = ventana activa).
        render_delay:   Segundos extra de espera para renders pesados.
        description:    Texto descriptivo para el log.
    """
    filename:      str
    view_type:     ViewType      = ViewType.ISO_3D
    display_type:  DisplayType   = DisplayType.GEOMETRY_ONLY
    case_name:     str           = ""
    mode_number:   int           = 1
    window_number: int           = 0
    render_delay:  float         = 0.0
    description:   str           = ""


# ── Controlador de vistas ─────────────────────────────────────────────────────

class ViewController:
    """Aplica ViewConfig al modelo SAP2000 antes de capturar."""

    def __init__(self, sap_model, base_render_delay: float = 0.5):
        """
        Args:
            sap_model:         cSapModel de la conexión COM.
            base_render_delay: Espera mínima tras cambiar vista (segundos).
        """
        self._m     = sap_model
        self._delay = base_render_delay

    def apply(self, cfg: ViewConfig) -> None:
        """Aplica la configuración y espera el render.

        Orden:
          1. Cambiar el caso/patrón visible.
          2. Cambiar el ángulo de la cámara.
          3. Zoom-to-fit.
          4. RefreshView para forzar redibujado.
          5. Esperar render_delay total.
        """
        self._set_display(cfg)
        self._set_view_angle(cfg)
        self._zoom_all(cfg.window_number)
        self._refresh(cfg.window_number)
        total_wait = self._delay + cfg.render_delay
        logger.debug(f"Esperando {total_wait:.2f}s para render de '{cfg.filename}'")
        time.sleep(total_wait)

    # ── Helpers privados ──────────────────────────────────────────────────────

    def _set_display(self, cfg: ViewConfig) -> None:
        """Activa el caso o patrón de carga correspondiente."""
        if cfg.display_type == DisplayType.GEOMETRY_ONLY:
            # No hay un "reset display" directo; se puede desactivar
            # el display de resultados llamando sin caso
            return

        if cfg.display_type in (DisplayType.LOAD_PATTERN, DisplayType.LOAD_CASE,
                                 DisplayType.DEFORMED):
            if not cfg.case_name:
                logger.warning(f"display_type={cfg.display_type.name} requiere case_name")
                return
            ret = self._m.View.SetActiveDisplayCase(cfg.case_name)
            if ret != 0:
                logger.error(f"SetActiveDisplayCase('{cfg.case_name}') retornó {ret}")
            else:
                logger.debug(f"Caso activo: '{cfg.case_name}'")

        elif cfg.display_type == DisplayType.MODE_SHAPE:
            if not cfg.case_name:
                logger.warning("MODE_SHAPE requiere case_name (nombre del caso modal)")
                return
            ret = self._m.View.SetCaseModalShape(cfg.case_name, cfg.mode_number)
            if ret != 0:
                logger.error(f"SetCaseModalShape retornó {ret}")
            else:
                logger.debug(f"Modo {cfg.mode_number} del caso '{cfg.case_name}'")

        elif cfg.display_type == DisplayType.FRAME_FORCES:
            # SetDisplayResultsFrame requiere que el análisis esté corrido.
            # Los parámetros varían por versión; aquí se activa el caso y
            # el usuario debe haber configurado el display type antes.
            if cfg.case_name:
                self._m.View.SetActiveDisplayCase(cfg.case_name)

    def _set_view_angle(self, cfg: ViewConfig) -> None:
        """Cambia el ángulo de cámara."""
        ret = self._m.View.SetView(cfg.window_number, int(cfg.view_type))
        if ret != 0:
            logger.warning(
                f"SetView({cfg.window_number}, {cfg.view_type}) retornó {ret}. "
                "En algunas versiones el número de ventana puede diferir."
            )

    def _zoom_all(self, window_number: int) -> None:
        """Ajusta el zoom para mostrar todo el modelo."""
        ret = self._m.View.UnzoomAll(window_number)
        if ret != 0:
            logger.warning(f"UnzoomAll({window_number}) retornó {ret}")

    def _refresh(self, window_number: int) -> None:
        """Fuerza redibujado de la ventana."""
        # Mantener el zoom actual evita que SAP2000 vuelva a un encuadre
        # distinto después de UnzoomAll().
        ret = self._m.View.RefreshView(window_number, True)
        if ret != 0:
            logger.warning(f"RefreshView retornó {ret}")
