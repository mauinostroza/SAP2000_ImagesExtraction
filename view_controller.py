"""
view_controller.py
Control de vistas, casos de carga y angulos de camara en SAP2000 via OAPI.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from enum import IntEnum

logger = logging.getLogger(__name__)


class ViewType(IntEnum):
    ISO_3D = 0
    PLAN_XY = 1
    ELEV_XZ = 2
    ELEV_YZ = 3


class DisplayType(IntEnum):
    GEOMETRY_ONLY = 0
    LOAD_PATTERN = 1
    LOAD_CASE = 2
    MODE_SHAPE = 3
    FRAME_FORCES = 4
    DEFORMED = 5


@dataclass
class ViewConfig:
    filename: str
    view_type: ViewType = ViewType.ISO_3D
    display_type: DisplayType = DisplayType.GEOMETRY_ONLY
    case_name: str = ""
    mode_number: int = 1
    window_number: int = 0
    render_delay: float = 0.0
    description: str = ""


class ViewController:
    def __init__(self, sap_model, base_render_delay: float = 0.5):
        self._m = sap_model
        self._delay = base_render_delay

    def apply(self, cfg: ViewConfig) -> None:
        self._set_display(cfg)
        target_window = self._set_view_angle(cfg)
        self._zoom_all(target_window)
        self._refresh(target_window)
        total_wait = self._delay + cfg.render_delay
        logger.debug("Esperando %.2fs para render de '%s'", total_wait, cfg.filename)
        time.sleep(total_wait)

    @staticmethod
    def _window_candidates(window_number: int) -> list[int]:
        ordered: list[int] = []
        for candidate in (window_number, 0, 1):
            if candidate not in ordered:
                ordered.append(candidate)
        return ordered

    def _set_display(self, cfg: ViewConfig) -> None:
        if cfg.display_type == DisplayType.GEOMETRY_ONLY:
            return

        if cfg.display_type in (DisplayType.LOAD_PATTERN, DisplayType.LOAD_CASE, DisplayType.DEFORMED):
            if not cfg.case_name:
                raise ValueError(f"display_type={cfg.display_type.name} requiere case_name")
            ret = self._m.View.SetActiveDisplayCase(cfg.case_name)
            if ret != 0:
                raise RuntimeError(f"SetActiveDisplayCase('{cfg.case_name}') retorno {ret}")
            return

        if cfg.display_type == DisplayType.MODE_SHAPE:
            if not cfg.case_name:
                raise ValueError("MODE_SHAPE requiere case_name")
            method = getattr(self._m.View, "SetCaseModalShape", None)
            if method is None:
                raise RuntimeError("La API actual no expone View.SetCaseModalShape")
            ret = method(cfg.case_name, cfg.mode_number)
            if ret != 0:
                raise RuntimeError(
                    f"SetCaseModalShape('{cfg.case_name}', {cfg.mode_number}) retorno {ret}"
                )
            return

        if cfg.display_type == DisplayType.FRAME_FORCES and cfg.case_name:
            ret = self._m.View.SetActiveDisplayCase(cfg.case_name)
            if ret != 0:
                raise RuntimeError(f"SetActiveDisplayCase('{cfg.case_name}') retorno {ret}")

    def _set_view_angle(self, cfg: ViewConfig) -> int:
        failures: list[str] = []
        for window_number in self._window_candidates(cfg.window_number):
            ret = self._m.View.SetView(window_number, int(cfg.view_type))
            if ret == 0:
                if window_number != cfg.window_number:
                    logger.warning(
                        "SetView funciono con window_number=%s; el plan pedia %s",
                        window_number,
                        cfg.window_number,
                    )
                return window_number
            failures.append(f"{window_number}->{ret}")

        logger.warning(
            "SetView fallo para '%s' (%s). Se capturara la vista actual de SAP2000.",
            cfg.filename,
            ", ".join(failures),
        )
        return cfg.window_number

    def _zoom_all(self, window_number: int) -> None:
        ret = self._m.View.UnzoomAll(window_number)
        if ret != 0:
            logger.warning("UnzoomAll(%s) retorno %s", window_number, ret)

    def _refresh(self, window_number: int) -> None:
        ret = self._m.View.RefreshView(window_number, True)
        if ret != 0:
            logger.warning("RefreshView(%s) retorno %s", window_number, ret)
