"""
capture_plan.py
Lee el plan de capturas desde un archivo JSON o Excel (.xlsx).
"""

from __future__ import annotations

import json
import logging
from pathlib import Path

from view_controller import DisplayType, ViewConfig, ViewType

logger = logging.getLogger(__name__)


def load_plan(path: str | Path) -> list[ViewConfig]:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Plan no encontrado: {path}")

    suffix = path.suffix.lower()
    if suffix == ".json":
        return _load_json(path)
    if suffix in (".xlsx", ".xlsm"):
        return _load_excel(path)
    if suffix == ".xls":
        raise ValueError("Formato .xls no soportado; usa .xlsx, .xlsm o .json")
    raise ValueError(f"Formato no soportado: '{suffix}'. Usa .json, .xlsx o .xlsm")


def _load_json(path: Path) -> list[ViewConfig]:
    with path.open(encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, list):
        raise ValueError("El JSON debe ser una lista de objetos de captura")
    return [_dict_to_config(item, i) for i, item in enumerate(data, start=1)]


def _load_excel(path: Path) -> list[ViewConfig]:
    try:
        import openpyxl
    except ImportError as exc:
        raise ImportError("openpyxl no está instalado. Ejecuta: pip install openpyxl") from exc

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    headers = [str(cell.value).strip().lower() if cell.value else "" for cell in ws[1]]
    if "filename" not in headers:
        raise ValueError(
            f"El Excel debe tener al menos la columna 'filename'. Encontradas: {headers}"
        )

    configs: list[ViewConfig] = []
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(value is None for value in row):
            continue
        row_dict = {headers[i]: (row[i] if i < len(row) else None) for i in range(len(headers))}
        try:
            configs.append(_dict_to_config(row_dict, row_num))
        except Exception as exc:
            logger.warning("Fila %s ignorada: %s", row_num, exc)
    return configs


def _dict_to_config(data: dict, index: int) -> ViewConfig:
    def _str(key: str, default: str = "") -> str:
        value = data.get(key)
        return str(value).strip() if value is not None else default

    def _int(key: str, default: int = 0) -> int:
        value = data.get(key)
        try:
            return int(value) if value is not None else default
        except (TypeError, ValueError):
            return default

    def _float(key: str, default: float = 0.0) -> float:
        value = data.get(key)
        try:
            return float(value) if value is not None else default
        except (TypeError, ValueError):
            return default

    filename = _str("filename")
    if not filename:
        raise ValueError(f"Entrada {index}: 'filename' es obligatorio")

    vt_str = _str("view_type", "ISO_3D").upper()
    try:
        view_type = ViewType[vt_str]
    except KeyError as exc:
        valid = [item.name for item in ViewType]
        raise ValueError(f"view_type='{vt_str}' inválido. Opciones: {valid}") from exc

    dt_str = _str("display_type", "GEOMETRY_ONLY").upper()
    try:
        display_type = DisplayType[dt_str]
    except KeyError as exc:
        valid = [item.name for item in DisplayType]
        raise ValueError(f"display_type='{dt_str}' inválido. Opciones: {valid}") from exc

    return ViewConfig(
        filename=filename,
        view_type=view_type,
        display_type=display_type,
        case_name=_str("case_name"),
        mode_number=_int("mode_number", 1),
        window_number=_int("window_number", 0),
        render_delay=_float("render_delay", 0.0),
        description=_str("description"),
    )


def generate_sample_plan(output_path: str | Path = "capture_plan.json") -> Path:
    sample = [
        {
            "filename": "geometria_3d",
            "view_type": "ISO_3D",
            "display_type": "GEOMETRY_ONLY",
            "description": "Vista isométrica, solo geometría",
        },
        {
            "filename": "geometria_planta",
            "view_type": "PLAN_XY",
            "display_type": "GEOMETRY_ONLY",
            "description": "Vista en planta",
        },
        {
            "filename": "geometria_elevacion_xz",
            "view_type": "ELEV_XZ",
            "display_type": "GEOMETRY_ONLY",
            "description": "Elevación X-Z",
        },
        {
            "filename": "carga_muerta_3d",
            "view_type": "ISO_3D",
            "display_type": "LOAD_CASE",
            "case_name": "DEAD",
            "description": "Carga muerta, vista 3D",
        },
        {
            "filename": "carga_viva_3d",
            "view_type": "ISO_3D",
            "display_type": "LOAD_CASE",
            "case_name": "LIVE",
            "description": "Carga viva, vista 3D",
        },
        {
            "filename": "sismo_x_planta",
            "view_type": "PLAN_XY",
            "display_type": "LOAD_CASE",
            "case_name": "QUAKE-X",
            "description": "Sismo X, vista en planta",
        },
        {
            "filename": "sismo_y_planta",
            "view_type": "PLAN_XY",
            "display_type": "LOAD_CASE",
            "case_name": "QUAKE-Y",
            "description": "Sismo Y, vista en planta",
        },
        {
            "filename": "modo_1_3d",
            "view_type": "ISO_3D",
            "display_type": "MODE_SHAPE",
            "case_name": "MODAL",
            "mode_number": 1,
            "render_delay": 0.3,
            "description": "Modo 1 de vibración",
        },
        {
            "filename": "modo_2_3d",
            "view_type": "ISO_3D",
            "display_type": "MODE_SHAPE",
            "case_name": "MODAL",
            "mode_number": 2,
            "description": "Modo 2 de vibración",
        },
        {
            "filename": "deformada_combo_env",
            "view_type": "ISO_3D",
            "display_type": "DEFORMED",
            "case_name": "ENVELOPE",
            "render_delay": 0.5,
            "description": "Forma deformada, envolvente",
        },
    ]
    output_path = Path(output_path)
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(sample, handle, ensure_ascii=False, indent=2)
    logger.info("Plan de ejemplo generado en: %s", output_path)
    return output_path


if __name__ == "__main__":
    generated = generate_sample_plan()
    print(f"Plan de ejemplo generado: {generated}")
