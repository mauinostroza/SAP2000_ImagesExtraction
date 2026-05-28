"""
output_writer.py
Gestiona el naming de archivos, la carpeta de salida y el log de capturas.

Genera:
  outputs/
    ├── 001_geometria_3d.png
    ├── 002_carga_muerta_3d.png
    ├── ...
    └── capture_log.json          ← registro de cada captura con metadata

El log JSON permite auditar resultados, detectar capturas fallidas y
regenerar informes sin volver a ejecutar el script.
"""

from __future__ import annotations

import json
import logging
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

from view_controller import ViewConfig

logger = logging.getLogger(__name__)


@dataclass
class CaptureRecord:
    """Registro de una captura individual en el log."""
    index:        int
    filename:     str          # nombre sin extensión
    output_file:  str          # ruta relativa al output_dir
    view_type:    str
    display_type: str
    case_name:    str
    description:  str
    status:       str          # "ok" | "error"
    error_msg:    str = ""
    timestamp:    str = field(default_factory=lambda: datetime.now().isoformat())
    elapsed_ms:   int = 0


class OutputWriter:
    """Gestiona la carpeta de salida y el log de capturas.

    Uso:
        writer = OutputWriter("outputs/mi_proyecto")
        writer.start_session()
        path = writer.get_output_path(cfg, index=1)
        # ... captura ...
        writer.record_ok(cfg, index=1, output_path=path, elapsed_ms=320)
        writer.save_log()
    """

    def __init__(
        self,
        output_dir: str | Path = "outputs",
        prefix_index: bool = True,
        log_filename: str = "capture_log.json",
    ):
        """
        Args:
            output_dir:    Carpeta donde se guardan los PNG y el log.
            prefix_index:  Si True, antepone "001_", "002_" al filename.
            log_filename:  Nombre del archivo de log JSON.
        """
        self.output_dir   = Path(output_dir)
        self.prefix_index = prefix_index
        self.log_path     = self.output_dir / log_filename
        self._records: list[CaptureRecord] = []
        self._session_start: str = ""

    def start_session(self) -> None:
        """Crea la carpeta de salida y registra el inicio de sesión."""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._session_start = datetime.now().isoformat()
        logger.info(f"Output dir: {self.output_dir.resolve()}")

    def get_output_path(self, cfg: ViewConfig, index: int) -> Path:
        """Construye la ruta completa del PNG para una ViewConfig."""
        base = f"{index:03d}_{cfg.filename}" if self.prefix_index else cfg.filename
        # Sanitizar nombre (eliminar caracteres problemáticos en Windows)
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in base)
        return self.output_dir / f"{safe}.png"

    def record_ok(
        self,
        cfg: ViewConfig,
        index: int,
        output_path: Path,
        elapsed_ms: int = 0,
    ) -> None:
        """Registra una captura exitosa."""
        rec = CaptureRecord(
            index       = index,
            filename    = cfg.filename,
            output_file = str(output_path.relative_to(self.output_dir)),
            view_type   = cfg.view_type.name,
            display_type= cfg.display_type.name,
            case_name   = cfg.case_name,
            description = cfg.description,
            status      = "ok",
            elapsed_ms  = elapsed_ms,
        )
        self._records.append(rec)
        logger.info(f"[{index:03d}] OK  {output_path.name}  ({elapsed_ms} ms)")

    def record_error(
        self,
        cfg: ViewConfig,
        index: int,
        error: Exception,
    ) -> None:
        """Registra una captura fallida (no aborta el resto del plan)."""
        rec = CaptureRecord(
            index       = index,
            filename    = cfg.filename,
            output_file = "",
            view_type   = cfg.view_type.name,
            display_type= cfg.display_type.name,
            case_name   = cfg.case_name,
            description = cfg.description,
            status      = "error",
            error_msg   = str(error),
        )
        self._records.append(rec)
        logger.error(f"[{index:03d}] ERR {cfg.filename}: {error}")

    def save_log(self) -> Path:
        """Escribe el log JSON con todos los registros de la sesión."""
        payload = {
            "session_start": self._session_start,
            "session_end":   datetime.now().isoformat(),
            "total":         len(self._records),
            "ok":            sum(1 for r in self._records if r.status == "ok"),
            "errors":        sum(1 for r in self._records if r.status == "error"),
            "captures":      [asdict(r) for r in self._records],
        }
        with open(self.log_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        logger.info(f"Log guardado: {self.log_path}")
        return self.log_path

    def print_summary(self) -> None:
        """Imprime resumen en consola al terminar."""
        ok     = sum(1 for r in self._records if r.status == "ok")
        errors = sum(1 for r in self._records if r.status == "error")
        total  = len(self._records)
        print(f"\n{'─'*48}")
        print(f"  Capturas completadas : {ok}/{total}")
        if errors:
            print(f"  Errores              : {errors}")
            for r in self._records:
                if r.status == "error":
                    print(f"    ✗  {r.filename}: {r.error_msg}")
        print(f"  Output dir           : {self.output_dir.resolve()}")
        print(f"  Log                  : {self.log_path.name}")
        print(f"{'─'*48}\n")
