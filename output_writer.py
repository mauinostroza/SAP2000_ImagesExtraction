"""
output_writer.py
Gestiona el naming de archivos, la carpeta de salida y el log de capturas.
"""

from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

from view_controller import ViewConfig

logger = logging.getLogger(__name__)


@dataclass
class CaptureRecord:
    index: int
    filename: str
    output_file: str
    view_type: str
    display_type: str
    case_name: str
    description: str
    status: str
    error_msg: str = ""
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())
    elapsed_ms: int = 0


class OutputWriter:
    def __init__(
        self,
        output_dir: str | Path = "outputs",
        prefix_index: bool = True,
        log_filename: str = "capture_log.json",
    ):
        self.output_dir = Path(output_dir)
        self.prefix_index = prefix_index
        self.log_path = self.output_dir / log_filename
        self._records: list[CaptureRecord] = []
        self._session_start = ""

    def start_session(self) -> None:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._session_start = datetime.now().isoformat()
        logger.info("Output dir: %s", self.output_dir.resolve())

    def get_output_path(self, cfg: ViewConfig, index: int) -> Path:
        base = f"{index:03d}_{cfg.filename}" if self.prefix_index else cfg.filename
        safe = "".join(char if char.isalnum() or char in "-_" else "_" for char in base)
        return self.output_dir / f"{safe}.png"

    def record_ok(
        self,
        cfg: ViewConfig,
        index: int,
        output_path: Path,
        elapsed_ms: int = 0,
    ) -> None:
        record = CaptureRecord(
            index=index,
            filename=cfg.filename,
            output_file=str(output_path.relative_to(self.output_dir)),
            view_type=cfg.view_type.name,
            display_type=cfg.display_type.name,
            case_name=cfg.case_name,
            description=cfg.description,
            status="ok",
            elapsed_ms=elapsed_ms,
        )
        self._records.append(record)
        logger.info("[%03d] OK  %s  (%s ms)", index, output_path.name, elapsed_ms)

    def record_error(self, cfg: ViewConfig, index: int, error: Exception) -> None:
        record = CaptureRecord(
            index=index,
            filename=cfg.filename,
            output_file="",
            view_type=cfg.view_type.name,
            display_type=cfg.display_type.name,
            case_name=cfg.case_name,
            description=cfg.description,
            status="error",
            error_msg=str(error),
        )
        self._records.append(record)
        logger.error("[%03d] ERR %s: %s", index, cfg.filename, error)

    def save_log(self) -> Path:
        payload = {
            "session_start": self._session_start,
            "session_end": datetime.now().isoformat(),
            "total": len(self._records),
            "ok": sum(1 for record in self._records if record.status == "ok"),
            "errors": sum(1 for record in self._records if record.status == "error"),
            "captures": [asdict(record) for record in self._records],
        }
        with self.log_path.open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
        logger.info("Log guardado: %s", self.log_path)
        return self.log_path

    def print_summary(self) -> None:
        ok = sum(1 for record in self._records if record.status == "ok")
        errors = sum(1 for record in self._records if record.status == "error")
        total = len(self._records)
        print("\n" + "-" * 48)
        print(f"  Capturas completadas : {ok}/{total}")
        if errors:
            print(f"  Errores              : {errors}")
            for record in self._records:
                if record.status == "error":
                    print(f"    x  {record.filename}: {record.error_msg}")
        print(f"  Output dir           : {self.output_dir.resolve()}")
        print(f"  Log                  : {self.log_path.name}")
        print("-" * 48 + "\n")
