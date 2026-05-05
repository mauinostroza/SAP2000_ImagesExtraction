"""Interfaz simple en Tkinter para conectar con SAP2000 y extraer capturas.

La GUI usa el backend existente de `sap_imagenes.py` sin cambiar la lógica
principal de capturas. Solo añade el pegamento para:
- elegir el Excel de configuración,
- conectar SAP2000 por separado,
- ejecutar la extracción,
- mostrar logs y errores en pantalla.
"""

import argparse
import logging
import os
import queue
import threading
import traceback
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import openpyxl

import sap_imagenes as backend


class _QueueLogHandler(logging.Handler):
    """Envia mensajes de log a una cola para pintarlos en la UI."""

    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
        except Exception:
            msg = record.getMessage()
        self.log_queue.put(("log", record.levelname, msg))


class SAP2000GuiApp:
    def __init__(self, root: tk.Tk, config_path: Optional[str] = None, allow_unsafe_output: bool = False):
        self.root = root
        self.root.title("SAP2000 Capture GUI")
        self.root.geometry("920x620")
        self.root.minsize(820, 560)

        self.log_queue: queue.Queue = queue.Queue()
        self.worker_thread = None
        self.conector = None
        self.last_config = None
        self.last_config_path: str = ""
        self.log_handler: Optional[_QueueLogHandler] = None
        self.poll_after_id: Optional[str] = None
        self.is_closing = False

        self.config_path_var = tk.StringVar(value=config_path or self._default_config_path())
        self.status_var = tk.StringVar(value="Desconectado")
        self.details_var = tk.StringVar(value="Selecciona un Excel y conecta SAP2000.")
        self.allow_unsafe_var = tk.BooleanVar(value=allow_unsafe_output)

        self._build_ui()
        self._attach_log_handler()
        self._append_log(
            "INFO",
            "GUI lista. Selecciona el Excel y usa 'Conectar a SAP2000'.",
        )
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.poll_after_id = self.root.after(100, self._poll_log_queue)

    @staticmethod
    def _is_inactive_result(resultado: dict) -> bool:
        if resultado.get("estado") == "inactiva":
            return True
        return (
            resultado.get("ok") is None
            and resultado.get("mensaje") == "INACTIVO"
            and not resultado.get("error_tipo")
        )

    def _has_connection(self) -> bool:
        return self.conector is not None and getattr(self.conector, "sap_model", None) is not None

    def _default_config_path(self) -> str:
        base_dir = Path(__file__).resolve().parent
        candidate = base_dir / "SAP2000_Capturas.xlsx"
        return str(candidate if candidate.exists() else Path.cwd() / "SAP2000_Capturas.xlsx")

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        main = ttk.Frame(self.root, padding=12)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(1, weight=1)
        main.rowconfigure(2, weight=1)

        title = ttk.Label(
            main,
            text="SAP2000 Capture GUI",
            font=("Segoe UI", 15, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        ttk.Label(main, text="Excel de configuración:").grid(row=1, column=0, sticky="w")
        config_entry = ttk.Entry(main, textvariable=self.config_path_var)
        config_entry.grid(row=1, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(main, text="Buscar...", command=self._browse_config).grid(
            row=1, column=2, sticky="e"
        )

        options = ttk.Frame(main)
        options.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10, 10))
        options.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            options,
            text="Permitir salida insegura",
            variable=self.allow_unsafe_var,
        ).grid(row=0, column=0, sticky="w")

        buttons = ttk.Frame(main)
        buttons.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=1)

        self.connect_btn = ttk.Button(
            buttons,
            text="Conectar a SAP2000",
            command=self._on_connect,
        )
        self.connect_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.extract_btn = ttk.Button(
            buttons,
            text="Extraer fotos",
            command=self._on_extract,
            state="disabled",
        )
        self.extract_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        status_frame = ttk.LabelFrame(main, text="Estado", padding=10)
        status_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        status_frame.columnconfigure(0, weight=1)

        ttk.Label(status_frame, textvariable=self.status_var, font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(status_frame, textvariable=self.details_var, wraplength=840).grid(
            row=1, column=0, sticky="w", pady=(4, 0)
        )

        log_frame = ttk.LabelFrame(main, text="Log", padding=8)
        log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew")
        main.rowconfigure(5, weight=1)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = ScrolledText(log_frame, height=18, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text.tag_configure("INFO", foreground="#1a1a1a")
        self.log_text.tag_configure("WARNING", foreground="#a66a00")
        self.log_text.tag_configure("ERROR", foreground="#b00020")
        self.log_text.tag_configure("DEBUG", foreground="#666666")

    def _attach_log_handler(self) -> None:
        if self.log_handler is not None:
            return

        handler = _QueueLogHandler(self.log_queue)
        handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", "%H:%M:%S"))
        backend.log.addHandler(handler)
        backend.log.setLevel(logging.INFO)
        self.log_handler = handler

    def _detach_log_handler(self) -> None:
        if self.log_handler is None:
            return

        backend.log.removeHandler(self.log_handler)
        self.log_handler.close()
        self.log_handler = None

    def _schedule_ui(self, callback, delay_ms: int = 0) -> None:
        if self.is_closing:
            return

        try:
            self.root.after(delay_ms, callback)
        except tk.TclError:
            self.is_closing = True

    def _poll_log_queue(self) -> None:
        if self.is_closing:
            return

        try:
            while True:
                kind, level, message = self.log_queue.get_nowait()
                if kind == "log":
                    self._append_log(level, message)
        except queue.Empty:
            pass
        if not self.is_closing:
            try:
                self.poll_after_id = self.root.after(100, self._poll_log_queue)
            except tk.TclError:
                self.is_closing = True

    def _append_log(self, level: str, message: str) -> None:
        if self.is_closing:
            return
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n", level if level in ("INFO", "WARNING", "ERROR", "DEBUG") else "INFO")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_status(self, status: str, details: str) -> None:
        self.status_var.set(status)
        self.details_var.set(details)

    def _browse_config(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecciona el Excel de configuración",
            filetypes=[
                ("Excel", "*.xlsx *.xlsm"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if path:
            self.config_path_var.set(path)

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.connect_btn.configure(state=state)
        self.extract_btn.configure(state="disabled" if busy else ("normal" if self._has_connection() else "disabled"))

    def _run_worker(self, target, busy_status: str) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("SAP2000", "Ya hay una tarea en ejecución.")
            return

        self._set_status(busy_status, "Procesando. Revisa el panel de log para más detalles.")
        self._set_busy(True)

        def runner():
            try:
                target()
            except Exception as exc:
                backend.log.error(str(exc))
                backend.log.error(traceback.format_exc())
                self._schedule_ui(lambda: self._set_status("Error", str(exc)))
                self._schedule_ui(
                    lambda: messagebox.showerror("SAP2000 Capture GUI", str(exc))
                )
            finally:
                self._schedule_ui(self._finish_worker)

        self.worker_thread = threading.Thread(target=runner, daemon=True)
        self.worker_thread.start()

    def _finish_worker(self) -> None:
        self._set_busy(False)
        if self._has_connection():
            self.extract_btn.configure(state="normal")
        else:
            self.extract_btn.configure(state="disabled")
            if self.status_var.get() == "Conectando":
                self._set_status("Desconectado", "Conecta SAP2000 antes de extraer fotos.")

    def _read_config_workbook(self) -> openpyxl.Workbook:
        ruta_config = self.config_path_var.get().strip()
        if not ruta_config:
            raise ValueError("Selecciona un archivo Excel de configuración.")
        if not os.path.exists(ruta_config):
            raise FileNotFoundError(f"No existe el archivo: {ruta_config}")

        return openpyxl.load_workbook(ruta_config, data_only=True)

    def _read_connection_config(self) -> dict:
        ruta_config = self.config_path_var.get().strip()
        wb = self._read_config_workbook()
        try:
            try:
                sh_cfg = wb["CONFIG"]
            except KeyError as exc:
                raise ValueError("El Excel no tiene la hoja CONFIG.") from exc

            sap_dll = sh_cfg["B2"].value or backend.SAP_DLL_PATH
            return {
                "ruta_config": ruta_config,
                "sap_dll": sap_dll,
            }
        finally:
            wb.close()

    def _read_extraction_config(self) -> dict:
        ruta_config = self.config_path_var.get().strip()
        if not ruta_config:
            raise ValueError("Selecciona un archivo Excel de configuración.")

        config = backend.cargar_configuracion_desde_excel(
            ruta_config,
            allow_unsafe_output=self.allow_unsafe_var.get(),
        )
        config["ruta_config"] = ruta_config
        return config

    def _connect_task(self) -> None:
        config = self._read_connection_config()
        self.last_config = config
        self.last_config_path = config["ruta_config"]

        backend.log.info(f"Configuración cargada: {config['ruta_config']}")
        backend.log.info(f"Ruta DLL SAP2000: {config['sap_dll']}")

        if self._has_connection():
            backend.log.info("Se reemplazará la conexión SAP2000 anterior por una nueva.")

        # Invalidar la conexión anterior antes de reconectar para no dejar
        # una sesión obsoleta habilitada si el nuevo intento falla.
        self.conector = None

        conector = backend.SAP2000Conector(config["sap_dll"])
        conector.conectar()

        self.conector = conector
        self._schedule_ui(
            lambda: self._set_status(
                "Conectado",
                "SAP2000 conectado. Ya puedes extraer fotos.",
            )
        )
        backend.log.info("Conexión a SAP2000 lista desde la GUI.")

    def _extract_task(self) -> None:
        if not self._has_connection():
            raise RuntimeError("Primero debes conectar a SAP2000.")

        config = self._read_extraction_config()
        self.last_config = config
        self.last_config_path = config["ruta_config"]

        capturas = config["capturas"]
        if not capturas:
            raise RuntimeError("No hay filas válidas en la hoja CAPTURAS.")

        activas = [c for c in capturas if c.get("activo", False)]
        backend.log.info(f"Capturas configuradas: {len(capturas)} | Activas: {len(activas)}")

        resultado = backend.ejecutar_trabajo_capturas(
            config,
            sap_model=self.conector.sap_model,
            conectar_si_falta=False,
        )

        if resultado["stage"] != "captura":
            self.conector = None
            raise RuntimeError(f"{resultado['stage']}: {resultado['mensaje']}")

        resultados = resultado.get("resultados", [])
        resumen_backend = resultado.get("resumen", {})
        omitidas = resumen_backend.get(
            "inactivas",
            sum(1 for item in resultados if self._is_inactive_result(item)),
        )
        ok_count = resumen_backend.get(
            "ok",
            sum(1 for item in resultados if item.get("estado") == "ok"),
        )
        err_count = resumen_backend.get(
            "errores",
            sum(1 for item in resultados if item.get("contabiliza_error")),
        )

        if err_count > 0:
            detalles_error = []
            for item in resultados:
                if item.get("ok") or self._is_inactive_result(item):
                    continue
                nombre = item.get("nombre_imagen") or item.get("archivo") or "(sin nombre)"
                mensaje = item.get("mensaje") or "Error sin detalle"
                detalles_error.append(f"- {nombre}: {mensaje}")

            resumen = f"Extracción completada con errores: {ok_count} OK, {err_count} errores"
            if omitidas:
                resumen += f", {omitidas} omitidas"
            resumen += "."

            backend.log.warning(resumen)
            if detalles_error:
                backend.log.warning("Errores por captura:\n" + "\n".join(detalles_error))

            self._schedule_ui(
                lambda: self._set_status("Completado con errores", resumen)
            )
            self._schedule_ui(
                lambda: messagebox.showwarning("SAP2000 Capture GUI", resumen)
            )
            return

        resumen = f"Extracción terminada: {ok_count} OK, 0 errores"
        if omitidas:
            resumen += f", {omitidas} omitidas"
        resumen += "."

        self._schedule_ui(
            lambda: self._set_status(
                "Listo",
                resumen,
            )
        )

    def _on_connect(self) -> None:
        self._run_worker(self._connect_task, "Conectando")

    def _on_extract(self) -> None:
        self._run_worker(self._extract_task, "Extrayendo")

    def _on_close(self) -> None:
        if self.is_closing:
            return

        self.is_closing = True

        if self.poll_after_id is not None:
            try:
                self.root.after_cancel(self.poll_after_id)
            except tk.TclError:
                pass
            self.poll_after_id = None

        self._detach_log_handler()

        try:
            self.root.destroy()
        except tk.TclError:
            pass


def main(config_path: Optional[str] = None, allow_unsafe_output: bool = False) -> None:
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        raise RuntimeError(
            "La interfaz gráfica requiere un entorno con pantalla disponible."
        ) from exc
    app = SAP2000GuiApp(
        root,
        config_path=config_path,
        allow_unsafe_output=allow_unsafe_output,
    )
    root.mainloop()


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="SAP2000 Capture GUI"
    )
    parser.add_argument(
        "--config",
        metavar="EXCEL",
        help="Ruta al Excel de configuración que abrirá la GUI.",
    )
    parser.add_argument(
        "--allow-unsafe-output",
        action="store_true",
        help=(
            "Permitir una carpeta de salida absoluta o fuera de la carpeta base del Excel. "
            "La GUI abrirá con esa opción activada."
        ),
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    return parser.parse_args(argv)


if __name__ == "__main__":
    try:
        args = parse_args()
        main(
            config_path=args.config,
            allow_unsafe_output=args.allow_unsafe_output,
        )
    except Exception as exc:
        print(str(exc))
        raise SystemExit(1)
