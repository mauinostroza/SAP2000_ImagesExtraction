from __future__ import annotations

import argparse
from collections.abc import Callable
import logging
import queue
import sys
import threading
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from capture_plan import load_plan, save_plan
from sap_bridge import SapBridge
from view_controller import DisplayType, ViewConfig, ViewType


CaptureRunner = Callable[
    [list[ViewConfig], str | Path, str | Path | None, float, bool, bool, object | None],
    int,
]


def _default_capture_runner(
    configs: list[ViewConfig],
    output_dir: str | Path,
    sap_dll_path: str | Path | None,
    render_delay: float,
    verbose: bool,
    ui_automation_enabled: bool = False,
    ui_stop_requested: object | None = None,
) -> int:
    main_module = sys.modules.get("main")
    if main_module and hasattr(main_module, "run_capture_configs"):
        return main_module.run_capture_configs(
            configs, output_dir, sap_dll_path, render_delay, verbose, ui_automation_enabled, ui_stop_requested
        )

    main_module = sys.modules.get("__main__")
    if main_module and hasattr(main_module, "run_capture_configs"):
        return main_module.run_capture_configs(
            configs, output_dir, sap_dll_path, render_delay, verbose, ui_automation_enabled, ui_stop_requested
        )

    from main import run_capture_configs

    return run_capture_configs(
        configs,
        output_dir,
        sap_dll_path,
        render_delay,
        verbose,
        ui_automation_enabled,
        ui_stop_requested,
    )


APP_VERSION = "dev-unknown"


class _QueueLogHandler(logging.Handler):
    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record: logging.LogRecord) -> None:
        try:
            message = self.format(record)
        except Exception:
            message = record.getMessage()
        self.log_queue.put(("log", record.levelname, message))


class SapCaptureGui:
    def __init__(
        self,
        root: tk.Tk,
        plan_path: str | None = None,
        output_dir: str | None = None,
        sap_dll_path: str | None = None,
        render_delay: float = 0.5,
        verbose: bool = False,
        capture_runner: CaptureRunner | None = None,
        app_version: str = APP_VERSION,
    ):
        self.root = root
        self.root.title("SAP2000 Capture")
        self.root.geometry("1460x860")
        self.root.minsize(1280, 760)

        self.log_queue: queue.Queue = queue.Queue()
        self.log_handler = _QueueLogHandler(self.log_queue)
        self.log_handler.setFormatter(
            logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", "%H:%M:%S")
        )
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.DEBUG if verbose else logging.INFO)

        self.is_closing = False
        self.worker_thread: threading.Thread | None = None
        self.ui_stop_event = threading.Event()
        self.capture_runner = capture_runner or _default_capture_runner
        self.app_version = app_version
        self.catalog: dict[str, list[str]] = {"load_patterns": [], "load_cases": [], "combos": []}
        self.configs: list[ViewConfig] = []
        self.selected_index: int | None = None
        self.current_plan_path: Path | None = Path(plan_path) if plan_path else None
        default_sap_dll = Path(sap_dll_path) if sap_dll_path else SapBridge.find_default_dll_path()

        self.sap_dll_var = tk.StringVar(value=str(default_sap_dll) if default_sap_dll else "")
        self.output_dir_var = tk.StringVar(value=output_dir or str(Path.cwd() / "outputs"))
        self.render_delay_var = tk.StringVar(value=str(render_delay))
        self.verbose_var = tk.BooleanVar(value=verbose)
        self.ui_automation_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Listo. Conecta SAP2000 para cargar patrones y casos.")
        self.catalog_var = tk.StringVar(value="Sin catálogo cargado")
        self.plan_var = tk.StringVar(value=str(self.current_plan_path) if self.current_plan_path else "")
        self.runtime_var = tk.StringVar(value=f"Runtime: {self.app_version}")

        self.filename_var = tk.StringVar()
        self.description_var = tk.StringVar()
        self.view_type_var = tk.StringVar(value=ViewType.ISO_3D.name)
        self.display_type_var = tk.StringVar(value=DisplayType.GEOMETRY_ONLY.name)
        self.case_name_var = tk.StringVar()
        self.mode_number_var = tk.StringVar(value="1")
        self.window_number_var = tk.StringVar(value="0")
        self.item_delay_var = tk.StringVar(value="0.0")

        self._build_ui()
        logging.info("Interfaz SAP2000 Capture abierta")
        logging.info("SAP2000 Capture runtime: %s", self.app_version)
        if self.sap_dll_var.get():
            logging.info("SAP2000 DLL por defecto: %s", self.sap_dll_var.get())
        else:
            logging.info("SAP2000 DLL por defecto no encontrada; usa Buscar para seleccionarla")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._poll_queue)

        if self.current_plan_path and self.current_plan_path.exists():
            self._load_plan_file(self.current_plan_path)
        self._sync_case_controls()

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        container = ttk.Frame(self.root, padding=12)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=0)
        container.rowconfigure(2, weight=1)

        main_panel = ttk.Frame(container)
        main_panel.grid(row=0, column=0, rowspan=4, sticky="nsew", padx=(0, 12))
        main_panel.columnconfigure(0, weight=1)
        main_panel.rowconfigure(2, weight=1)

        side_panel = ttk.Frame(container)
        side_panel.grid(row=0, column=1, rowspan=4, sticky="ns")
        side_panel.columnconfigure(0, weight=1)
        side_panel.rowconfigure(1, weight=1)

        top = ttk.Frame(main_panel)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        top.columnconfigure(1, weight=1)

        conn = ttk.LabelFrame(top, text="Conexión y salida", padding=10)
        conn.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        conn.columnconfigure(1, weight=1)

        ttk.Label(conn, text="SAP2000 DLL").grid(row=0, column=0, sticky="w")
        ttk.Entry(conn, textvariable=self.sap_dll_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(conn, text="Buscar", command=self._browse_dll).grid(row=0, column=2, sticky="ew")

        ttk.Label(conn, text="Carpeta salida").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(conn, textvariable=self.output_dir_var).grid(row=1, column=1, sticky="ew", padx=6, pady=(8, 0))
        ttk.Button(conn, text="Elegir", command=self._browse_output_dir).grid(row=1, column=2, sticky="ew", pady=(8, 0))

        ttk.Label(conn, text="Render delay base").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(conn, textvariable=self.render_delay_var, width=12).grid(row=2, column=1, sticky="w", padx=6, pady=(8, 0))
        ttk.Checkbutton(conn, text="Verbose", variable=self.verbose_var).grid(row=2, column=2, sticky="e", pady=(8, 0))

        ttk.Checkbutton(
            conn,
            text="Permitir automatizacion de teclado",
            variable=self.ui_automation_var,
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Button(conn, text="Detener UI", command=self._stop_ui_automation).grid(
            row=3, column=2, sticky="ew", pady=(8, 0)
        )

        actions = ttk.Frame(conn)
        actions.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        actions.columnconfigure(2, weight=1)
        ttk.Button(actions, text="Conectar SAP2000", command=self._connect_sap).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(actions, text="Guardar plan JSON", command=self._save_plan_dialog).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(actions, text="Ejecutar capturas", command=self._run_capture).grid(row=0, column=2, sticky="ew", padx=(4, 0))

        ttk.Label(conn, textvariable=self.catalog_var, foreground="#555555").grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(10, 0)
        )

        plan_box = ttk.LabelFrame(top, text="Plan", padding=10)
        plan_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        plan_box.columnconfigure(1, weight=1)

        ttk.Label(plan_box, text="Archivo").grid(row=0, column=0, sticky="w")
        ttk.Entry(plan_box, textvariable=self.plan_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(plan_box, text="Abrir", command=self._load_plan_dialog).grid(row=0, column=2, sticky="ew")

        ttk.Label(
            plan_box,
            text="Puedes armar el plan desde esta GUI sin crear un JSON primero. Guardarlo es opcional.",
            wraplength=430,
            foreground="#555555",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 0))

        editor = ttk.LabelFrame(main_panel, text="Definición de captura", padding=10)
        editor.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        for idx in range(7):
            editor.columnconfigure(idx, weight=1)

        ttk.Label(editor, text="Nombre archivo").grid(row=0, column=0, sticky="w")
        ttk.Entry(editor, textvariable=self.filename_var).grid(row=1, column=0, sticky="ew", padx=(0, 6))

        ttk.Label(editor, text="Descripción").grid(row=0, column=1, sticky="w")
        ttk.Entry(editor, textvariable=self.description_var).grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(editor, text="Vista").grid(row=0, column=2, sticky="w")
        self.view_combo = ttk.Combobox(
            editor,
            textvariable=self.view_type_var,
            values=[item.name for item in ViewType],
            state="readonly",
        )
        self.view_combo.grid(row=1, column=2, sticky="ew", padx=6)

        ttk.Label(editor, text="Display").grid(row=0, column=3, sticky="w")
        self.display_combo = ttk.Combobox(
            editor,
            textvariable=self.display_type_var,
            values=[item.name for item in DisplayType],
            state="readonly",
        )
        self.display_combo.grid(row=1, column=3, sticky="ew", padx=6)
        self.display_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_case_controls())

        ttk.Label(editor, text="Caso / patrón / combo").grid(row=0, column=4, sticky="w")
        self.case_selector_frame = ttk.Frame(editor)
        self.case_selector_frame.grid(row=1, column=4, sticky="ew", padx=6)
        self.case_selector_frame.columnconfigure(0, weight=1)

        self.case_combo = ttk.Combobox(
            self.case_selector_frame,
            textvariable=self.case_name_var,
            state="readonly",
        )
        self.case_combo.grid(row=0, column=0, sticky="ew")
        self.case_combo.bind("<<ComboboxSelected>>", self._on_case_combo_selected)

        self.load_pattern_listbox = tk.Listbox(
            self.case_selector_frame,
            height=6,
            exportselection=False,
        )
        self.load_pattern_listbox.grid(row=0, column=0, sticky="nsew")
        self.load_pattern_listbox.bind("<<ListboxSelect>>", self._on_load_pattern_select)
        self.load_pattern_scroll = ttk.Scrollbar(
            self.case_selector_frame,
            orient="vertical",
            command=self.load_pattern_listbox.yview,
        )
        self.load_pattern_listbox.configure(yscrollcommand=self.load_pattern_scroll.set)
        self.load_pattern_scroll.grid(row=0, column=1, sticky="ns")
        self.load_pattern_listbox.grid_remove()
        self.load_pattern_scroll.grid_remove()

        ttk.Label(editor, text="Modo").grid(row=0, column=5, sticky="w")
        self.mode_spin = ttk.Spinbox(editor, from_=1, to=999, textvariable=self.mode_number_var, width=8)
        self.mode_spin.grid(row=1, column=5, sticky="w", padx=(6, 0))

        ttk.Label(editor, text="Ventana").grid(row=0, column=6, sticky="w")
        ttk.Spinbox(editor, from_=0, to=99, textvariable=self.window_number_var, width=8).grid(row=1, column=6, sticky="w", padx=(6, 0))

        ttk.Label(editor, text="Delay extra").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(editor, textvariable=self.item_delay_var, width=12).grid(row=3, column=0, sticky="w", padx=(0, 6))

        form_actions = ttk.Frame(editor)
        form_actions.grid(row=3, column=1, columnspan=6, sticky="e")
        ttk.Button(form_actions, text="Limpiar", command=self._clear_form).grid(row=0, column=0, padx=(0, 4))
        ttk.Button(form_actions, text="Actualizar fila", command=self._update_selected,).grid(row=0, column=1, padx=4)
        ttk.Button(form_actions, text="Agregar fila", command=self._add_config).grid(row=0, column=2, padx=(4, 0))

        list_box = ttk.LabelFrame(main_panel, text="Capturas programadas", padding=10)
        list_box.grid(row=2, column=0, sticky="nsew")
        list_box.columnconfigure(0, weight=1)
        list_box.rowconfigure(0, weight=1)

        columns = ("idx", "filename", "view", "display", "case", "mode", "window", "delay", "description")
        self.tree = ttk.Treeview(list_box, columns=columns, show="headings", height=16)
        headings = {
            "idx": "#",
            "filename": "Archivo",
            "view": "Vista",
            "display": "Display",
            "case": "Caso/Patrón/Combo",
            "mode": "Modo",
            "window": "Ventana",
            "delay": "Delay",
            "description": "Descripción",
        }
        widths = {"idx": 50, "filename": 180, "view": 90, "display": 120, "case": 180, "mode": 55, "window": 65, "delay": 65, "description": 260}
        for key in columns:
            self.tree.heading(key, text=headings[key])
            self.tree.column(key, width=widths[key], anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        tree_scroll = ttk.Scrollbar(list_box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.grid(row=0, column=1, sticky="ns")

        list_actions = ttk.Frame(list_box)
        list_actions.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        for idx in range(5):
            list_actions.columnconfigure(idx, weight=1)
        ttk.Button(list_actions, text="Subir", command=lambda: self._move_selected(-1)).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(list_actions, text="Bajar", command=lambda: self._move_selected(1)).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(list_actions, text="Duplicar", command=self._duplicate_selected).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(list_actions, text="Eliminar", command=self._delete_selected).grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Button(list_actions, text="Nuevo plan", command=self._reset_plan).grid(row=0, column=4, sticky="ew", padx=(4, 0))

        runtime_box = ttk.LabelFrame(side_panel, text="Sesión", padding=10)
        runtime_box.grid(row=0, column=0, sticky="ew")
        runtime_box.columnconfigure(0, weight=1)
        ttk.Label(runtime_box, textvariable=self.runtime_var, foreground="#555555").grid(
            row=0, column=0, sticky="w"
        )

        log_box = ttk.LabelFrame(side_panel, text="Log", padding=10)
        log_box.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)

        self.log_text = ScrolledText(log_box, wrap="word", width=42, state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        status = ttk.Label(main_panel, textvariable=self.status_var)
        status.grid(row=3, column=0, sticky="ew", pady=(8, 0))

    def _browse_dll(self) -> None:
        current = self.sap_dll_var.get().strip()
        initialdir = str(Path(current).parent) if current else None
        path = filedialog.askopenfilename(
            title="Seleccionar SAP2000v1.dll",
            initialdir=initialdir,
            filetypes=[("SAP2000 DLL", "*.dll"), ("Todos los archivos", "*.*")],
        )
        if path:
            self.sap_dll_var.set(path)
            logging.info("SAP2000 DLL seleccionada manualmente: %s", path)

    def _browse_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if path:
            self.output_dir_var.set(path)

    def _append_log(self, level: str, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _poll_queue(self) -> None:
        if self.is_closing:
            return
        try:
            while True:
                kind, *payload = self.log_queue.get_nowait()
                if kind == "log":
                    level, message = payload
                    self._append_log(level, message)
                elif kind == "catalog":
                    catalog = payload[0]
                    catalog = {key: self._unique_options(value) for key, value in catalog.items()}
                    self.catalog = catalog
                    logging.info(
                        "Catálogo SAP2000 cargado: %s patrones, %s casos, %s combos",
                        len(catalog["load_patterns"]),
                        len(catalog["load_cases"]),
                        len(catalog["combos"]),
                    )
                    self.catalog_var.set(
                        f"Patrones: {len(catalog['load_patterns'])} | "
                        f"Casos: {len(catalog['load_cases'])} | "
                        f"Combos: {len(catalog['combos'])}"
                    )
                    self.status_var.set("Catálogo SAP2000 cargado.")
                    self._sync_case_controls()
                    self.worker_thread = None
                elif kind == "status":
                    self.status_var.set(payload[0])
                elif kind == "error":
                    self.status_var.set(payload[0])
                    messagebox.showerror("SAP2000 Capture", payload[0])
                    self.worker_thread = None
                elif kind == "run_done":
                    exit_code = payload[0]
                    self.status_var.set("Capturas completadas." if exit_code == 0 else "Capturas finalizadas con errores.")
                    if exit_code == 0:
                        messagebox.showinfo("SAP2000 Capture", "Capturas completadas.")
                    else:
                        messagebox.showwarning("SAP2000 Capture", "La ejecución terminó con errores. Revisa el log.")
                    self.worker_thread = None
        except queue.Empty:
            pass
        self.root.after(120, self._poll_queue)

    def _start_worker(self, target, busy_text: str) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("SAP2000 Capture", "Ya hay una operación en curso.")
            return
        self.status_var.set(busy_text)
        self.worker_thread = threading.Thread(target=target, daemon=True)
        self.worker_thread.start()

    def _stop_ui_automation(self) -> None:
        self.ui_stop_event.set()
        self.status_var.set("Detencion UI solicitada. La automatizacion abortara en el siguiente paso.")
        logging.warning("Detencion manual de automatizacion UI solicitada.")

    def _connect_sap(self) -> None:
        sap_dll = self.sap_dll_var.get().strip() or None
        logging.info("Conectando a SAP2000 desde la GUI")
        if sap_dll:
            logging.info("Usando SAP2000 DLL: %s", sap_dll)
        else:
            logging.info("Sin DLL explícita; se intentará conexión COM directa")

        def worker() -> None:
            try:
                bridge = SapBridge(sap_dll_path=sap_dll)
                bridge.connect()
                try:
                    catalog = bridge.get_model_catalog()
                finally:
                    bridge.disconnect()
                self.log_queue.put(("catalog", catalog))
            except Exception as exc:
                self.log_queue.put(("error", f"No se pudo cargar el catálogo SAP2000: {exc}"))

        self._start_worker(worker, "Conectando a SAP2000 y leyendo catálogo...")

    @staticmethod
    def _unique_options(values: list[str]) -> list[str]:
        options: list[str] = []
        seen: set[str] = set()
        for value in values:
            for name in SapBridge._split_name_text(str(value)):
                if name not in seen:
                    options.append(name)
                    seen.add(name)
        return options

    def _catalog_values_for_display(self, display_name: str) -> list[str]:
        if display_name == DisplayType.LOAD_PATTERN.name:
            return self._unique_options(self.catalog.get("load_patterns", []))
        if display_name in {
            DisplayType.LOAD_CASE.name,
            DisplayType.DEFORMED.name,
            DisplayType.FRAME_FORCES.name,
        }:
            return sorted(
                self._unique_options(
                    self.catalog.get("load_cases", []) + self.catalog.get("combos", [])
                )
            )
        if display_name == DisplayType.MODE_SHAPE.name:
            return self._unique_options(self.catalog.get("load_cases", []))
        return []

    def _sync_case_controls(self) -> None:
        display_name = self.display_type_var.get()
        needs_case = display_name != DisplayType.GEOMETRY_ONLY.name
        options = self._catalog_values_for_display(display_name)

        if not needs_case:
            self.case_name_var.set("")
            self.case_combo.grid_remove()
            self.load_pattern_listbox.grid_remove()
            self.load_pattern_scroll.grid_remove()
            self.case_combo.configure(values=[], state="disabled")
            self.case_combo.set("")
        else:
            if display_name == DisplayType.LOAD_PATTERN.name:
                self.case_combo.grid_remove()
                self.load_pattern_listbox.grid(row=0, column=0, sticky="nsew")
                self.load_pattern_scroll.grid(row=0, column=1, sticky="ns")
                self._refresh_load_pattern_selector()
                if options:
                    self.load_pattern_listbox.configure(state="normal")
                else:
                    self.load_pattern_listbox.configure(state="disabled")
            else:
                self.load_pattern_listbox.grid_remove()
                self.load_pattern_scroll.grid_remove()
                self.case_combo.grid(row=0, column=0, sticky="ew")
                if options:
                    self.case_combo.configure(values=options, state="readonly")
                    if self.case_name_var.get() not in options:
                        self.case_name_var.set(options[0])
                else:
                    self.case_combo.configure(values=[], state="normal")
                    self.case_combo.set("")
                if self.case_name_var.get() in options:
                    self.case_combo.set(self.case_name_var.get())

        if display_name == DisplayType.MODE_SHAPE.name:
            self.mode_spin.configure(state="normal")
        else:
            self.mode_number_var.set("1")
            self.mode_spin.configure(state="disabled")

    def _refresh_load_pattern_selector(self) -> None:
        options = self.catalog.get("load_patterns", [])
        current = self.case_name_var.get().strip()

        self.load_pattern_listbox.delete(0, tk.END)
        for option in options:
            self.load_pattern_listbox.insert(tk.END, option)

        if current in options:
            index = options.index(current)
            self.load_pattern_listbox.selection_set(index)
            self.load_pattern_listbox.see(index)
        elif options:
            self.case_name_var.set(options[0])
            self.load_pattern_listbox.selection_set(0)
            self.load_pattern_listbox.see(0)
        else:
            self.case_name_var.set("")
            self.load_pattern_listbox.selection_clear(0, tk.END)

    def _on_case_combo_selected(self, _event=None) -> None:
        self.case_name_var.set(self.case_combo.get().strip())

    def _on_load_pattern_select(self, _event=None) -> None:
        selected = self.load_pattern_listbox.curselection()
        if not selected:
            return
        self.case_name_var.set(self.load_pattern_listbox.get(selected[0]))

    def _config_from_form(self) -> ViewConfig:
        filename = self.filename_var.get().strip()
        if not filename:
            raise ValueError("El nombre de archivo es obligatorio.")

        display_type = DisplayType[self.display_type_var.get()]
        case_name = self.case_name_var.get().strip()
        if display_type != DisplayType.GEOMETRY_ONLY and not case_name:
            raise ValueError("Debes indicar el caso, patrón o combinación para este display.")

        mode_number = 1
        if display_type == DisplayType.MODE_SHAPE:
            try:
                mode_number = int(self.mode_number_var.get())
            except ValueError as exc:
                raise ValueError("El número de modo debe ser entero.") from exc

        try:
            render_delay = float(self.item_delay_var.get().strip() or "0.0")
        except ValueError as exc:
            raise ValueError("El delay extra debe ser numérico.") from exc
        try:
            window_number = int(self.window_number_var.get().strip() or "0")
        except ValueError as exc:
            raise ValueError("El número de ventana debe ser entero.") from exc

        return ViewConfig(
            filename=filename,
            description=self.description_var.get().strip(),
            view_type=ViewType[self.view_type_var.get()],
            display_type=display_type,
            case_name=case_name,
            mode_number=mode_number,
            window_number=window_number,
            render_delay=render_delay,
        )

    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, cfg in enumerate(self.configs, start=1):
            self.tree.insert(
                "",
                "end",
                iid=str(idx - 1),
                values=(
                    idx,
                    cfg.filename,
                    cfg.view_type.name,
                    cfg.display_type.name,
                    cfg.case_name,
                    cfg.mode_number,
                    cfg.window_number,
                    cfg.render_delay,
                    cfg.description,
                ),
            )

    def _clear_form(self) -> None:
        self.selected_index = None
        self.filename_var.set("")
        self.description_var.set("")
        self.view_type_var.set(ViewType.ISO_3D.name)
        self.display_type_var.set(DisplayType.GEOMETRY_ONLY.name)
        self.case_name_var.set("")
        self.mode_number_var.set("1")
        self.window_number_var.set("0")
        self.item_delay_var.set("0.0")
        self.tree.selection_remove(self.tree.selection())
        self.load_pattern_listbox.selection_clear(0, tk.END)
        self._sync_case_controls()

    def _add_config(self) -> None:
        try:
            cfg = self._config_from_form()
        except Exception as exc:
            messagebox.showerror("SAP2000 Capture", str(exc))
            return
        self.configs.append(cfg)
        self._refresh_tree()
        self._clear_form()
        self.status_var.set(f"Fila agregada. Total: {len(self.configs)}")

    def _update_selected(self) -> None:
        if self.selected_index is None:
            messagebox.showwarning("SAP2000 Capture", "Selecciona una fila para actualizar.")
            return
        try:
            cfg = self._config_from_form()
        except Exception as exc:
            messagebox.showerror("SAP2000 Capture", str(exc))
            return
        self.configs[self.selected_index] = cfg
        self._refresh_tree()
        self.tree.selection_set(str(self.selected_index))
        self.status_var.set("Fila actualizada.")

    def _on_tree_select(self, _event=None) -> None:
        selected = self.tree.selection()
        if not selected:
            return
        index = int(selected[0])
        cfg = self.configs[index]
        self.selected_index = index
        self.filename_var.set(cfg.filename)
        self.description_var.set(cfg.description)
        self.view_type_var.set(cfg.view_type.name)
        self.display_type_var.set(cfg.display_type.name)
        self.case_name_var.set(cfg.case_name)
        self.mode_number_var.set(str(cfg.mode_number))
        self.window_number_var.set(str(cfg.window_number))
        self.item_delay_var.set(str(cfg.render_delay))
        self._sync_case_controls()

    def _delete_selected(self) -> None:
        if self.selected_index is None:
            messagebox.showwarning("SAP2000 Capture", "Selecciona una fila para eliminar.")
            return
        del self.configs[self.selected_index]
        self._refresh_tree()
        self._clear_form()
        self.status_var.set(f"Fila eliminada. Total: {len(self.configs)}")

    def _duplicate_selected(self) -> None:
        if self.selected_index is None:
            messagebox.showwarning("SAP2000 Capture", "Selecciona una fila para duplicar.")
            return
        cfg = self.configs[self.selected_index]
        duplicate = ViewConfig(
            filename=f"{cfg.filename}_copy",
            view_type=cfg.view_type,
            display_type=cfg.display_type,
            case_name=cfg.case_name,
            mode_number=cfg.mode_number,
            window_number=cfg.window_number,
            render_delay=cfg.render_delay,
            description=cfg.description,
        )
        self.configs.insert(self.selected_index + 1, duplicate)
        self._refresh_tree()
        self.status_var.set("Fila duplicada.")

    def _move_selected(self, offset: int) -> None:
        if self.selected_index is None:
            messagebox.showwarning("SAP2000 Capture", "Selecciona una fila para mover.")
            return
        new_index = self.selected_index + offset
        if new_index < 0 or new_index >= len(self.configs):
            return
        self.configs[self.selected_index], self.configs[new_index] = (
            self.configs[new_index],
            self.configs[self.selected_index],
        )
        self.selected_index = new_index
        self._refresh_tree()
        self.tree.selection_set(str(new_index))

    def _reset_plan(self) -> None:
        self.configs.clear()
        self.current_plan_path = None
        self.plan_var.set("")
        self._refresh_tree()
        self._clear_form()
        self.status_var.set("Plan reiniciado.")

    def _load_plan_file(self, path: str | Path) -> None:
        configs = load_plan(path)
        self.configs = list(configs)
        self.current_plan_path = Path(path)
        self.plan_var.set(str(self.current_plan_path))
        self._refresh_tree()
        self._clear_form()
        self.status_var.set(f"Plan cargado: {self.current_plan_path.name}")

    def _load_plan_dialog(self) -> None:
        path = filedialog.askopenfilename(
            title="Abrir plan",
            filetypes=[
                ("Planes compatibles", "*.json *.xlsx *.xlsm"),
                ("JSON", "*.json"),
                ("Excel", "*.xlsx *.xlsm"),
            ],
        )
        if not path:
            return
        try:
            self._load_plan_file(path)
        except Exception as exc:
            messagebox.showerror("SAP2000 Capture", f"No se pudo cargar el plan: {exc}")

    def _save_plan_dialog(self) -> None:
        if not self.configs:
            messagebox.showwarning("SAP2000 Capture", "No hay capturas programadas para guardar.")
            return
        initial = self.current_plan_path.name if self.current_plan_path else "capture_plan.json"
        path = filedialog.asksaveasfilename(
            title="Guardar plan JSON",
            defaultextension=".json",
            initialfile=initial,
            filetypes=[("JSON", "*.json")],
        )
        if not path:
            return
        try:
            saved = save_plan(self.configs, path)
        except Exception as exc:
            messagebox.showerror("SAP2000 Capture", f"No se pudo guardar el plan: {exc}")
            return
        self.current_plan_path = saved
        self.plan_var.set(str(saved))
        self.status_var.set(f"Plan guardado en {saved}")

    def _run_capture(self) -> None:
        if not self.configs:
            messagebox.showwarning("SAP2000 Capture", "No hay capturas programadas.")
            return

        try:
            render_delay = float(self.render_delay_var.get().strip() or "0.5")
        except ValueError:
            messagebox.showerror("SAP2000 Capture", "El render delay base debe ser numérico.")
            return

        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            messagebox.showerror("SAP2000 Capture", "Debes indicar una carpeta de salida.")
            return

        configs = list(self.configs)
        sap_dll = self.sap_dll_var.get().strip() or None
        verbose = self.verbose_var.get()
        ui_automation_enabled = self.ui_automation_var.get()
        self.ui_stop_event.clear()

        def worker() -> None:
            try:
                exit_code = self.capture_runner(
                    configs=configs,
                    output_dir=output_dir,
                    sap_dll_path=sap_dll,
                    render_delay=render_delay,
                    verbose=verbose,
                    ui_automation_enabled=ui_automation_enabled,
                    ui_stop_requested=self.ui_stop_event,
                )
                self.log_queue.put(("run_done", exit_code))
            except Exception as exc:
                self.log_queue.put(("error", f"Error ejecutando capturas: {exc}"))

        self._start_worker(worker, "Ejecutando capturas...")

    def _on_close(self) -> None:
        self.is_closing = True
        logging.getLogger().removeHandler(self.log_handler)
        self.log_handler.close()
        self.root.destroy()


def launch_gui_app(
    argv: list[str] | None = None,
    capture_runner: CaptureRunner | None = None,
    app_version: str = APP_VERSION,
) -> None:
    parser = argparse.ArgumentParser(
        prog="sap_capture_gui",
        description="Interfaz gráfica para definir y ejecutar capturas SAP2000",
    )
    parser.add_argument("--plan")
    parser.add_argument("--output")
    parser.add_argument("--sap-dll")
    parser.add_argument("--render-delay", type=float, default=0.5)
    parser.add_argument("--verbose", action="store_true")
    args, _unknown = parser.parse_known_args(argv)

    root = tk.Tk()
    SapCaptureGui(
        root,
        plan_path=args.plan,
        output_dir=args.output,
        sap_dll_path=args.sap_dll,
        render_delay=args.render_delay,
        verbose=args.verbose,
        capture_runner=capture_runner,
        app_version=app_version,
    )
    root.mainloop()


if __name__ == "__main__":
    launch_gui_app()
