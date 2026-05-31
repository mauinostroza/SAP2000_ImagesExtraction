"""
main.py
Orquestador principal del script de captura SAP2000.

Uso:
  python main.py --plan capture_plan.json --output outputs/proyecto_01
  python main.py --plan capture_plan.json --output outputs --sap-dll "C:/SAP2000 23/SAP2000v1.dll"
  python main.py --generate-plan
  python main.py --list-cases
"""

from __future__ import annotations

import argparse
import logging
import platform
import sys
import time
from pathlib import Path

from capture_plan import generate_sample_plan, load_plan
from output_writer import OutputWriter
from sap_bridge import SapBridge
from view_controller import ViewConfig
from view_controller import ViewController

APP_VERSION = "dev-2026-05-31-01"


def setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    fmt = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    logging.basicConfig(level=level, format=fmt, datefmt="%H:%M:%S")


def ensure_windows(required_for: str) -> None:
    if platform.system() != "Windows":
        raise RuntimeError(
            f"{required_for} solo está soportado en Windows porque depende de Win32 y SAP2000 COM."
        )


def cmd_list_cases(bridge: SapBridge) -> None:
    print("\n-- Patrones de carga --")
    for name in bridge.get_load_pattern_names():
        print(f"  LOAD_PATTERN  {name}")

    print("\n-- Casos de análisis --")
    for name in bridge.get_load_case_names():
        print(f"  LOAD_CASE     {name}")

    print("\n-- Combinaciones --")
    for name in bridge.get_combo_names():
        print(f"  COMBO         {name}")
    print()


def launch_gui(argv: list[str] | None = None) -> int:
    from sap2000_gui import launch_gui_app

    launch_gui_app(argv, capture_runner=run_capture_configs, app_version=APP_VERSION)
    return 0


def run_capture_configs(
    configs: list[ViewConfig],
    output_dir: str | Path,
    sap_dll_path: str | Path | None,
    render_delay: float,
    verbose: bool,
    ui_automation_enabled: bool = False,
    ui_stop_requested=None,
) -> int:
    setup_logging(verbose)
    log = logging.getLogger("main")
    log.info("SAP2000 Capture runtime: %s", APP_VERSION)
    log.info(
        "UI automation: %s",
        "armada" if ui_automation_enabled else "desarmada",
    )

    if not configs:
        log.error("El plan está vacío.")
        return 1

    ensure_windows("La captura")
    from sap_ui_automation import SAP2000UIController
    from win32_capture import Win32CaptureEngine, find_sap2000_hwnd, prepare_window_for_capture

    hwnd = find_sap2000_hwnd()
    if hwnd is None:
        log.error(
            "SAP2000 no está abierto o no se encontró la ventana. "
            "Abre SAP2000 y carga el modelo antes de ejecutar este script."
        )
        return 1
    log.info("Ventana SAP2000 encontrada: hwnd=%s", hwnd)
    prepare_window_for_capture(hwnd)

    log.info("Conectando a SAP2000 via COM...")
    bridge = SapBridge(sap_dll_path=sap_dll_path)
    try:
        bridge.connect()
    except Exception as exc:
        log.error("Conexión COM fallida: %s", exc)
        return 1

    ui_ctrl = SAP2000UIController(
        hwnd,
        enabled=ui_automation_enabled,
        stop_requested=ui_stop_requested,
    )
    view_ctrl = ViewController(bridge.model, base_render_delay=render_delay, ui_controller=ui_ctrl)
    capture = Win32CaptureEngine(render_delay=0.0)
    writer = OutputWriter(output_dir)
    writer.start_session()

    errors = 0
    try:
        for idx, cfg in enumerate(configs, start=1):
            desc = cfg.description or f"{cfg.view_type.name} / {cfg.display_type.name}"
            log.info("[%03d/%03d] %s - %s", idx, len(configs), cfg.filename, desc)

            t0 = time.perf_counter()
            try:
                view_ctrl.apply(cfg)
                out_path = writer.get_output_path(cfg, idx)
                capture.capture(hwnd, out_path)
                elapsed = int((time.perf_counter() - t0) * 1000)
                writer.record_ok(cfg, idx, out_path, elapsed_ms=elapsed)
            except Exception as exc:
                writer.record_error(cfg, idx, exc)
                errors += 1
    finally:
        writer.save_log()
        writer.print_summary()
        bridge.disconnect()

    return 0 if errors == 0 else 1


def run_capture_pipeline(
    plan_path: str | Path,
    output_dir: str | Path,
    sap_dll_path: str | Path | None,
    render_delay: float,
    verbose: bool,
) -> int:
    setup_logging(verbose)
    log = logging.getLogger("main")

    log.info("Cargando plan: %s", plan_path)
    try:
        configs = load_plan(plan_path)
    except Exception as exc:
        log.error("No se pudo cargar el plan: %s", exc)
        return 1

    return run_capture_configs(
        configs=configs,
        output_dir=output_dir,
        sap_dll_path=sap_dll_path,
        render_delay=render_delay,
        verbose=verbose,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sap_capture",
        description="Captura autónoma de vistas SAP2000 usando PrintWindow (Win32)",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--plan",
        "-p",
        default="capture_plan.json",
        help="Ruta al plan de capturas (.json o .xlsx)\n(default: capture_plan.json)",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="outputs",
        help="Carpeta de salida para los PNG y el log\n(default: outputs/)",
    )
    parser.add_argument(
        "--sap-dll",
        default=None,
        help="Ruta explícita a SAP2000v1.dll\n(se detecta automáticamente si se omite)",
    )
    parser.add_argument(
        "--render-delay",
        type=float,
        default=0.5,
        help="Segundos de espera tras cambiar vista (default: 0.5)",
    )
    parser.add_argument(
        "--generate-plan",
        action="store_true",
        help="Genera un capture_plan.json de ejemplo y termina",
    )
    parser.add_argument(
        "--list-cases",
        action="store_true",
        help="Lista casos/combos disponibles en el modelo SAP2000 y termina",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Logging detallado (DEBUG)",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Abre la interfaz gráfica para armar y ejecutar capturas",
    )
    return parser


def _arg_was_provided(names: tuple[str, ...]) -> bool:
    for token in sys.argv[1:]:
        for name in names:
            if token == name or token.startswith(f"{name}="):
                return True
    return False


def main() -> None:
    if len(sys.argv) == 1:
        sys.exit(launch_gui([]))

    parser = build_parser()
    args = parser.parse_args()

    setup_logging(args.verbose)
    log = logging.getLogger("main")

    if args.gui:
        gui_argv: list[str] = []
        if _arg_was_provided(("--plan", "-p")):
            gui_argv.extend(["--plan", args.plan])
        if _arg_was_provided(("--output", "-o")):
            gui_argv.extend(["--output", args.output])
        if _arg_was_provided(("--sap-dll",)):
            gui_argv.extend(["--sap-dll", args.sap_dll])
        if _arg_was_provided(("--render-delay",)):
            gui_argv.extend(["--render-delay", str(args.render_delay)])
        if args.verbose:
            gui_argv.append("--verbose")
        sys.exit(launch_gui(gui_argv))

    if args.generate_plan:
        out = generate_sample_plan("capture_plan.json")
        print(f"Plan de ejemplo generado: {out}")
        print("Edita los 'case_name' con los nombres reales de tu modelo SAP2000.")
        print("Luego ejecuta: python main.py --list-cases")
        sys.exit(0)

    if args.list_cases:
        try:
            ensure_windows("La conexión SAP2000")
        except RuntimeError as exc:
            log.error("%s", exc)
            sys.exit(1)
        bridge = SapBridge(sap_dll_path=args.sap_dll)
        try:
            bridge.connect()
        except Exception as exc:
            log.error("No se pudo conectar a SAP2000: %s", exc)
            sys.exit(1)
        cmd_list_cases(bridge)
        bridge.disconnect()
        sys.exit(0)

    try:
        ret = run_capture_pipeline(
            plan_path=args.plan,
            output_dir=args.output,
            sap_dll_path=args.sap_dll,
            render_delay=args.render_delay,
            verbose=args.verbose,
        )
    except RuntimeError as exc:
        log.error("%s", exc)
        ret = 1
    sys.exit(ret)


if __name__ == "__main__":
    main()
