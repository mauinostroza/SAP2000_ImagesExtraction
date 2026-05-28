"""
main.py
Orquestador principal del script de captura SAP2000.

Uso:
  python main.py --plan capture_plan.json --output outputs/proyecto_01
  python main.py --plan capture_plan.json --output outputs --sap-dll "C:/SAP2000 23/SAP2000v1.dll"
  python main.py --generate-plan               ← solo genera capture_plan.json de ejemplo
  python main.py --list-cases                  ← lista casos/combos disponibles en el modelo

Dependencias:
  pip install pywin32 Pillow comtypes openpyxl
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from pathlib import Path

from capture_plan import load_plan, generate_sample_plan
from output_writer import OutputWriter
from sap_bridge import SapBridge
from view_controller import ViewController
from win32_capture import Win32CaptureEngine, find_sap2000_hwnd


# ── Logging ───────────────────────────────────────────────────────────────────

def setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    fmt   = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    logging.basicConfig(level=level, format=fmt, datefmt="%H:%M:%S")


# ── Comando: listar casos ──────────────────────────────────────────────────────

def cmd_list_cases(bridge: SapBridge) -> None:
    print("\n── Patrones de carga ─────────────────────────────")
    for name in bridge.get_load_pattern_names():
        print(f"  LOAD_PATTERN  {name}")

    print("\n── Casos de análisis ─────────────────────────────")
    for name in bridge.get_load_case_names():
        print(f"  LOAD_CASE     {name}")

    print("\n── Combinaciones ─────────────────────────────────")
    for name in bridge.get_combo_names():
        print(f"  COMBO         {name}")
    print()


# ── Pipeline de captura ────────────────────────────────────────────────────────

def run_capture_pipeline(
    plan_path:    str | Path,
    output_dir:   str | Path,
    sap_dll_path: str | Path | None,
    render_delay: float,
    verbose:      bool,
) -> int:
    """Ejecuta el pipeline completo. Retorna 0 si todo OK, 1 si hubo errores."""

    setup_logging(verbose)
    log = logging.getLogger("main")

    # 1. Cargar plan
    log.info(f"Cargando plan: {plan_path}")
    try:
        configs = load_plan(plan_path)
    except Exception as e:
        log.error(f"No se pudo cargar el plan: {e}")
        return 1

    if not configs:
        log.error("El plan está vacío.")
        return 1

    log.info(f"{len(configs)} capturas en el plan")

    # 2. Localizar ventana SAP2000 (win32)
    hwnd = find_sap2000_hwnd()
    if hwnd is None:
        log.error(
            "SAP2000 no está abierto o no se encontró la ventana. "
            "Abre SAP2000 y carga el modelo antes de ejecutar este script."
        )
        return 1
    log.info(f"Ventana SAP2000 encontrada: hwnd={hwnd}")

    # 3. Conectar COM
    log.info("Conectando a SAP2000 via COM...")
    bridge = SapBridge(sap_dll_path=sap_dll_path)
    try:
        bridge.connect()
    except Exception as e:
        log.error(f"Conexión COM fallida: {e}")
        return 1

    # 4. Inicializar módulos
    view_ctrl = ViewController(bridge.model, base_render_delay=render_delay)
    capture   = Win32CaptureEngine(render_delay=0.0)  # el delay lo maneja ViewController
    writer    = OutputWriter(output_dir)
    writer.start_session()

    # 5. Loop de capturas
    errors = 0
    for idx, cfg in enumerate(configs, start=1):
        desc = cfg.description or f"{cfg.view_type.name} / {cfg.display_type.name}"
        log.info(f"[{idx:03d}/{len(configs):03d}] {cfg.filename}  —  {desc}")

        t0 = time.perf_counter()
        try:
            # a) Aplicar vista (cambia caso, ángulo, zoom, espera render)
            view_ctrl.apply(cfg)

            # b) Capturar ventana SAP2000
            out_path = writer.get_output_path(cfg, idx)
            capture.capture(hwnd, out_path)

            elapsed = int((time.perf_counter() - t0) * 1000)
            writer.record_ok(cfg, idx, out_path, elapsed_ms=elapsed)

        except Exception as e:
            writer.record_error(cfg, idx, e)
            errors += 1

    # 6. Log y resumen
    writer.save_log()
    writer.print_summary()
    bridge.disconnect()

    return 0 if errors == 0 else 1


# ── CLI ────────────────────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="sap_capture",
        description="Captura autónoma de vistas SAP2000 usando PrintWindow (win32)",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    p.add_argument(
        "--plan", "-p",
        default="capture_plan.json",
        help="Ruta al plan de capturas (.json o .xlsx)\n(default: capture_plan.json)",
    )
    p.add_argument(
        "--output", "-o",
        default="outputs",
        help="Carpeta de salida para los PNG y el log\n(default: outputs/)",
    )
    p.add_argument(
        "--sap-dll",
        default=None,
        help="Ruta explícita a SAP2000v1.dll\n(se detecta automáticamente si se omite)",
    )
    p.add_argument(
        "--render-delay",
        type=float,
        default=0.5,
        help="Segundos de espera tras cambiar vista (default: 0.5)",
    )
    p.add_argument(
        "--generate-plan",
        action="store_true",
        help="Genera un capture_plan.json de ejemplo y termina",
    )
    p.add_argument(
        "--list-cases",
        action="store_true",
        help="Lista casos/combos disponibles en el modelo SAP2000 y termina",
    )
    p.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Logging detallado (DEBUG)",
    )
    return p


def main() -> None:
    parser = build_parser()
    args   = parser.parse_args()

    setup_logging(args.verbose)
    log = logging.getLogger("main")

    # ── Modo: generar plan de ejemplo ─────────────────────────────────────────
    if args.generate_plan:
        out = generate_sample_plan("capture_plan.json")
        print(f"Plan de ejemplo generado: {out}")
        print("Edita los 'case_name' con los nombres reales de tu modelo SAP2000.")
        print("Luego ejecuta: python main.py --list-cases  para verificarlos.")
        sys.exit(0)

    # ── Modo: listar casos ────────────────────────────────────────────────────
    if args.list_cases:
        bridge = SapBridge(sap_dll_path=args.sap_dll)
        try:
            bridge.connect()
        except Exception as e:
            log.error(f"No se pudo conectar a SAP2000: {e}")
            sys.exit(1)
        cmd_list_cases(bridge)
        bridge.disconnect()
        sys.exit(0)

    # ── Modo: captura ─────────────────────────────────────────────────────────
    ret = run_capture_pipeline(
        plan_path    = args.plan,
        output_dir   = args.output,
        sap_dll_path = args.sap_dll,
        render_delay = args.render_delay,
        verbose      = args.verbose,
    )
    sys.exit(ret)


if __name__ == "__main__":
    main()
