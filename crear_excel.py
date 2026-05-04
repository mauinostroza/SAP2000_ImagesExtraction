"""
crear_excel.py
Genera el archivo Excel de configuración SAP2000_Capturas.xlsx
Ejecutar una sola vez para crear la plantilla.

Uso:
    python crear_excel.py
    python crear_excel.py --ruta "C:/Proyectos/Mi_Proyecto/SAP2000_Capturas.xlsx"
"""
import sys
import os
import argparse

sys.path.insert(0, os.path.dirname(__file__))
from sap_imagenes import crear_excel_configuracion

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Crea el Excel de config SAP2000")
    parser.add_argument(
        "--ruta",
        default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "SAP2000_Capturas.xlsx"),
        help="Ruta completa del archivo Excel a crear"
    )
    args = parser.parse_args()

    crear_excel_configuracion(args.ruta)
    print(f"\nExcel creado en: {args.ruta}")
    print("\nPasos siguientes:")
    print("  1. Abre el archivo en Excel")
    print("  2. Completa la hoja CONFIG y la hoja CAPTURAS")
    print("  3. Abre SAP2000 con tu modelo cargado")
    print(f"  4. Ejecuta: python sap_imagenes.py --config \"{args.ruta}\"")
    print(f"     o usa: sap2000_capture.exe --config \"{args.ruta}\"")
