"""Entrypoint para el ejecutable portable de SAP2000 Image Capture."""

import sys

from sap_imagenes import main_cli


if __name__ == "__main__":
    sys.exit(main_cli())
