# Mejoras Pendientes de GUI y Flujo

## Hallazgos confirmados

1. Las filas con `ACTIVO = NO` se cuentan como error y hacen fallar la extraccion aunque solo debieran omitirse.
2. La GUI trata cualquier lote con errores parciales como fallo total, en vez de distinguir entre:
   - fallo de conexion o inicializacion,
   - extraccion completada con errores parciales,
   - extraccion completada sin errores.
3. El estado final de la GUI queda incoherente porque despues de extraer vuelve a mostrar `Conectado` y pisa `Listo`.
4. Un reconectado fallido no limpia la conexion previa y puede dejar habilitada una extraccion sobre un `sap_model` obsoleto.
5. `cargar_configuracion_desde_excel()` no cierra el workbook de `openpyxl`, con riesgo de handles abiertos en Windows.
6. El EXE con `--gui` sigue siendo un binario de consola por `console=True` en PyInstaller.
7. El launcher `.bat` distribuible no expone el flujo GUI.
8. `python sap2000_gui.py` no acepta argumentos mientras que el EXE si acepta `--config` y `--allow-unsafe-output`.

## Criterios de correccion

- Las filas inactivas deben quedar marcadas como omitidas, no como error.
- La GUI debe usar el resumen estructurado del backend y reflejar:
  - `Listo` si no hubo errores,
  - `Completado con errores` si hubo errores parciales,
  - `Error` solo si no se pudo ejecutar el trabajo.
- La GUI no debe conservar conexiones viejas despues de un intento de reconexion fallido.
- El backend debe cerrar los workbooks que abre.
- La documentacion y los entrypoints deben describir el flujo real entregado al usuario.

## Verificaciones minimas a mantener

- `python -m py_compile sap_imagenes.py sap2000_gui.py`
- `python sap_imagenes.py --help`
- Revision de consistencia entre:
  - `sap_imagenes.py`
  - `sap2000_gui.py`
  - `sap2000_portable.spec`
  - `ejecutar_capturas.bat`
  - `README.md`

## Tarea 3: entrypoints, empaquetado y documentacion

Alcance de esta iteracion:

- Mantener consistente el acceso a la GUI desde:
  - `sap2000_capture.exe`
  - `sap2000_capture_gui.exe`
  - `python sap2000_gui.py`
  - `ejecutar_capturas.bat --gui`
- Alinear los argumentos soportados entre EXE, Python y launcher.
- Documentar el comportamiento real del paquete portable, sin prometer un flujo que no exista.

Cambios esperados:

1. `sap2000_gui.py` debe aceptar `--config` y `--allow-unsafe-output`, y tolerar `--gui` para no romper wrappers.
2. `sap2000_portable.spec` debe generar un EXE GUI real sin consola, separado del EXE CLI.
3. `ejecutar_capturas.bat` debe detectar `--gui` y elegir el entrypoint correcto sin duplicar argumentos.
4. `README.md` debe describir con precision:
   - cuando usar CLI,
   - cuando usar la GUI,
   - que hace el launcher por defecto,
   - y como invocar la GUI desde Python, EXE y `.bat`.
