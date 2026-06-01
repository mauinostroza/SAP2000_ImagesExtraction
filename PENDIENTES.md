# Pendientes

## Estado actual

- La captura de imagen por pantalla funciona y ya no depende de `PrintWindow`.
- La conexión COM a SAP2000 23 funciona cargando `SAP2000v1.tlb`.
- `SapModel.View` existe, pero en este runtime solo expone `RefreshView` y `RefreshWindow`.
- `SetView` y `UnzoomAll` no están disponibles en la interfaz COM observada.
- La automatización UI tiene salvaguardas activas:
  - requiere armado explícito desde la GUI
  - aborta con `Esc`
  - aborta si cambia el foco fuera del contexto SAP2000
  - no envía `Escape` de limpieza si aborta antes del primer input

## Hallazgos UI

- SAP2000 no expone un menú Win32 clásico usable con `GetMenu`.
- La ventana principal sí expone muchos controles hijos WinForms.
- Entre los hijos detectados aparecen al menos:
  - `View`
  - `Display`
  - `Define`
  - `Draw`
  - `Design`
  - `MenuStrip1`
  - `3-D View`
- Al hacer un click seguro sobre el hijo `View`, aparecen nuevas ventanas top-level del mismo proceso con clase:
  - `WindowsForms10.Window.20808.app.0.141b42a_r6_ad1`
- Esas ventanas no tienen título y probablemente correspondan a popups o menús flotantes WinForms.

## Conclusión técnica actual

- La ruta correcta ya no es automatizar por:
  - métodos COM de vista, porque no están disponibles
  - menú Win32 tradicional, porque SAP2000 no lo expone en esta instalación
- La ruta más factible es automatizar sobre controles WinForms reales del proceso SAP2000.

## Próximos pasos

1. Detectar qué ventanas `WindowsForms10.Window.20808...` aparecen únicamente después de clickear el hijo `View`.
2. Enumerar los hijos y textos de esas ventanas popup.
3. Buscar en esos popups cadenas o controles equivalentes a:
   - `Set 3D View`
   - `3-D View`
   - `Rotate`
   - `Zoom`
4. Si aparecen, interactuar por `hwnd`/mensaje Win32 dentro del proceso SAP2000, sin teclas globales.
5. Repetir el mismo patrón para `Display`, con foco en:
   - `Show Load Assigns`
   - selección de patrones/casos

## Mejoras funcionales pendientes

- Implementar control visual real de `ISO_3D` con `azimuth` y `elevation`.
- Implementar `PLAN_XY`, `ELEV_XZ` y `ELEV_YZ` por UI segura.
- Implementar `is_extruded` por UI segura.
- Implementar selección visual de `LOAD_PATTERN` y, más adelante, displays tipo `FRAME_FORCES`.

## Criterios de seguridad que deben mantenerse

- No enviar teclas globales si no hay confirmación explícita del usuario.
- No actuar si la ventana foreground no es SAP2000 o un diálogo hijo/propietario del mismo proceso.
- Preferir interacción por control `hwnd` o popup WinForms antes que navegación ciega por teclado.
- Mantener abortos por `Esc` y por cambio de foco.
