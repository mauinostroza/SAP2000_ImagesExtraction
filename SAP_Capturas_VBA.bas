Attribute VB_Name = "SAP_Capturas"
' ==============================================================================
' SAP_Capturas.bas  v4.0
' Captura automatica de imagenes SAP2000 v23 - Solo VBA + PowerShell
'
' MACROS disponibles:
'   CapturarImagenes    -> ejecuta las capturas marcadas SI
'   AbrirCarpetaSalida  -> abre la carpeta de salida en el Explorador
'   DiagnosticoSAP2000  -> verifica conexion COM, HWND y API de vistas
'   LimpiarResultados   -> borra columnas M y N de la hoja CAPTURAS
'   CalibracionMenus    -> ayuda a encontrar la posicion correcta en menus
' ==============================================================================
Option Explicit

' ── Win32 API ─────────────────────────────────────────────────────────────────
#If VBA7 Then
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function BringWindowToTop Lib "user32" _
        (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function SetCursorPos Lib "user32" _
        (ByVal X As Long, ByVal Y As Long) As Long
#Else
    Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" _
        (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function BringWindowToTop Lib "user32" _
        (ByVal hwnd As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
    Private Declare Function SetCursorPos Lib "user32" _
        (ByVal X As Long, ByVal Y As Long) As Long
#End If

Private Const SW_RESTORE  As Long = 9
Private Const SW_MAXIMIZE As Long = 3

' Angulos: "azimut,elevacion"
Private Const ANG_PLANTA As String = "270,89"
Private Const ANG_ELEV_X As String = "270,0"
Private Const ANG_ELEV_Y As String = "0,0"
Private Const ANG_ISO_NE As String = "225,30"
Private Const ANG_ISO_NO As String = "315,30"
Private Const ANG_ISO_SE As String = "135,30"
Private Const ANG_ISO_SO As String = "45,30"

' ── Posiciones de menus SendKeys (ajustar si no funcionan en tu version) ──────
' Numero de veces que se pulsa DOWN en el menu View hasta "Set 3D View"
Private Const MENU_VIEW_SET3D_FLECHAS As Integer = 2

' Numero de veces en menu Display hasta "Show Load Assigns"
Private Const MENU_DISP_LOADASSIGN_FLECHAS As Integer = 3

' Numero de flechas en subMenu de Load Assigns hasta el tipo deseado
' 1 = Frame Loads,  2 = Shell Loads,  3 = Joint Loads
Private Const MENU_LOADASSIGN_TIPO_FLECHAS As Integer = 1


' ==============================================================================
'  MACRO PRINCIPAL
' ==============================================================================
Sub CapturarImagenes()

    ' ── Hojas ─────────────────────────────────────────────────────────────────
    Dim shCfg As Worksheet
    Dim shCap As Worksheet
    On Error Resume Next
    Set shCfg = ThisWorkbook.Sheets("CONFIG")
    Set shCap = ThisWorkbook.Sheets("CAPTURAS")
    On Error GoTo 0
    If shCfg Is Nothing Or shCap Is Nothing Then
        MsgBox "No se encontraron las hojas CONFIG y/o CAPTURAS.", vbCritical
        Exit Sub
    End If

    ' ── Configuracion ─────────────────────────────────────────────────────────
    Dim proyecto   As String
    Dim carpetaRel As String
    proyecto   = NzStr(shCfg.Range("B3").Value, "Modelo")
    carpetaRel = NzStr(shCfg.Range("B4").Value, "Capturas_SAP")

    Dim carpeta As String
    carpeta = ResolverCarpetaLocal(carpetaRel)
    CrearCarpetaFSO carpeta
    If Not CarpetaExiste(carpeta) Then
        MsgBox "No se pudo crear la carpeta:" & vbNewLine & carpeta & vbNewLine & vbNewLine & _
               "Escribe una ruta absoluta en B4 de CONFIG" & vbNewLine & _
               "Ejemplo:  C:\Proyectos\Capturas_SAP", vbCritical
        Exit Sub
    End If

    Dim psDir As String
    psDir = Environ("TEMP") & "\SAP_PS_TEMP"
    CrearCarpetaFSO psDir

    ' ── SAP2000 ───────────────────────────────────────────────────────────────
    Dim SapModel As Object
    Set SapModel = ConectarSAP2000()
    If SapModel Is Nothing Then
        MsgBox "No se pudo conectar a SAP2000." & vbNewLine & _
               "Ejecuta DiagnosticoSAP2000 para mas informacion.", vbCritical
        Exit Sub
    End If

    ' ── Contar activas ────────────────────────────────────────────────────────
    Dim uf As Long
    uf = shCap.Cells(shCap.Rows.Count, "B").End(xlUp).Row
    If uf < 3 Then MsgBox "No hay datos en CAPTURAS.", vbInformation: Exit Sub

    Dim totalActivas As Long
    Dim f As Long
    For f = 3 To uf
        If EsActivo(shCap.Cells(f, 1).Value) Then totalActivas = totalActivas + 1
    Next f
    If totalActivas = 0 Then
        MsgBox "Ninguna fila marcada como SI en columna A.", vbInformation
        Exit Sub
    End If

    If MsgBox("Se procesaran " & totalActivas & " capturas." & vbNewLine & vbNewLine & _
              "Carpeta de salida:" & vbNewLine & carpeta & vbNewLine & vbNewLine & _
              "IMPORTANTE: No mover el mouse ni teclear" & vbNewLine & _
              "mientras se ejecuta el proceso." & vbNewLine & vbNewLine & _
              "Deseas continuar?", vbQuestion + vbYesNo, "SAP2000 Image Capture") = vbNo Then
        Exit Sub
    End If

    ' ── Obtener y preparar ventana SAP2000 ────────────────────────────────────
    Dim hwndSAP As Long
    hwndSAP = ObtenerHwndSAP2000()
    If hwndSAP = 0 Then
        MsgBox "Ventana de SAP2000 no encontrada. " & _
               "Ejecuta DiagnosticoSAP2000.", vbCritical
        Exit Sub
    End If

    ' Maximizar SAP2000 para captura consistente
    ShowWindow hwndSAP, SW_MAXIMIZE
    Sleep 800
    ActivarVentana hwndSAP

    Call LimpiarResultados

    ' Log de inicio
    Dim logPath As String
    logPath = carpeta & "\capture_log.txt"
    EscribirLog logPath, "INICIO " & Now & " | carpeta=" & carpeta, True
    EscribirLog logPath, "hwndSAP=" & hwndSAP, False

    ' ── Variables del bucle ───────────────────────────────────────────────────
    Dim okCount     As Long
    Dim errCount    As Long
    Dim nombreImg   As String
    Dim tipoVista   As String
    Dim az          As Double
    Dim el          As Double
    Dim modoDisplay As String
    Dim patronCarga As String
    Dim tipoVentana As String
    Dim cropIzq     As Double
    Dim cropSup     As Double
    Dim cropDer     As Double
    Dim cropInf     As Double
    Dim ang         As String
    Dim sufijo      As String
    Dim nombreArch  As String
    Dim rutaArch    As String
    Dim okCap       As Boolean
    Dim msgEstado   As String

    ' ── Bucle principal ───────────────────────────────────────────────────────
    For f = 3 To uf
        If Not EsActivo(shCap.Cells(f, 1).Value) Then GoTo SiguienteFila

        ' Leer fila
        nombreImg   = NzStr(shCap.Cells(f, 2).Value,  "img_" & f)
        tipoVista   = UCase(NzStr(shCap.Cells(f, 3).Value,  "ISO_NE"))
        az          = NzDbl(shCap.Cells(f, 4).Value,  225)
        el          = NzDbl(shCap.Cells(f, 5).Value,  30)
        modoDisplay = UCase(NzStr(shCap.Cells(f, 6).Value,  "MODELO"))
        patronCarga = NzStr(shCap.Cells(f, 7).Value,  "")
        tipoVentana = UCase(NzStr(shCap.Cells(f, 8).Value,  "COMPLETA"))
        cropIzq     = NzDbl(shCap.Cells(f, 9).Value,  0)
        cropSup     = NzDbl(shCap.Cells(f, 10).Value, 0)
        cropDer     = NzDbl(shCap.Cells(f, 11).Value, 100)
        cropInf     = NzDbl(shCap.Cells(f, 12).Value, 100)
        If cropDer <= 0 Then cropDer = 100
        If cropInf <= 0 Then cropInf = 100

        If tipoVista <> "CUSTOM" Then
            ang = AngulosDeVista(tipoVista)
            az  = CDbl(Split(ang, ",")(0))
            el  = CDbl(Split(ang, ",")(1))
        End If

        EscribirLog logPath, "--- Fila " & f & ": " & nombreImg & _
                             " vista=" & tipoVista & " az=" & az & " el=" & el & _
                             " display=" & modoDisplay & " patron=" & patronCarga, False

        msgEstado = ""

        ' 1. Cambiar vista 3D (via API)
        Dim retVista As Long
        retVista = CambiarVista3D(SapModel, az, el)
        If retVista <> 0 Then
            msgEstado = "Vista API fallo(ret=" & retVista & ");"
            EscribirLog logPath, "  WARNING: Set3DView ret=" & retVista, False
        End If

        ' 2. Zoom ajustado al modelo
        On Error Resume Next
        SapModel.View.RefreshView 0, True
        On Error GoTo 0
        Sleep 700

        ' 3. Display (modelo o cargas)
        If modoDisplay = "CARGAS" And Trim(patronCarga) <> "" Then
            Dim okDisplay As Boolean
            okDisplay = MostrarCargas(hwndSAP, SapModel, patronCarga)
            If Not okDisplay Then
                msgEstado = msgEstado & "Display cargas fallo;"
                EscribirLog logPath, "  WARNING: display cargas fallo", False
            End If
        Else
            ' Vista limpia del modelo
            On Error Resume Next
            SapModel.View.RefreshView 0, False
            On Error GoTo 0
            Sleep 400
        End If

        Sleep 500

        ' 4. Nombre del archivo
        If modoDisplay = "CARGAS" And patronCarga <> "" Then
            sufijo = "_CARGAS_" & SanitizarNombre(patronCarga)
        Else
            sufijo = "_MODELO"
        End If
        nombreArch = SanitizarNombre(proyecto) & "_" & _
                     SanitizarNombre(nombreImg) & "_" & tipoVista & sufijo & ".png"
        rutaArch   = carpeta & "\" & nombreArch

        ' 5. Capturar (asegurando que SAP2000 este al frente)
        ActivarVentana hwndSAP
        Sleep 400

        If tipoVentana = "PARCIAL" Then
            okCap = CapturarConPS(hwndSAP, rutaArch, psDir, cropIzq, cropSup, cropDer, cropInf)
        Else
            okCap = CapturarConPS(hwndSAP, rutaArch, psDir, 0, 0, 100, 100)
        End If

        ' 6. Resultado
        If okCap Then
            shCap.Cells(f, 13).Value          = IIf(msgEstado = "", "OK", "OK*")
            shCap.Cells(f, 13).Interior.Color = IIf(msgEstado = "", RGB(198, 239, 206), RGB(255, 235, 156))
            shCap.Cells(f, 14).Value          = nombreArch & IIf(msgEstado <> "", " [" & msgEstado & "]", "")
            okCount = okCount + 1
            EscribirLog logPath, "  -> OK: " & nombreArch, False
        Else
            shCap.Cells(f, 13).Value          = "ERROR"
            shCap.Cells(f, 13).Interior.Color = RGB(255, 199, 206)
            shCap.Cells(f, 14).Value          = "Captura fallida"
            errCount = errCount + 1
            EscribirLog logPath, "  -> ERROR: captura fallida para " & nombreArch, False
        End If
        shCap.Cells(f, 13).Font.Bold           = True
        shCap.Cells(f, 13).HorizontalAlignment = xlCenter

SiguienteFila:
        DoEvents
    Next f

    EscribirLog logPath, "FIN | OK=" & okCount & " ERR=" & errCount, False

    MsgBox "Capturas completadas." & vbNewLine & vbNewLine & _
           "OK     : " & okCount & vbNewLine & _
           "Errores: " & errCount & vbNewLine & vbNewLine & _
           "Carpeta: " & carpeta & vbNewLine & _
           "(* = imagen guardada con advertencias, ver log)", _
           vbInformation, "SAP2000 Image Capture"
End Sub


' ==============================================================================
'  CONECTAR A SAP2000  (3 estrategias)
' ==============================================================================
Private Function ConectarSAP2000() As Object

    Dim helper As Object
    Dim sapObj As Object
    Dim sapMod As Object

    ' 1: Helper COM (metodo oficial CSI)
    On Error Resume Next
    Set helper = CreateObject("SAP2000v1.Helper")
    On Error GoTo 0
    If Not helper Is Nothing Then
        On Error Resume Next
        Set sapObj = helper.GetObject("CSI.SAP2000.API.SapObject")
        On Error GoTo 0
        If Not sapObj Is Nothing Then
            On Error Resume Next
            Set sapMod = sapObj.SapModel
            On Error GoTo 0
            If Not sapMod Is Nothing Then
                Set ConectarSAP2000 = sapMod: Exit Function
            End If
        End If
    End If

    ' 2: GetObject ProgID generico
    On Error Resume Next
    Set sapObj = GetObject(, "CSI.SAP2000.API.SapObject")
    On Error GoTo 0
    If Not sapObj Is Nothing Then
        On Error Resume Next
        Set sapMod = sapObj.SapModel
        On Error GoTo 0
        If Not sapMod Is Nothing Then
            Set ConectarSAP2000 = sapMod: Exit Function
        End If
    End If

    ' 3: ProgIDs con version
    Dim pids(3) As String
    pids(0) = "CSI.SAP2000v23.API.SapObject"
    pids(1) = "CSI.SAP2000v24.API.SapObject"
    pids(2) = "CSI.SAP2000v22.API.SapObject"
    pids(3) = "CSI.SAP2000v21.API.SapObject"
    Dim i As Integer
    For i = 0 To 3
        On Error Resume Next
        Set sapObj = GetObject(, pids(i))
        On Error GoTo 0
        If Not sapObj Is Nothing Then
            On Error Resume Next
            Set sapMod = sapObj.SapModel
            On Error GoTo 0
            If Not sapMod Is Nothing Then
                Set ConectarSAP2000 = sapMod: Exit Function
            End If
        End If
    Next i

    Set ConectarSAP2000 = Nothing
End Function


' ==============================================================================
'  RESOLVER CARPETA LOCAL
'  CORRECCION: solo verifica http/https. Las rutas C:\...\OneDrive\... son locales.
' ==============================================================================
Private Function ResolverCarpetaLocal(carpetaRel As String) As String

    ' Si ya es ruta absoluta (C:\... o \\...)
    If Len(carpetaRel) >= 2 Then
        If Mid(carpetaRel, 2, 1) = ":" Or Left(carpetaRel, 2) = "\\" Then
            ResolverCarpetaLocal = carpetaRel: Exit Function
        End If
    End If

    Dim wbPath As String
    wbPath = ThisWorkbook.Path

    ' Es URL si NO comienza con letra de unidad y NO comienza con \\
    ' Las rutas OneDrive sincronizadas (C:\Users\...\OneDrive\...) son locales
    Dim esRutaLocal As Boolean
    esRutaLocal = False
    If Len(wbPath) >= 2 Then
        If Mid(wbPath, 2, 1) = ":" Then esRutaLocal = True   ' C:\...
        If Left(wbPath, 2) = "\\" Then esRutaLocal = True    ' \\server\...
    End If

    If esRutaLocal Then
        ResolverCarpetaLocal = wbPath & "\" & carpetaRel
    Else
        ' URL de OneDrive web / SharePoint -> usar Documents local
        ResolverCarpetaLocal = Environ("USERPROFILE") & "\Documents\" & carpetaRel
    End If
End Function


' ==============================================================================
'  FSO helpers
' ==============================================================================
Private Sub CrearCarpetaFSO(ruta As String)
    If ruta = "" Then Exit Sub
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(ruta) Then fso.CreateFolder ruta
    On Error GoTo 0
End Sub

Private Function CarpetaExiste(ruta As String) As Boolean
    If ruta = "" Then CarpetaExiste = False: Exit Function
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    CarpetaExiste = fso.FolderExists(ruta)
    If Err.Number <> 0 Then CarpetaExiste = False
    On Error GoTo 0
End Function


' ==============================================================================
'  OBTENER HWND DE SAP2000
'  CORRECCION: filtra por ventana visible con titulo no vacio (ventana principal)
' ==============================================================================
Private Function ObtenerHwndSAP2000() As Long

    Dim psPath   As String
    Dim resPath  As String
    psPath  = Environ("TEMP") & "\get_sap_hwnd.ps1"
    resPath = Environ("TEMP") & "\sap_hwnd_result.txt"

    On Error Resume Next
    If Dir(resPath) <> "" Then Kill resPath
    If Dir(psPath)  <> "" Then Kill psPath
    On Error GoTo 0

    ' Script PS: busca proceso SAP2000 con ventana principal visible (titulo no vacio)
    Dim fnum As Integer
    fnum = FreeFile
    Open psPath For Output As #fnum
    Print #fnum, "$hw = 0"
    Print #fnum, "$procs = Get-Process | Where-Object { $_.Name -like '*SAP2000*' }"
    Print #fnum, "foreach ($p in $procs) {"
    Print #fnum, "    if ($p.MainWindowHandle.ToInt64() -gt 0 -and $p.MainWindowTitle -ne '') {"
    Print #fnum, "        $hw = $p.MainWindowHandle.ToInt64(); break"
    Print #fnum, "    }"
    Print #fnum, "}"
    Print #fnum, "$hw | Out-File '" & resPath & "' -Encoding ascii"
    Close #fnum

    Shell "powershell.exe -NonInteractive -ExecutionPolicy Bypass " & _
          "-WindowStyle Hidden -File """ & psPath & """", vbHide

    EsperarArchivo resPath, 8000

    Dim linea As String
    linea = ""
    If Dir(resPath) <> "" Then
        fnum = FreeFile
        Open resPath For Input As #fnum
        If Not EOF(fnum) Then Line Input #fnum, linea
        Close #fnum
        Kill resPath
    End If
    If Dir(psPath) <> "" Then Kill psPath

    On Error Resume Next
    ObtenerHwndSAP2000 = CLng(Trim(linea))
    If Err.Number <> 0 Then ObtenerHwndSAP2000 = 0
    On Error GoTo 0
End Function


' ==============================================================================
'  ACTIVAR VENTANA SAP2000  (mas robusto que solo AppActivate)
'  CORRECCION: usa SetForegroundWindow + BringWindowToTop + verifica el foco
' ==============================================================================
Private Sub ActivarVentana(hwnd As Long)
    ShowWindow hwnd, SW_RESTORE
    BringWindowToTop hwnd
    SetForegroundWindow hwnd
    Sleep 600
    ' Verificar que realmente tiene el foco
    Dim intentos As Integer
    intentos = 0
    Do While GetForegroundWindow <> hwnd And intentos < 5
        SetForegroundWindow hwnd
        Sleep 300
        intentos = intentos + 1
    Loop
End Sub


' ==============================================================================
'  ANGULOS POR TIPO DE VISTA
' ==============================================================================
Private Function AngulosDeVista(t As String) As String
    Select Case UCase(Trim(t))
        Case "PLANTA":  AngulosDeVista = ANG_PLANTA
        Case "ELEV_X":  AngulosDeVista = ANG_ELEV_X
        Case "ELEV_Y":  AngulosDeVista = ANG_ELEV_Y
        Case "ISO_NE":  AngulosDeVista = ANG_ISO_NE
        Case "ISO_NO":  AngulosDeVista = ANG_ISO_NO
        Case "ISO_SE":  AngulosDeVista = ANG_ISO_SE
        Case "ISO_SO":  AngulosDeVista = ANG_ISO_SO
        Case Else:      AngulosDeVista = ANG_ISO_NE
    End Select
End Function


' ==============================================================================
'  CAMBIAR VISTA 3D  -  SOLO via API
'  Devuelve 0 si tuvo exito, otro valor si fallo.
'  No usa SendKeys para vistas (causa seleccion de nodos y menus erroneos).
' ==============================================================================
Private Function CambiarVista3D(SapModel As Object, az As Double, el As Double) As Long

    Dim ret As Long

    ' Intento 1: Set3DView(HorizAngle, VertAngle)
    On Error Resume Next
    ret = SapModel.View.Set3DView(az, el)
    On Error GoTo 0
    If ret = 0 Then CambiarVista3D = 0: Exit Function

    ' Intento 2: parametros invertidos (algunos builds los tienen al reves)
    On Error Resume Next
    ret = SapModel.View.Set3DView(el, az)
    On Error GoTo 0
    If ret = 0 Then CambiarVista3D = 0: Exit Function

    ' Si API falla: usar SendKeys SOLO para cambio de vista
    ' (mas controlado y aislado que para menus de display)
    Dim retSK As Boolean
    retSK = CambiarVistaConSendKeys(az, el)
    CambiarVista3D = IIf(retSK, 0, -1)
End Function


' ==============================================================================
'  CAMBIAR VISTA VIA MENU (SendKeys) - solo como fallback de Set3DView
'  CORRECCION: activa ventana con Win32, mueve mouse a zona segura,
'  escapa dialogos abiertos, y usa la posicion correcta del menu View.
' ==============================================================================
Private Function CambiarVistaConSendKeys(az As Double, el As Double) As Boolean

    On Error GoTo ErrorSendKeys

    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")

    ' Activar SAP2000 via Win32 (mas fiable que AppActivate solo)
    Dim hw As Long
    hw = ObtenerHwndSAP2000()
    If hw = 0 Then CambiarVistaConSendKeys = False: Exit Function
    ActivarVentana hw

    ' Mover mouse a esquina superior izquierda de la ventana (zona segura)
    ' para evitar activar herramientas del viewport al recibir teclas Alt
    SetCursorPos 10, 10
    Sleep 200

    ' Cerrar cualquier dialogo o menu abierto antes de navegar
    ws.SendKeys "{ESCAPE}", True:  Sleep 150
    ws.SendKeys "{ESCAPE}", True:  Sleep 150

    ' Abrir menu View con Alt+V
    ws.SendKeys "%V", True
    Sleep 400

    ' Bajar hasta "Set 3D View..." en el menu View
    ' MENU_VIEW_SET3D_FLECHAS se puede ajustar en las constantes al inicio
    Dim i As Integer
    For i = 1 To MENU_VIEW_SET3D_FLECHAS
        ws.SendKeys "{DOWN}", True
        Sleep 100
    Next i
    ws.SendKeys "{ENTER}", True
    Sleep 700

    ' Dialogo "Set 3D View": campo "Plan Rotation" (azimut)
    ws.SendKeys "^a", True:  Sleep 100
    ws.SendKeys CStr(CLng(az)), True
    ws.SendKeys "{TAB}", True:  Sleep 200

    ' Campo "Elevation Angle"
    ws.SendKeys "^a", True:  Sleep 100
    ws.SendKeys CStr(CLng(el)), True
    ws.SendKeys "{ENTER}", True
    Sleep 500

    CambiarVistaConSendKeys = True
    Exit Function

ErrorSendKeys:
    CambiarVistaConSendKeys = False
End Function


' ==============================================================================
'  MOSTRAR CARGAS VIA MENU (Display > Show Load Assigns)
'  CORRECCION: activa ventana, escapa primero, evita interferencias.
'  Devuelve True si navego el menu sin errores.
' ==============================================================================
Private Function MostrarCargas(hwnd As Long, SapModel As Object, patron As String) As Boolean

    On Error GoTo ErrorCargas

    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")

    ActivarVentana hwnd
    SetCursorPos 10, 10   ' mouse a zona segura
    Sleep 200

    ' Cerrar estado previo
    ws.SendKeys "{ESCAPE}", True:  Sleep 150
    ws.SendKeys "{ESCAPE}", True:  Sleep 150

    ' Abrir menu Display con Alt+D
    ws.SendKeys "%D", True
    Sleep 400

    ' Bajar hasta "Show Load Assigns"
    For i = 1 To MENU_DISP_LOADASSIGN_FLECHAS
        ws.SendKeys "{DOWN}", True
        Sleep 100
    Next i

    ' Abrir subMenu
    ws.SendKeys "{RIGHT}", True
    Sleep 350

    ' Bajar hasta tipo de carga (Frame, Shell o Joint)
    For i = 1 To MENU_LOADASSIGN_TIPO_FLECHAS
        ws.SendKeys "{DOWN}", True
        Sleep 100
    Next i
    ws.SendKeys "{ENTER}", True
    Sleep 800

    ' En el dialogo: seleccionar patron de carga
    ' Intentar Ctrl+A para limpiar y escribir el nombre
    ws.SendKeys "^a", True:  Sleep 100
    ws.SendKeys patron, True: Sleep 150
    ws.SendKeys "{ENTER}", True
    Sleep 700

    MostrarCargas = True
    Exit Function

ErrorCargas:
    ' Si falla: refrescar vista limpia del modelo
    On Error Resume Next
    SapModel.View.RefreshView 0, False
    On Error GoTo 0
    MostrarCargas = False
End Function


' ==============================================================================
'  CAPTURAR VENTANA SAP2000 CON POWERSHELL
'  CORRECCION:
'   1. Llama SetProcessDPIAware() para que GetWindowRect devuelva coordenadas
'      fisicas (soluciona el problema de "ventana cortada" en monitores HiDPI)
'   2. Captura pantalla completa y recorta a los limites de la ventana SAP2000
'   3. Aplica crop adicional opcional
' ==============================================================================
Private Function CapturarConPS(hwnd As Long, filePath As String, psDir As String, _
    cropIzq As Double, cropSup As Double, cropDer As Double, cropInf As Double) As Boolean

    ActivarVentana hwnd
    Sleep 300

    Dim ts      As String
    ts = Format(Now, "HHmmss") & Right(CStr(CLng(Timer * 100)), 4)
    Dim psPath  As String
    Dim resPath As String
    psPath  = psDir & "\cap_" & ts & ".ps1"
    resPath = psDir & "\res_" & ts & ".txt"

    ' Numeros con punto decimal (independiente del locale de Windows)
    Dim sL As String, sT As String, sR As String, sB As String
    sL = Replace(Format(cropIzq, "0.00"), ",", ".")
    sT = Replace(Format(cropSup, "0.00"), ",", ".")
    sR = Replace(Format(cropDer, "0.00"), ",", ".")
    sB = Replace(Format(cropInf, "0.00"), ",", ".")

    Dim fpPS As String
    Dim rpPS As String
    fpPS = Replace(filePath, "'", "''")
    rpPS = Replace(resPath,  "'", "''")

    Dim L  As String
    Dim ps As String
    L = vbLf

    ' Bloque 1: DPI awareness (CORRECCION principal para ventana cortada)
    ps = "Add-Type @'" & L
    ps = ps & "using System;" & L
    ps = ps & "using System.Runtime.InteropServices;" & L
    ps = ps & "public class DpiWin {" & L
    ps = ps & "    [DllImport(" & Chr(34) & "user32.dll" & Chr(34) & ")]" & L
    ps = ps & "    public static extern bool SetProcessDPIAware();" & L
    ps = ps & "    [DllImport(" & Chr(34) & "user32.dll" & Chr(34) & ")]" & L
    ps = ps & "    public static extern bool GetWindowRect(IntPtr h, ref RECT r);" & L
    ps = ps & "    public struct RECT { public int L, T, R, B; }" & L
    ps = ps & "    public static int[] Bounds(long hwnd) {" & L
    ps = ps & "        SetProcessDPIAware();" & L
    ps = ps & "        RECT r = new RECT();" & L
    ps = ps & "        GetWindowRect(new IntPtr(hwnd), ref r);" & L
    ps = ps & "        return new int[] { r.L, r.T, r.R, r.B };" & L
    ps = ps & "    }" & L
    ps = ps & "}" & L
    ps = ps & "'@" & L

    ' Bloque 2: obtener coordenadas fisicas de la ventana
    ps = ps & "$rc = [DpiWin]::Bounds(" & CStr(hwnd) & ")" & L
    ps = ps & "$wx=$rc[0]; $wy=$rc[1]; $ww=$rc[2]-$rc[0]; $wh=$rc[3]-$rc[1]" & L
    ps = ps & "if ($ww -lt 10 -or $wh -lt 10) {" & L
    ps = ps & "    'ERR:WindowTooSmall' | Out-File '" & rpPS & "' -Enc ascii; exit" & L
    ps = ps & "}" & L

    ' Bloque 3: capturar con CopyFromScreen (coordenadas fisicas)
    ps = ps & "Add-Type -AssemblyName System.Drawing" & L
    ps = ps & "$bmp = New-Object System.Drawing.Bitmap($ww, $wh)" & L
    ps = ps & "$g = [System.Drawing.Graphics]::FromImage($bmp)" & L
    ps = ps & "$g.CopyFromScreen($wx, $wy, 0, 0, (New-Object System.Drawing.Size($ww,$wh)))" & L
    ps = ps & "$g.Dispose()" & L

    ' Bloque 4: crop adicional
    ps = ps & "$x1=[int]($ww*" & sL & "/100); $y1=[int]($wh*" & sT & "/100)" & L
    ps = ps & "$x2=[int]($ww*" & sR & "/100); $y2=[int]($wh*" & sB & "/100)" & L
    ps = ps & "$cw=$x2-$x1; $ch=$y2-$y1" & L
    ps = ps & "if ($cw -gt 0 -and $ch -gt 0) {" & L
    ps = ps & "    $bmp=$bmp.Clone((New-Object System.Drawing.Rectangle($x1,$y1,$cw,$ch)),$bmp.PixelFormat)" & L
    ps = ps & "}" & L

    ' Bloque 5: guardar PNG
    ps = ps & "try {" & L
    ps = ps & "    $bmp.Save('" & fpPS & "',[System.Drawing.Imaging.ImageFormat]::Png)" & L
    ps = ps & "    'OK' | Out-File '" & rpPS & "' -Enc ascii" & L
    ps = ps & "} catch {" & L
    ps = ps & "    ('ERR:'+$_.Exception.Message)|Out-File '" & rpPS & "' -Enc ascii" & L
    ps = ps & "}" & L

    ' Guardar y ejecutar
    Dim fnum As Integer
    fnum = FreeFile
    Open psPath For Output As #fnum
    Print #fnum, ps
    Close #fnum

    Shell "powershell.exe -NonInteractive -ExecutionPolicy Bypass " & _
          "-WindowStyle Hidden -File """ & psPath & """", vbHide

    EsperarArchivo resPath, 15000

    ' Leer resultado para log
    Dim resultado As String
    resultado = ""
    If Dir(resPath) <> "" Then
        fnum = FreeFile
        Open resPath For Input As #fnum
        Dim ln As String
        Do While Not EOF(fnum): Line Input #fnum, ln: resultado = resultado & Trim(ln): Loop
        Close #fnum
        Kill resPath
    End If
    On Error Resume Next
    If Dir(psPath) <> "" Then Kill psPath
    On Error GoTo 0

    CapturarConPS = (Len(filePath) > 0) And (Len(Dir(filePath)) > 0)
End Function


' ==============================================================================
'  ESPERAR ARCHIVO con timeout
' ==============================================================================
Private Sub EsperarArchivo(ruta As String, maxMs As Long)
    Dim t0 As Single
    t0 = Timer
    Do While Dir(ruta) = ""
        Sleep 400
        DoEvents
        If (Timer - t0) * 1000 >= maxMs Then Exit Do
    Loop
End Sub


' ==============================================================================
'  LOG
' ==============================================================================
Private Sub EscribirLog(logPath As String, msg As String, nuevo As Boolean)
    On Error Resume Next
    Dim fnum As Integer
    fnum = FreeFile
    If nuevo Then
        Open logPath For Output As #fnum
    Else
        Open logPath For Append As #fnum
    End If
    Print #fnum, Format(Now, "hh:mm:ss") & "  " & msg
    Close #fnum
    On Error GoTo 0
End Sub


' ==============================================================================
'  HELPERS
' ==============================================================================
Private Function EsActivo(val As Variant) As Boolean
    If IsNull(val) Or IsEmpty(val) Then EsActivo = False: Exit Function
    Dim s As String
    s = UCase(Trim(CStr(val)))
    EsActivo = (s = "SI" Or s = "YES" Or s = "1" Or s = "X")
End Function

Private Function NzStr(val As Variant, def As String) As String
    If IsNull(val) Or IsEmpty(val) Or Trim(CStr(val)) = "" Then
        NzStr = def
    Else
        NzStr = Trim(CStr(val))
    End If
End Function

Private Function NzDbl(val As Variant, def As Double) As Double
    If IsNull(val) Or IsEmpty(val) Or Trim(CStr(val)) = "" Then
        NzDbl = def: Exit Function
    End If
    On Error Resume Next
    NzDbl = CDbl(val)
    If Err.Number <> 0 Then NzDbl = def
    On Error GoTo 0
End Function

Private Function SanitizarNombre(n As String) As String
    Dim res As String
    Dim j   As Integer
    Dim c   As String
    res = n
    Dim chars As String
    chars = "\/:*?""<>| "
    For j = 1 To Len(chars)
        c   = Mid(chars, j, 1)
        res = Replace(res, c, "_")
    Next j
    SanitizarNombre = res
End Function


' ==============================================================================
'  LIMPIAR RESULTADOS
' ==============================================================================
Sub LimpiarResultados()
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Sheets("CAPTURAS")
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub
    Dim uf As Long
    uf = sh.Cells(sh.Rows.Count, "B").End(xlUp).Row
    If uf < 3 Then Exit Sub
    sh.Range("M3:N" & uf).ClearContents
    sh.Range("M3:N" & uf).Interior.ColorIndex = xlNone
    With sh.Range("M3:N" & uf).Borders
        .LineStyle = xlContinuous
        .Color     = RGB(136, 136, 136)
        .Weight    = xlThin
    End With
End Sub


' ==============================================================================
'  ABRIR CARPETA DE SALIDA
' ==============================================================================
Sub AbrirCarpetaSalida()
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Sheets("CONFIG")
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub
    Dim carpeta As String
    carpeta = ResolverCarpetaLocal(NzStr(sh.Range("B4").Value, "Capturas_SAP"))
    If CarpetaExiste(carpeta) Then
        Shell "explorer.exe """ & carpeta & """", vbNormalFocus
    Else
        MsgBox "Carpeta no existe todavia:" & vbNewLine & carpeta, vbInformation
    End If
End Sub


' ==============================================================================
'  DIAGNOSTICO SAP2000
' ==============================================================================
Sub DiagnosticoSAP2000()

    Dim msg As String
    msg = "DIAGNOSTICO SAP2000" & vbNewLine & String(55, "-") & vbNewLine & vbNewLine

    ' Ruta y carpeta resuelta
    msg = msg & "Excel:   " & ThisWorkbook.Path & vbNewLine
    Dim carpetaRel As String
    On Error Resume Next
    carpetaRel = NzStr(ThisWorkbook.Sheets("CONFIG").Range("B4").Value, "Capturas_SAP")
    On Error GoTo 0
    Dim carpetaRes As String
    carpetaRes = ResolverCarpetaLocal(carpetaRel)
    msg = msg & "Carpeta: " & carpetaRes & vbNewLine & vbNewLine

    ' Proceso
    Dim found As Boolean
    Dim oWMI  As Object
    Dim procs As Object
    Dim proc  As Object
    On Error Resume Next
    Set oWMI  = GetObject("winmgmts:")
    Set procs = oWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name LIKE '%SAP2000%'")
    On Error GoTo 0
    If Not procs Is Nothing Then
        For Each proc In procs
            found = True
            msg = msg & "[OK] Proceso: " & proc.Name & " PID=" & proc.ProcessId & vbNewLine
        Next
    End If
    If Not found Then
        msg = msg & "[ERR] SAP2000 no encontrado. Abrelo con un modelo." & vbNewLine
        MsgBox msg, vbCritical, "Diagnostico": Exit Sub
    End If
    msg = msg & vbNewLine

    ' Conexion COM
    Dim sapM As Object
    Set sapM = ConectarSAP2000()
    If sapM Is Nothing Then
        msg = msg & "[ERR] Conexion COM fallo." & vbNewLine
        msg = msg & "  -> Cierra SAP2000 y vuelve a abrirlo con un modelo." & vbNewLine
    Else
        msg = msg & "[OK] Conexion COM exitosa." & vbNewLine
        Dim numPts As Long
        Dim ptN As Variant
        On Error Resume Next
        sapM.PointObj.GetNameList numPts, ptN
        On Error GoTo 0
        msg = msg & "[OK] Nodos en el modelo: " & numPts & vbNewLine

        ' Probar Set3DView con distintos parametros
        Dim r1 As Long, r2 As Long
        On Error Resume Next
        r1 = sapM.View.Set3DView(225, 30)
        On Error GoTo 0
        On Error Resume Next
        r2 = sapM.View.Set3DView(30, 225)
        On Error GoTo 0
        msg = msg & "[INFO] Set3DView(225,30) = " & r1 & _
              IIf(r1 = 0, " (OK - vistas via API)", " (fallo)") & vbNewLine
        msg = msg & "[INFO] Set3DView(30,225) = " & r2 & _
              IIf(r2 = 0, " (OK - params invertidos)", " (fallo)") & vbNewLine

        If r1 <> 0 And r2 <> 0 Then
            msg = msg & vbNewLine & "AVISO: Set3DView no funciona via API." & vbNewLine
            msg = msg & "El cambio de vista usara menus (SendKeys)." & vbNewLine
            msg = msg & "Si los angulos no cambian, ejecuta CalibracionMenus." & vbNewLine
        End If
    End If

    ' HWND
    msg = msg & vbNewLine
    Dim hw As Long
    hw = ObtenerHwndSAP2000()
    If hw <> 0 Then
        msg = msg & "[OK] Ventana SAP2000 HWND=" & hw & vbNewLine
    Else
        msg = msg & "[ERR] Ventana no encontrada. SAP2000 debe ser visible." & vbNewLine
    End If

    msg = msg & vbNewLine & String(55, "-")
    MsgBox msg, vbInformation, "Diagnostico SAP2000"
End Sub


' ==============================================================================
'  CALIBRACION DE MENUS
'  Abre el menu View de SAP2000 con 1 flecha hacia abajo, muestra
'  el resultado, y pregunta si hay que subir o bajar.
'  Ejecutar esta macro para encontrar la posicion correcta de
'  "Set 3D View" y actualizar MENU_VIEW_SET3D_FLECHAS.
' ==============================================================================
Sub CalibracionMenus()

    Dim hw As Long
    hw = ObtenerHwndSAP2000()
    If hw = 0 Then
        MsgBox "Abre SAP2000 primero.", vbCritical: Exit Sub
    End If

    Dim msg As String
    msg = "CALIBRACION DE POSICION DEL MENU VIEW" & vbNewLine & vbNewLine & _
          "Esta macro abrira el menu View de SAP2000 y" & vbNewLine & _
          "bajara N posiciones hasta 'Set 3D View'." & vbNewLine & vbNewLine & _
          "Actualmente N = " & MENU_VIEW_SET3D_FLECHAS & vbNewLine & vbNewLine & _
          "Asegurate de que SAP2000 este abierto y visible." & vbNewLine & vbNewLine & _
          "Presiona OK para abrir el menu y contar la posicion."

    If MsgBox(msg, vbOKCancel, "Calibracion") = vbCancel Then Exit Sub

    ActivarVentana hw
    SetCursorPos 10, 10
    Sleep 300

    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    ws.SendKeys "{ESCAPE}", True: Sleep 150
    ws.SendKeys "{ESCAPE}", True: Sleep 150

    ' Abrir menu View
    ws.SendKeys "%V", True
    Sleep 400

    ' Bajar N posiciones
    Dim i As Integer
    For i = 1 To MENU_VIEW_SET3D_FLECHAS
        ws.SendKeys "{DOWN}", True
        Sleep 150
    Next i

    MsgBox "El menu esta abierto y el cursor esta en la posicion " & _
           MENU_VIEW_SET3D_FLECHAS & "." & vbNewLine & vbNewLine & _
           "Mira SAP2000: el item resaltado deberia ser 'Set 3D View...'." & vbNewLine & vbNewLine & _
           "Si NO es ese item, presiona ESCAPE en SAP2000, " & vbNewLine & _
           "ajusta el valor de MENU_VIEW_SET3D_FLECHAS en el codigo VBA, " & vbNewLine & _
           "y vuelve a ejecutar CalibracionMenus." & vbNewLine & vbNewLine & _
           "Presiona OK para cerrar el menu.", _
           vbInformation, "Calibracion"

    ' Cerrar menu
    ws.SendKeys "{ESCAPE}", True: Sleep 200
    ws.SendKeys "{ESCAPE}", True

End Sub
