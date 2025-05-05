Attribute VB_Name = "modBasics"
'Modulo modBasics
'Origen: Basado en codigos personales y algunas recopilaciones de de codigo internet con modificaciones propias.
'Version: 1.0.1 (18/06/2019)
'
'Descripcion:
'   Modulo con funciones basicas y escenciales para la mayoria de las aplicaciones
'
'Modo de uso:

'Funciones (Publicas):
'   DetectIDE:              Inicializa la variable global IDE (boolean) y devuelve True
'                           En caso de error devolvera False (y probablemente la variable IDE no se inicialice, dependera del error ocurrido)
'                           Opcionalmente en caso de error podra mostrar mensaje (con numero y descripcion de error) y/o cerrar la aplicacion,
'                           Nota: Se utilizan 2 metodos para detectar el IDE, si difieren los resultados entre ellos, generará un error (sin numero ni descripcion de error)
'
'   InitEnv:                Carga Shell32.dll en memoria
'                           Inicializa Microsoft Common Controls
'                           Inicializa las variables basicas del programa:
'                               Variable AppPath (string):  Haciendo uso de LoadAppPath (Ver LoadAppPath)
'                               Variable IDE (boolean):     Haciendo uso de DetectIDE (Ver DetectIDE)
'                           Opcionalmente en caso de error podra mostrar mensaje (con numero y descripcion de error) y/o cerrar la aplicacion
'
'   TerminateEnv:           Debe ser llamado antes de finalizar el programa siempre y cuando se hubiera utilizado InitEnv.
'                           Se encarga de descargar Shell32.dll y otras cosas cargadas por InitEnv.
'                           Opcionalmente en caso de error podra mostrar mensaje (con numero y descripcion de error) y/o cerrar la aplicacion
'
'   PrevInstance:           Devuelve true si existe una instancia previa de la APP o false si no.
'                           *Opcionalmente muestra un MsgBox si se especifica el valor opcional sMsg > Tambien se puede especificar Type (VBMSGBOXSTYLE) y sTitle (string)
'                           *Se ignora la instancia previa si: "comctlrestart = true" y command$ contiene "/comctl"
'
'   FileExist:              Devolvera true si el archivo file (string) existe y es del tipo FileType (VbFileAttribute)
'
'Funciones (privadas):
'   inIDE:                  Devuelve true si se esta ejecutando desde el IDE de VB6 (metodo 1) (usado por DetectIDE)
'   RunningIDE:             Devuelve true si se esta ejecutando desde el IDE de VB6 (metodo 2) (usado por DetectIDE)
'
'
'Subs (Publicas):
'   AsyncSleep:             Detiene un procedimiento X milisegundos (lngMilliseconds).
'                           *Opcional "Once=true" intenta ejecutar la rutina una sola vez (no es infalible)
'   AsyncSleepAllStop:      Detiene cualquier rutina AsyncSleep en ejecucion.
'
'   Pause:                  Similar a AsyncSleep, pero usando API SleepEx, duerme todo el programa, opcionalmente se puede usar DoEvents.
'                           Permite hacer una pausa de X milisegundos (lngMilliseconds), pudiendo optar por detener todo el programa (Async = False)
'                           o simplemente el procedimiento (haciendo uso de DoEvents para seguir capturando eventos (Async = True)
'
'   LogSysInfo:             LogSysInfo creará un archivo de nombre sysinfo.log en el directorio "Path" con la siguiente informacion:
'                           NT:     (Verdadero | Falso)
'                           Server: (Verdadero | Falso)
'                           x64:    (Verdadero | Falso)
'                           Name:   (Ej: Windows 10)
'                           Version:(Ej: 10.0.17763)
'                           Major:  (Ej: 10)
'                           Minor:  (Ej: 0)
'                           Rev:    (Ej: 17763)
'                           *NOTAS:
'                                   -Requiere modulo modWindowsVersion para funcionar
'                                   -Si la constante de compilacion condicional del mismo nombre (LogSysInfo) es 0 esta funcion estara deshabilitada
'                                   -Si es Windows 8 o superior puede informar incorrectamente la version si no se crea el manifiesto de compatibilidad
'                                   -Requiere del modulo modWindowsVersion, de no estar presente en el proyecto no podra compilar. Deje la constante
'                                    LogSysInfo=0 si no cuenta con este modulo
'                                   -El parametro Path es opcional y podra representar cualquier ruta a directorio real, si no se especifica Path se
'                                    utilizara AppPath. Si AppPath no se hubiera inicializado, se llamará a LoadAppPath para inicializar la
'                                    variable (que es equivalente a App.Path & "\").
'                                   -Si Path no terminase con "\" entonces se agregará automaticamente para evitar errores y simplificar la programación.
'                                   -Se producirá un error si Path no es una ruta valida o no se tiene permiso de acceso
'
'   DoManifest:             Creara el archivo manifiesto para la aplicacion, en AppPath & App.EXEName & ".exe.manifest"
'                           (se requiere permisos de escritura en el directorio, caso contrario no creará el archivo. Será ignorado cualquier error)
'                           Parametros:
'                                   -AppName:               Opcional,   por valor,  String,                                             por defecto App.EXEName & ".exe"
'                                   -WindowsReadyManifest:  Opcional,   por valor,  Long (eWindowsReadyManifest) (Ver enumerador),      por defecto "WinAll"
'                                       Indica la compatibilidad a especificar en el manifiesto.
'                                       Si no se especifica, no se insertaran detalles de compatibilidad con el SO
'                                       los valores admisibles (y combinables con mascaras de bits) son
'                                           eWindowsReadyManifest.WinVista  (Windows Vista)
'                                           eWindowsReadyManifest.Win7      (Windows 7)
'                                           eWindowsReadyManifest.Win8      (Windows 8)
'                                           eWindowsReadyManifest.Win81     (Windows 8.1)
'                                           eWindowsReadyManifest.Win10     (Windows 10)
'                                           eWindowsReadyManifest.WinAll    (Todos los anteriores)
'
'                                       Puede combinar varios valores de la siguiente forma:
'                                           'Asignacion:
'                                           Dim WindowsReadyManifest As eWindowsReadyManifest
'                                           WindowsReadyManifest = eWindowsReadyManifest.Win7 Or eWindowsReadyManifest.Win8 Or eWindowsReadyManifest.Win10
'
'                                           'Verificacion:
'                                           If (WindowsReadyManifest And eWindowsReadyManifest.Win7) then
'                                               'Ejecutar Verdadero si esta presente Win7
'                                           Else
'                                               'Ejecutar Flaso si esta presente Win7
'                                           End If
'
'                                   -ExecutionLevel:        Opcional,   por valor,  Long (eExecutionLevel) (Ver enumerador),            por defecto 0
'                                       Los niveles de ejecucion posibles son:
'                                           eExecutionLevel.asInvoker:              Como quien lo ejecuta, no eleva privilegios
'                                           eExecutionLevel.requireAdministrator:   Como Administrador, eleva privilegios
'                                           eExecutionLevel.highestAvailable:       Eleva privilegios si la cuenta lo permite.
'
'                                   -OpcionalCloseAndRun:   Opcional,   por valor,  Boolean,                                            por defecto "False"
'                                       Si es True se ejecutara asimismo pasando como parametro adicional /comctl, respetando la existencia previa
'                                       de parametros y luego se cerrará. Se podrá utilizar /comctl para ignorar restricciones de instancia previa "PrevInstance()"
'                                       Ejemplo: Shell """" & AppPath & App.EXEName & ".exe"" /comctl " & Command$
'                                       NOTA: Esta opcion intenta hacer uso de la variable global IDE (Boolean), que en caso de ser True ignorarà el reinicio.
'                                             IDE debe ser inicializada con DetectIDE o InitEnv (que llama a DetectIDE)
'
'   LoadAppPath:        Inicializa la variable global AppPath, cargando App.Path y agregando una barra invertida al final.
'
Option Explicit

'Habilita (1) o deshabilita (0) la sub LogSysInfo
'**:Requiere modulo modWindowsVersion
#Const LogSysInfo = 0

'Declaracion de APIs Win32
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function SleepEx Lib "kernel32" ( _
    ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
'Private Declare Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
   ByVal hLibModule As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
    
'Variables globales
Public IDE As Boolean                   'Para saber si se esta ejecutando en IDE VB6 o en Ejecutable
Public AppPath As String                'String, ruta al directorio de la app con barra final.

'Variables privadas del modulo
Private m_hMod As Long                  'Para usar en InitEnv. Handle de LoadLibrary
Private bAsyncSleepStop As Boolean      'Para usar en AsyncSleep: premature stop flag
Private bAsyncSleepStopOnce As Boolean  'Para usar en AsyncSleep: Once execution at time
#If LogSysInfo = 1 Then                 'Compilacion condicional
Private WinInfo As OSVERSIONDATA        'Variable WinInfo para LogSysInfo()
#End If

'Enumeradores y tipos definidos por el usuario
Public Enum eWindowsReadyManifest
    'Windows Vista and Windows Server 2008
    WinVista = 1
    'Windows 7 and Windows Server 2008 R2
    Win7 = 2
    'Windows 8 and Windows Server 2012
    Win8 = 4
    'Windows 8.1 and Windows Server 2012 R2
    Win81 = 8
    'Windows 10 and Windows Server 2016
    Win10 = 16
    
    'Todos los anteriores
    WinAll = 4096
End Enum

Public Type tPrevInstance
    sMsg    As String
    Style   As VbMsgBoxStyle
    sTitle  As String
End Type
   
Public Enum eExecutionLevel
    asInvoker = 1
    requireAdministrator = 2
    highestAvailable = 4
End Enum


'Suspende un procedimiento de forma asincrona (con un bucle).
'Permite que el resto de procedimientos y eventos continuen.
Public Sub AsyncSleep(ByVal lngMilliseconds As Long, Optional ByVal Once As Boolean = True)
   Dim lngStartTicks As Long

   'Si el parametro Once es true, verificamos si se esta ejecutanto AsyncSleep, de ser asi salimos del procedimiento
   If Once = True And bAsyncSleepStopOnce = True Then Exit Sub
   
   bAsyncSleepStop = False      'make sure the flag isn't set
   lngStartTicks = GetTickCount 'save the current ticks
   
   bAsyncSleepStopOnce = True   'make the flag true when is running
   
   Do Until bAsyncSleepStop 'start the loop that stops when the flag is True
      'if the requested time has elapsed, leave the loop
      If (GetTickCount - lngStartTicks) >= lngMilliseconds Then Exit Do
      Sleep 1
      DoEvents 'keeps the system from hanging
   Loop
   
   bAsyncSleepStopOnce = False 'make the flag true at end
End Sub
 
Public Sub AsyncSleepAllStop()
   bAsyncSleepStop = True 'test button to make sure it stops on an action
End Sub
 
Public Sub Pause(ByVal lngMilliseconds As Long, Optional ByVal Async As Boolean = True)
   Dim timeout   As Single
   Dim PrevTimer As Single
   Dim SecsDelay As Single
   
   SecsDelay = lngMilliseconds / 1000
   
   PrevTimer = Timer
   timeout = PrevTimer + SecsDelay
   Do While PrevTimer < timeout
      SleepEx 4, False '-- Timer is only updated every 1/64 sec = 15.625 millisecs.
      If Async = True Then DoEvents
      If Timer < PrevTimer Then timeout = timeout - 86400 '-- pass midnight
      PrevTimer = Timer
   Loop
End Sub

#If LogSysInfo = 1 Then
Public Sub LogSysInfo(Optional ByVal path As String)
    Dim SoInfo As OSVERSIONDATA
    SoInfo = SoInfo
    
    If path = "" Then
        If AppPath = "" Then
            LoadAppPath
            path = AppPath
        Else
            path = AppPath
        End If
    Else
        If Not Right$(path, 1) = "\" Then path = path & "\"
    End If
    
    If Not FileExist(path & "sysinfo.log") Then
        fnum = FreeFile
        Open AppPath & "sysinfo.log" For Binary As #fnum
            Put #fnum, , "NT:" & SoInfo.IsNT & vbCrLf & _
                            "Server:" & SoInfo.IsServer & vbCrLf & _
                            "x64:" & SoInfo.IsX64 & vbCrLf & _
                            "Name:" & SoInfo.windows & vbCrLf & _
                            "Version:" & SoInfo.Version & vbCrLf & _
                            "Major:" & SoInfo.VerMajor & vbCrLf & _
                            "Minor:" & SoInfo.VerMinor & vbCrLf & _
                            "Rev:" & SoInfo.VerRevision
        Close #fnum
    End If
End Sub
#End If

Public Function InitEnv(Optional ByVal OnErrEnd As Boolean, Optional ByVal OnErrMsg As Boolean) As Boolean
On Error GoTo ErrH
    'para evitar error con "enviar informe de errores" en windows xp por usar manifest (es por las dudas)
    m_hMod = LoadLibrary("shell32.dll")
    
    'Iniciamos el UI de windows (xp+)
    InitCommonControls

    'Cargar la vatiable AppPath con App.Path & "\"
    LoadAppPath
    
    Call DetectIDE(OnErrEnd, OnErrMsg)
    
    InitEnv = True
    Exit Function
ErrH:
    InitEnv = False
    If OnErrMsg = True Then MsgBox "Se ha producido un error al iniciar." & vbCrLf & "Error Initializing Enviroment" & vbCrLf & "Error: " & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "Error"
    If OnErrEnd = True Then End
End Function

Public Function TerminateEnv(Optional ByVal OnErrEnd As Boolean, Optional ByVal OnErrMsg As Boolean) As Boolean
On Error GoTo ErrH
    If m_hMod Then
        FreeLibrary m_hMod
    End If
    TerminateEnv = True
    Exit Function
ErrH:
    TerminateEnv = False
    If OnErrMsg = True Then MsgBox "Se ha producido un error al cerrar." & vbCrLf & "Error Unload Enviroment" & vbCrLf & "Error: " & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "Error"
    If OnErrEnd = True Then End
End Function
 
Public Sub LoadAppPath()
    Dim Ret As String
    Ret = App.path
    If Len(Ret) = 3 Or Right$(Ret, 1) = "\" Then
        AppPath = Ret
    Else
        AppPath = Ret & "\"
    End If
End Sub



'Sub Main()
'
'modBasics.InitEnv True, True
'
'
'Call DoManifest("Organization.Division.Name", AppPath, eWindowsReadyManifest.WinAll, eExecutionLevel.asInvoker, True)
'
'modBasics.TerminateEnv True
'End Sub

Public Sub DoManifest(Optional ByVal AppName As String, Optional ByVal path As String, Optional ByVal WindowsReadyManifest As Long = eWindowsReadyManifest.WinAll, Optional ByVal ExecutionLevel As eExecutionLevel = asInvoker, Optional ByVal CloseAndRun As Boolean = False)
    On Error Resume Next
    
    If path = "" Then
        LoadAppPath
        path = AppPath
    Else
        path = AppPath
    End If
    
    If Not FileExist(path & App.EXEName & ".exe.manifest") Then
        Dim fnum As Long
        fnum = FreeFile
        Open App.EXEName & ".exe.manifest" For Binary As #fnum
            Put #fnum, , "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
            Put #fnum, , "<assembly manifestVersion=""1.0"" xmlns=""urn:schemas-microsoft-com:asm.v1"" xmlns:asmv3=""urn:schemas-microsoft-com:asm.v3"">" & vbCrLf
            Put #fnum, , "" & vbCrLf
            
            
            'Datos de la App >>
            Put #fnum, , "    <assemblyIdentity" & vbCrLf
            
            If AppName = "" Then
                'APPEXENAME
                Put #fnum, , "        name=""" & App.EXEName & ".exe""" & vbCrLf
            Else
                'Organization.Division.Name
                Put #fnum, , "        name=""" & AppName & """" & vbCrLf
            End If
            
            Put #fnum, , "        processorArchitecture=""x86""" & vbCrLf
            Put #fnum, , "        version=""" & App.Major & "." & App.Minor & ".0." & Format$("0000", App.Revision) & """" & vbCrLf
            Put #fnum, , "        type=""win32""" & vbCrLf
            Put #fnum, , "    />" & vbCrLf
            Put #fnum, , "    <description>" & App.Title & " " & App.LegalCopyright & "</description>" & vbCrLf
            'Datos de la App <<
            
            'Microsoft Windows Common Controls (Themes)
            Put #fnum, , "    <dependency>" & vbCrLf
            Put #fnum, , "        <dependentAssembly>" & vbCrLf
            Put #fnum, , "            <assemblyIdentity" & vbCrLf
            Put #fnum, , "                type=""win32""" & vbCrLf
            Put #fnum, , "                name=""Microsoft.Windows.Common-Controls""" & vbCrLf
            Put #fnum, , "                version=""6.0.0.0""" & vbCrLf
            Put #fnum, , "                processorArchitecture=""x86""" & vbCrLf
            Put #fnum, , "                publicKeyToken=""6595b64144ccf1df""" & vbCrLf
            Put #fnum, , "                language=""*""" & vbCrLf
            Put #fnum, , "            />" & vbCrLf
            Put #fnum, , "        </dependentAssembly>" & vbCrLf
            Put #fnum, , "    </dependency>" & vbCrLf
            
            
            If ExecutionLevel = eExecutionLevel.asInvoker Or ExecutionLevel = eExecutionLevel.requireAdministrator Or ExecutionLevel = eExecutionLevel.highestAvailable Then
                Put #fnum, , "    <trustInfo xmlns=""urn:schemas-microsoft-com:asm.v3"">" & vbCrLf
                Put #fnum, , "        <security>" & vbCrLf
                Put #fnum, , "            <requestedPrivileges>" & vbCrLf
                
                
                If ExecutionLevel = eExecutionLevel.asInvoker Then
                    'Como el invocador
                    Put #fnum, , "                <requestedExecutionLevel" & vbCrLf
                    Put #fnum, , "                    level=""asInvoker""" & vbCrLf
                    Put #fnum, , "                    uiAccess=""false""" & vbCrLf
                    Put #fnum, , "                />" & vbCrLf
                End If
                
                If ExecutionLevel = eExecutionLevel.requireAdministrator Then
                    'Como administrador
                    Put #fnum, , "                <requestedExecutionLevel" & vbCrLf
                    Put #fnum, , "                    level=""requireAdministrator""" & vbCrLf
                    Put #fnum, , "                    uiAccess=""false""" & vbCrLf
                    Put #fnum, , "                />" & vbCrLf
                End If
                
                If ExecutionLevel = eExecutionLevel.highestAvailable Then
                    'el mas alto posible
                    Put #fnum, , "                <requestedExecutionLevel" & vbCrLf
                    Put #fnum, , "                    level=""highestAvailable""" & vbCrLf
                    Put #fnum, , "                    uiAccess=""false""" & vbCrLf
                    Put #fnum, , "                />" & vbCrLf
                End If
                
                Put #fnum, , "            </requestedPrivileges>" & vbCrLf
                Put #fnum, , "        </security>" & vbCrLf
                Put #fnum, , "    </trustInfo>" & vbCrLf
            End If
            
            
            If Not WindowsReadyManifest = 0 Then
                '{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a} -> Windows 10 and Windows Server 2016
                '{1f676c76-80e1-4239-95bb-83d0f6d0da78} -> Windows 8.1 and Windows Server 2012 R2
                '{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38} -> Windows 8 and Windows Server 2012
                '{35138b9a-5d96-4fbd-8e2d-a2440225f93a} -> Windows 7 and Windows Server 2008 R2
                '{e2011457-1546-43c5-a5fe-008deee3d3f0} -> Windows Vista and Windows Server 2008
                Put #fnum, , "    <compatibility xmlns=""urn:schemas-microsoft-com:compatibility.v1"">" & vbCrLf
                Put #fnum, , "        <application>" & vbCrLf
                If (WindowsReadyManifest And eWindowsReadyManifest.Win10) Or (WindowsReadyManifest And eWindowsReadyManifest.WinAll) Then
                    Put #fnum, , "            <!-- Windows 10 -->" & vbCrLf
                    Put #fnum, , "            <supportedOS Id=""{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}""/>" & vbCrLf
                End If
                If (WindowsReadyManifest And eWindowsReadyManifest.Win81) Or (WindowsReadyManifest And eWindowsReadyManifest.WinAll) Then
                    Put #fnum, , "            <!-- Windows 8.1 -->" & vbCrLf
                    Put #fnum, , "            <supportedOS Id=""{1f676c76-80e1-4239-95bb-83d0f6d0da78}""/>" & vbCrLf
                End If
                If (WindowsReadyManifest And eWindowsReadyManifest.WinVista) Or (WindowsReadyManifest And eWindowsReadyManifest.WinAll) Then
                    Put #fnum, , "            <!-- Windows Vista -->" & vbCrLf
                    Put #fnum, , "            <supportedOS Id=""{e2011457-1546-43c5-a5fe-008deee3d3f0}""/>" & vbCrLf
                End If
                If (WindowsReadyManifest And eWindowsReadyManifest.Win7) Or (WindowsReadyManifest And eWindowsReadyManifest.WinAll) Then
                    Put #fnum, , "            <!-- Windows 7 -->" & vbCrLf
                    Put #fnum, , "            <supportedOS Id=""{35138b9a-5d96-4fbd-8e2d-a2440225f93a}""/>" & vbCrLf
                End If
                If (WindowsReadyManifest And eWindowsReadyManifest.Win8) Or (WindowsReadyManifest And eWindowsReadyManifest.WinAll) Then
                    Put #fnum, , "            <!-- Windows 8 -->" & vbCrLf
                    Put #fnum, , "            <supportedOS Id=""{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}""/>" & vbCrLf
                End If
                Put #fnum, , "        </application>" & vbCrLf
                Put #fnum, , "    </compatibility>" & vbCrLf
            End If
            
            Put #fnum, , "</assembly>" & vbCrLf
        Close #fnum
    
        If CloseAndRun = True And Not IDE Then
            Shell """" & AppPath & App.EXEName & ".exe"" /comctl " & Command$
            End
        End If
    End If
    
End Sub

Public Function FileExist(file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
    On Error GoTo Err 'Resume Next
    '*****************************************************************
    'Se fija si existe el archivo
    '*****************************************************************
    If Dir(file, FileType) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
    Exit Function
Err:
    FileExist = False
End Function

Public Function PrevInstance(Optional ByVal sMsg As String, Optional ByVal Style As VbMsgBoxStyle, Optional ByVal sTitle As String, Optional ByVal comctlrestart As Boolean = False) As Boolean
    If App.PrevInstance = True Then
        If comctlrestart = False Or (comctlrestart = True And Not Command$ Like "*/comctl*") Then
            If Not sMsg = "" Then
                MsgBox sMsg, Style, sTitle
            End If
            ' Finaliza el programa
            PrevInstance = True
        Else
            PrevInstance = False
        End If
    Else
        PrevInstance = False
    End If
End Function

Public Function DetectIDE(Optional ByVal OnErrEnd As Boolean, Optional ByVal OnErrMsg As Boolean) As Boolean
    On Error GoTo ErrH:
    If inIDE() And RunningIDE() Then
        IDE = True
        DetectIDE = True
    ElseIf Not inIDE() And Not RunningIDE() Then
        IDE = False
        DetectIDE = True
    Else
        DetectIDE = False
        If OnErrMsg = True Then MsgBox "Se ha producido un error al iniciar." & vbCrLf & "Error Detecting Enviroment", vbCritical + vbOKOnly, "Error"
        If OnErrEnd = True Then End
    End If
    Exit Function
ErrH:
    DetectIDE = False
    If OnErrMsg = True Then MsgBox "Se ha producido un error al iniciar." & vbCrLf & "Error Detecting Enviroment" & vbCrLf & "Error: " & Err.Number & " " & Err.Description, vbCritical + vbOKOnly, "Error"
    If OnErrEnd = True Then End
End Function

Private Function inIDE() As Boolean
    inIDE = CBool(App.LogMode = 0)
End Function

Private Function RunningIDE() As Boolean
'Returns whether we are running in vb(true), or compiled (false)
 
    Static counter As Variant
    If IsEmpty(counter) Then
        counter = 1
        Debug.Assert RunningIDE() Or True
        counter = counter - 1
    ElseIf counter = 1 Then
        counter = 0
    End If
    RunningIDE = counter
 
End Function


