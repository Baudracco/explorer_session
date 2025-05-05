VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Windows Explorer Session Centinel"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      ScaleHeight     =   5415
      ScaleWidth      =   7695
      TabIndex        =   0
      Top             =   -120
      Width           =   7695
      Begin VB.CommandButton Command3 
         Caption         =   "Sesiones"
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CheckBox chkShutdownRestoreConfirm 
         Caption         =   "Pedir confirmación antes de restaurar una sesión despues de un cierre incorrecto"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1680
         Width           =   6495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4920
         TabIndex        =   8
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   3360
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   7695
         TabIndex        =   5
         Top             =   -120
         Width           =   7695
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmMain.frx":180F2
            Top             =   480
            Width           =   480
         End
         Begin VB.Line Line1 
            X1              =   -480
            X2              =   10080
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Windows Explorer Session Centinel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   840
            TabIndex        =   6
            Top             =   480
            Width           =   6735
         End
      End
      Begin VB.OptionButton op1 
         Caption         =   "No restaurar ninguna ventana"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   3000
         Width           =   6735
      End
      Begin VB.OptionButton op1 
         Caption         =   "Restaurar ventanas de explorador siempre (incuso si el cierre del sistema fue correcto)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   6735
      End
      Begin VB.OptionButton op1 
         Caption         =   "Restaurar ventanas de explorador si el sistema no se cerró correctamente"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   6735
      End
      Begin VB.CheckBox chkShutdownInfo 
         Caption         =   "Avisarme si el sistema se cerró indebidamente aun si no hay ventanas para restaurar"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   2040
         Width           =   6495
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Mostrar Ventana"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadCFG()
On Error GoTo ErrHandler

    If FileExist(ConfigFile) Then
        
        Dim f As Long
        f = FreeFile
        Open ConfigFile For Binary As #f
            Get #f, , ConfigData
        Close #f
    Else
        ConfigData.OptionShutdownRestoreConfirm = True
        ConfigData.OptionRestore = 0
        ConfigData.OptionShutdownInfo = True
        ConfigData.Type = "ESCFG"
        'Espacio reservado para futuras versiones.
        ConfigData.Reserved1 = 450
        ConfigData.Reserved2 = 30024
        ConfigData.Reserved3 = 3004
        ConfigData.Reserved4 = 3122004
        ConfigData.Reserved5 = 5304
        ConfigData.Reserved6 = 9043
        ConfigData.Reserved7 = "82863876300287687294332342"
        ConfigData.Reserved8 = True
        ConfigData.Reserved9 = False
        ConfigData.Separator = "\\|>"
        SaveCFG
    End If
    
    Me.chkShutdownRestoreConfirm.Value = IIf(ConfigData.OptionShutdownRestoreConfirm, vbChecked, vbUnchecked)
    Me.op1(ConfigData.OptionRestore).Value = True
    Me.chkShutdownInfo.Value = IIf(ConfigData.OptionShutdownInfo, vbChecked, vbUnchecked)
    
    Exit Sub
ErrHandler:
    WriteLog "Error on lcfg: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
    End
End Sub

Private Sub SaveCFG()
On Error GoTo ErrHandler
        Dim f As Long
        f = FreeFile
        Open ConfigFile For Binary As #f
            Put #f, , ConfigData
        Close #f
    Exit Sub
ErrHandler:
    WriteLog "Error on scfg: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
    End
End Sub

Private Sub Command1_Click()

    Dim i As Integer
    For i = 0 To op1.Count - 1
        If Me.op1(i).Value = True Then
            ConfigData.OptionRestore = i
            Exit For
        End If
    Next i

    ConfigData.OptionShutdownInfo = IIf(Me.chkShutdownInfo = vbChecked, True, False)
    ConfigData.OptionShutdownRestoreConfirm = IIf(Me.chkShutdownRestoreConfirm = vbChecked, True, False)
    
    SaveCFG
    Me.Hide
    
End Sub

Private Sub Command2_Click()
    Me.Hide
    LoadCFG
End Sub

Private Sub Command3_Click()
    frmSesions.Show vbModal, Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Me.Hide
    
    If PrevInstance("Ya está ejecutando la aplicación", vbInformation + vbOKOnly, "Instancia previa", True) = True Then End
    
    Call DoManifest("net.baudracco.utils.explorersession", AppPath, eWindowsReadyManifest.WinAll, eExecutionLevel.asInvoker, True)

    InitEnv False, False
    
    SessionFile = AppPath & "session.failsafe"
    LogFile = AppPath & "app.log"
    ConfigFile = AppPath & "config.cfg"
    
    WriteLog "Iniciando"
    
    LoadCFG
    
    InitTrayIcon Me
    
    Dim result As Integer
    result = WasCleanExit
    
    If result = 0 Then
        WriteLog "Windows no se cerró correctamente, hay ventanas para restaurar"
        
        Dim restoresSession As String
        restoresSession = SessionFile & "_restored_" & Replace$(Date, "/", "") & "_" & Replace$(Time, ":", "") & ".bak"
            
        Dim nonRestoresSession As String
        nonRestoresSession = SessionFile & "_nonrestored_" & Replace$(Date, "/", "") & "_" & Replace$(Time, ":", "") & ".bak"
        
        If ConfigData.OptionShutdownRestoreConfirm Then
            
            If ConfigData.OptionRestore = 0 Then
                If MsgBox("Se detectó una sesión anterior sin cierre limpio." & vbCrLf & "¿Querés recuperar las ventanas del Explorador?", vbYesNo + vbQuestion, "Recuperar sesión") = vbYes Then
                    WriteLog "Restaurando ventanas (con confirmación)"
                    RestorePreviousSession (SessionFile)
                    
                    Name SessionFile As restoresSession
        
                    WriteLog "Session restaurada (archivo de sesión: " & restoresSession & ")"
                            
                Else
                    Name SessionFile As nonRestoresSession
                    
                    WriteLog "Omitiendo restauracion (denegada por usuario)"
                End If
            End If
        Else
            WriteLog "Restaurando ventanas (auto)"
            RestorePreviousSession (SessionFile)
            
            Name SessionFile As restoresSession

            WriteLog "Session restaurada (archivo de sesión: " & restoresSession & ")"
        End If
    ElseIf result = 1 Then
        
        WriteLog "Windows no se cerró correctamente, no hay ventanas para restaurar"
        If ConfigData.OptionShutdownInfo = True Then MsgBox "Se ha cerrado windows de forma incorrecta. No hay ventanas para recuperar.", vbInformation, "Cierre inesperado"
    Else
        
        If ConfigData.OptionRestore = 1 Then RestorePreviousSession (SessionFile)
    End If

    WriteLog "Guardando estado inicial"
    CleanExit = False
    SaveSession
    
    WriteLog "Iniciando monitorización"
    StartMonitoring
    Exit Sub
ErrHandler:
    WriteLog "Error on load: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandler
    WriteLog "Procediendo a cierre limpio"
    CleanExit = True
    SaveSession
    FileCopy SessionFile, SessionFile & "_" & Replace$(Date, "/", "") & "_" & Replace$(Time, ":", "") & ".bak"
    WriteLog "Sessión de backup guardada, cerrando"
    
    RemoveTrayIcon
    
    TerminateEnv
    
    End
    Exit Sub
ErrHandler:
    WriteLog "Error on close: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
    End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrHandler
    Dim msg As Long
    msg = X / Screen.TwipsPerPixelX

    Select Case msg
        Case WM_RBUTTONUP
            ShowTrayMenu Me
        Case WM_LBUTTONDBLCLK
            mnuShow_Click
    End Select
    Exit Sub
ErrHandler:
    WriteLog "Error on tray: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
End Sub

Private Sub mnuShow_Click()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub
