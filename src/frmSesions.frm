VERSION 5.00
Begin VB.Form frmSesions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de sesiones"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7245
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCloseAll 
      Caption         =   "Cerrar todas las ventanas"
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   6720
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   9375
      Left            =   0
      ScaleHeight     =   9375
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdExplorerListRefresh 
         Caption         =   "Actualizar"
         Height          =   255
         Left            =   6000
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdDeleteSess 
         Caption         =   "Borrar Sesión"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtSessName 
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   4560
         Width           =   3855
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1785
         Left            =   3240
         TabIndex        =   7
         Tag             =   "1"
         Top             =   1680
         Width           =   3855
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   3150
         Left            =   120
         Pattern         =   "*.session"
         TabIndex        =   5
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CommandButton cmdOpenSess 
         Caption         =   "Abrir Sesión"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Guardar ventanas"
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         Top             =   5040
         Width           =   2055
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   7695
         TabIndex        =   3
         Top             =   -120
         Width           =   7695
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Administrador de sesiones"
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
            TabIndex        =   4
            Top             =   480
            Width           =   6735
         End
         Begin VB.Line Line1 
            X1              =   -480
            X2              =   10080
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmSesions.frx":0000
            Top             =   480
            Width           =   480
         End
      End
      Begin VB.Line Line2 
         X1              =   3000
         X2              =   3000
         Y1              =   1320
         Y2              =   5400
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre de sesión:"
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label Label3 
         Caption         =   "Ventanas actuales:"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Sesiones:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmSesions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    Dim paths As Collection
    
    
    cmdExplorerListRefresh_Click
    
    Set paths = GetOpenExplorerPaths()

    Dim f As Integer
    Dim i As Long
    
    Dim FileToSave As String
    FileToSave = AppPath & txtSessName.Text & ".session"
    
    If Not IsValidPath(FileToSave) Then
        MsgBox "El nombre de archivo escogido no es válido", vbCritical + vbOKOnly, ""
    End If
    
    If FileExist(FileToSave) Then
        Dim result
        result = MsgBox("Ya existe un archivo de sesión con el mismo nombre." & vbCrLf & "¿Desea sobrescribirlo?", vbYesNo + vbQuestion, "")
        
        If result = vbYes Then
            Kill FileToSave
        Else
            Exit Sub
        End If
    End If

    f = FreeFile
    Open FileToSave For Output As #f

    Print #f, "User Session"

    For i = 1 To paths.Count
        Print #f, paths(i)
    Next i

    Close #f
    
    File1.Refresh
    Exit Sub
ErrHandler:
    WriteLog "Error on session_save: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
End Sub

Private Sub cmdCloseAll_Click()
    CloseAllWindowsExplorer
    cmdExplorerListRefresh_Click
End Sub

Private Sub cmdDeleteSess_Click()
'    Debug.Print "" = File1.List(File1.ListIndex), File1.ListIndex,
    If File1.ListCount = 0 Or File1.ListIndex = -1 Then Exit Sub
    
    Kill AppPath & File1.List(File1.ListIndex)
    
    File1.Refresh
    'AppPath
End Sub

Private Sub cmdOpenSess_Click()
    If File1.ListCount = 0 Or File1.ListIndex = -1 Then Exit Sub
    
    RestorePreviousSession AppPath & File1.List(File1.ListIndex)
    File1.Refresh
    
    cmdExplorerListRefresh_Click
End Sub

Private Sub cmdExplorerListRefresh_Click()

    Dim paths As Collection
    
    Set paths = GetOpenExplorerPaths()

    'Dim f As Integer
    Dim i As Long

    List1.Clear
    
    For i = 1 To paths.Count
    
        Debug.Print paths(i)
        If FileExist(paths(i), vbDirectory) Then
            Dim pathParts
            pathParts = Split(paths(i), "\")
            
            Dim newElement As String
            'pathParts(Len(pathParts) - 1) &
            If UBound(pathParts) > 0 Then
                newElement = pathParts(UBound(pathParts)) & " (" & paths(i) & ")"
            Else
                newElement = "-explorer- (" & paths(i) & ")"
            End If
            
            List1.AddItem newElement
                
        ElseIf Not FileExist(paths(i), vbDirectory) And IsShellPathValid(paths(i)) Then
        
            List1.AddItem "" & CStr(ToHumanName(paths(i)))
        End If
        
    Next i

    'Close #f
    Exit Sub
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    File1.path = AppPath
    File1.Refresh
    List1.Clear
    txtSessName.Text = "Session " & Replace$(Date, "/", "") & "_" & Replace$(Time, ":", "")
    cmdExplorerListRefresh_Click
End Sub

Private Sub Timer1_Timer()
    cmdExplorerListRefresh_Click
End Sub
