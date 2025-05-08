Attribute VB_Name = "modMain"
Option Explicit

Public CleanExit As Boolean
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public SessionFile As String
Public LogFile As String
Public ConfigFile As String
Public ConfigData As ConfigDataType

Public Type ConfigDataType
    Type As String * 5
    VersionMajor As Integer
    VersionMinor As Integer
    VersionRev As Integer
    Reserved1 As Long
    Reserved2 As Long
    Reserved3 As Long
    Reserved4 As Long
    Reserved5 As Long
    Reserved6 As Long
    Reserved7 As String * 50
    Reserved8 As Boolean
    Reserved9 As Boolean
    Separator As String * 5
    OptionRestore As Integer
    OptionShutdownInfo As Boolean
    OptionShutdownRestoreConfirm As Boolean
End Type


'---

Private Declare Function SHParseDisplayName Lib "shell32.dll" ( _
    ByVal pszName As Long, _
    ByVal pbc As Long, _
    ppidl As Long, _
    ByVal sfgaoIn As Long, _
    psfgaoOut As Long) As Long

Private Declare Function SHGetNameFromIDList Lib "shell32.dll" ( _
    ByVal pidl As Long, _
    ByVal sigdnName As Long, _
    ppszName As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function lstrcpyW Lib "kernel32.dll" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Const S_OK = 0&
Private Const SIGDN_NORMALDISPLAY = &H0


'---

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205

Public nid As NOTIFYICONDATA

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

'---

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1

'---

Public Function ToHumanName(path As String) As String
    Dim pidl As Long
    Dim namePtr As Long
    Dim buffer As String
    Dim result As Long

    ' Si la ruta realmente existe como carpeta, devolvemos tal cual
    If Dir(path, vbDirectory) <> "" Then
        ToHumanName = path
        Exit Function
    End If

    ' Si no es válida para el shell, devolvemos error
    result = SHParseDisplayName(StrPtr(path), 0, pidl, 0, 0)
    If result <> S_OK Or pidl = 0 Then
        ToHumanName = "%%ERROR_NOT_VALID%%"
        Exit Function
    End If

    ' Obtener nombre legible desde el PIDL
    result = SHGetNameFromIDList(pidl, SIGDN_NORMALDISPLAY, namePtr)
    If result = S_OK And namePtr <> 0 Then
        buffer = String$(260, vbNullChar)
        lstrcpyW buffer, namePtr
        CoTaskMemFree namePtr
        CoTaskMemFree pidl
        ToHumanName = Replace$(buffer, Chr(0), "") 'Left$(buffer, InStr(buffer, vbNullChar) - 1)
        
    Else
        CoTaskMemFree pidl
        ToHumanName = "%%ERROR_NOT_VALID%%"
    End If
End Function

'---
Public Sub InitTrayIcon(frm As Form)
On Error GoTo ErrHandler
    With nid
        .cbSize = Len(nid)
        .hWnd = frm.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frm.Icon
        .szTip = "Explorador Recovery" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    Exit Sub
ErrHandler:
    WriteLog "Error on init_tray: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
End Sub

Public Sub RemoveTrayIcon()
On Error GoTo ErrHandler
    Shell_NotifyIcon NIM_DELETE, nid
    
    Exit Sub
ErrHandler:
    WriteLog "Error on remove_tray: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
End Sub

Public Sub ShowTrayMenu(frm As Form)
On Error GoTo ErrHandler
    Dim mnu As Menu
    SetForegroundWindow frm.hWnd
    frm.PopupMenu frm.mnuTray
    Exit Sub
ErrHandler:
    WriteLog "Error on menu_tray: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
End Sub

Public Sub StartMonitoring()
On Error GoTo ErrHandler
    Dim IntervalSleep As Long
    
    Do While Not CleanExit
        SaveSession
        AsyncSleep 10000, True
    Loop
    Exit Sub
ErrHandler:
    WriteLog "Error on monitor: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
    End
End Sub

Public Sub CloseAllWindowsExplorer()
    Dim shellApp As Object
    Dim windows As Object
    Dim item As Object
    Dim attemptCount As Integer

    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    
    ' Intentar varias veces para asegurar que todas las ventanas se cierren
    For attemptCount = 1 To 100
    
        Set windows = shellApp.windows
        
        If windows.Count = 0 Then
            'Debug.Print "No se encontraron ventanas activas."
            Exit Sub
        End If
        
        For Each item In windows
            If Not item Is Nothing Then
                ' Verifica si el item es una ventana del Explorador de Windows
                If InStr(LCase(item.FullName), "explorer.exe") > 0 Then
                    Debug.Print vbCrLf & "es explorer:", item.FullName
                
                    Debug.Print Date & " " & Time
                    Debug.Print "Cerrando ventana: " & item.LocationURL
                    item.Quit
                Else
                    Debug.Print vbCrLf & "NO es explorer:", item.FullName
                End If
                DoEvents ' Deja que el sistema procese otras tareas
            End If
        Next
        
        DoEvents
    Next attemptCount
    
End Sub


Public Function GetOpenExplorerPaths() As Collection
On Error GoTo ErrHandler
    Dim shellApp As Object
    Dim windows As Object
    Dim item As Object
    Dim paths As New Collection

    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    Set windows = shellApp.windows

    For Each item In windows
        If IsShellPathValid(item.Document.folder.Self.path) Then
            paths.Add item.Document.folder.Self.path
        End If
    Next

    Set GetOpenExplorerPaths = paths
    Exit Function
ErrHandler:
    WriteLog "Error on get_explorers: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
    End
End Function


Public Sub WriteLog(data As String)
    On Error GoTo ErrHandler
    Dim f As Integer
    Dim i As Long

    f = FreeFile
    Open LogFile For Append As #f
        Print #f, Date & " " & Time & " - " & data
    Close #f
    
    Exit Sub
ErrHandler:
    MsgBox "Error on log: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")", vbCritical, "Error irrecuperable"
    End
End Sub

Public Function IsValidPath(path As String) As Boolean
    Dim invalidChars As String
    Dim c As String * 1
    Dim i As Integer

    invalidChars = "<>""|?*" & Chr(0)  ' barra invertida doblemente escapada

    ' 1. No vacío
    If Trim(path) = "" Then
        IsValidPath = False
        Exit Function
    End If

    ' 2. Longitud máxima (Win32)
    If Len(path) > 260 Then
        IsValidPath = False
        Exit Function
    End If

    ' 3. Caracteres inválidos
    For i = 1 To Len(invalidChars)
        c = Mid(invalidChars, i, 1)
        Debug.Print c
        If InStr(path, c) > 0 Then
            IsValidPath = False
            Exit Function
        End If
    Next i
    
    
    If InStr(path, ":") <> 2 And InStr(path, ":") > 0 Then
        IsValidPath = False
        Exit Function
    End If

    ' 4. Nombres reservados (sólo para archivos o carpetas simples)
    Dim nombreSolo As String
    nombreSolo = UCase$(Mid$(path, InStrRev(path, "\") + 1))
    
    If nombreSolo Like "CON" Or _
       nombreSolo Like "PRN" Or _
       nombreSolo Like "AUX" Or _
       nombreSolo Like "NUL" Or _
       nombreSolo Like "COM#" Or _
       nombreSolo Like "LPT#" Then
        IsValidPath = False
        Exit Function
    End If

    ' Todo bien
    IsValidPath = True
End Function


Public Sub SaveSession()
On Error GoTo ErrHandler
    Dim paths As Collection
    Set paths = GetOpenExplorerPaths()

    Dim f As Integer
    Dim i As Long

    f = FreeFile
    Open SessionFile For Output As #f

    Print #f, "clean_exit=" & IIf(CleanExit, "1", "0")

    For i = 1 To paths.Count
        Print #f, paths(i)
    Next i

    Close #f
    Exit Sub
ErrHandler:
    WriteLog "Error on session_save: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
End Sub

Public Function WasCleanExit() As Integer
    Dim f As Integer
    Dim lineHeader As String
    Dim lineRuta As String
    Dim ventanaCount As Integer

    On Error GoTo ErrorHandler

    f = FreeFile
    Open SessionFile For Input As #f

    ' Leer primera línea: clean_exit=...
    Line Input #f, lineHeader

    ' Buscar la primera ruta válida (no vacía)
    ventanaCount = 0
    Do While Not EOF(f) And ventanaCount = 0
        Line Input #f, lineRuta
        If Trim(lineRuta) <> "" Then
            ventanaCount = 1
        End If
    Loop
    Close #f

    ' Determinar código de retorno
    If InStr(lineHeader, "clean_exit=1") > 0 Then
        WasCleanExit = 2 ' salida limpia
    ElseIf ventanaCount = 0 Then
        WasCleanExit = 1 ' salida sucia sin ventanas
    Else
        WasCleanExit = 0 ' salida sucia con ventanas
    End If
    Exit Function

ErrorHandler:
    WasCleanExit = 2 ' asumimos salida limpia si hay error
End Function

Public Sub RestorePreviousSession(Optional FilePath As String = "")
On Error GoTo ErrHandler
    Dim f As Integer
    Dim line As String
    
    If Not FileExist(FilePath) Then
        WriteLog "Error on session_restore: file not found (" & FilePath & ")"
        Exit Sub
    End If
    
    f = FreeFile
    Open FilePath For Input As #f
        Line Input #f, line ' salteamos la primera línea

        Do While Not EOF(f)
            DoEvents
            Line Input #f, line
            If IsShellPathValid(line) Then
                ShellExecute frmMain.hWnd, vbNullString, line, vbNullString, "C:\", SW_SHOWNORMAL
            End If
        Loop
    Close #f
    Exit Sub
ErrHandler:
    WriteLog "Error on session_restore: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
    End
End Sub


Public Function IsShellPathValid(path As String) As Boolean
On Error GoTo ErrHandler
    Dim pidl As Long
    Dim result As Long

    ' Pasamos la dirección del string
    result = SHParseDisplayName(StrPtr(path), 0, pidl, 0, 0)

    If result = S_OK And pidl <> 0 Then
        IsShellPathValid = True
        CoTaskMemFree pidl ' liberamos el PIDL
    Else
        IsShellPathValid = False
    End If
    Exit Function
ErrHandler:
    WriteLog "Error on path_check: " & Err.Number & " " & Err.Description & "(" & Format$("0000", Erl()) & ")"
    End
End Function

