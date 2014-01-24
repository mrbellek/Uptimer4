Attribute VB_Name = "modSubClass"
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
        Bias As Long
        'StandardName(32) As Integer
        StandardName As String * 33
        StandardDate As SYSTEMTIME
        StandardBias As Long
        'DaylightName(32) As Integer
        DaylightName As String * 33
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Private Const WM_USER = &H400
Public Const WM_APPBARNOTIFY = WM_USER + 100
Private Const GWL_WNDPROC = (-4)

Private Const WM_ACTIVATE = &H6
Private Const WM_GETMINMAXINFO = &H24
Private Const WM_ENTERSIZEMOVE = &H231
Private Const WM_EXITSIZEMOVE = &H232
Private Const WM_MOVING = &H216
Private Const WM_NCHITTEST = &H84
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_SIZING = &H214
Private Const WM_TIMER = &H113
Private Const WM_WINDOWPOSCHANGED = &H47

Private Const ABN_POSCHANGED = &H1
Private Const ABN_FULLSCREENAPP = &H2
Private Const ABN_WINDOWARRANGE = &H3

Private Const ABM_ACTIVATE = &H6
Private Const ABM_GETAUTOHIDEBAR = &H7
Private Const ABM_GETSTATE = &H4
Private Const ABM_GETTASKBARPOS = &H5
Private Const ABM_NEW = &H0
Private Const ABM_QUERYPOS = &H2
Private Const ABM_REMOVE = &H1
Private Const ABM_SETAUTOHIDEBAR = &H8
Private Const ABM_SETPOS = &H3
Private Const ABM_WINDOWPOSCHANGED = &H9

Private Const WM_HOTKEY = &H312
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const MOD_WIN = &H8
Private Const VK_Z = &H5A   'previous
Private Const VK_X = &H58   'play
Private Const VK_C = &H43   'pause
Private Const VK_V = &H56   'stop
Private Const VK_B = &H42   'next
Private Const VK_L = &H4C   'add file(s)
Private Const VK_O = &H4F   'open file(s)

'User-defined identifying hotkeys
Private Const HOTKEY_Z = &H1000
Private Const HOTKEY_X = &H2000
Private Const HOTKEY_C = &H3000
Private Const HOTKEY_V = &H4000
Private Const HOTKEY_B = &H5000
Private Const HOTKEY_L = &H6000
Private Const HOTKEY_O = &H7000
'WinKey hotkey IDs are &HB00 thru &HB00+99

Private Const SE_ERR_ACCESSDENIED = 5
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FILENOTFOUND = 2
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_OUTOFMEMORY = 8
Private Const SE_ERR_PATHNOTFOUND = 3
Private Const SE_ERR_SHARE = 26

Private lOldWinProc As Long
Public lSocket& 'client socket for time sync
Private sWinKey$(99)

Public Function RunningInIDE() As Boolean
    RunningInIDE = False
    'So simple, yet so smart :)
    On Error Resume Next
    Debug.Print 1 / 0
    If Err Then RunningInIDE = True
    Err.Clear
    On Error GoTo 0
End Function

Public Sub HookForm(bStartStop As Boolean)
    If bStartStop Then
        lOldWinProc = SetWindowLong(frmMain.hwnd, GWL_WNDPROC, AddressOf MyWindowProc)
    Else
        SetWindowLong frmMain.hwnd, GWL_WNDPROC, lOldWinProc
        lOldWinProc = 0
    End If
End Sub

Public Function MyWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_APPBARNOTIFY: OnAppBarNotify wParam, lParam
        'Case WM_TIMER: OnTimer <- maybe later when I get
                            ' to code an autohide function
        Case WM_HOTKEY
            If wParam = HOTKEY_Z Or wParam = HOTKEY_X Or _
               wParam = HOTKEY_C Or wParam = HOTKEY_V Or _
               wParam = HOTKEY_B Or wParam = HOTKEY_O Or _
               wParam = HOTKEY_L Then
                OnWinampHotkey wParam ', lParam
            Else
                OnWinKeyHotKey wParam ', lParam
            End If
        Case WINSOCKMSG: OnWinSockMsg wParam, lParam
    End Select
    
    MyWindowProc = CallWindowProc(lOldWinProc, hwnd, uMsg, wParam, lParam)
    
    'Select Case uMsg
    '    Case WM_ACTIVATE: OnActivate wParam
    '    Case WM_WINDOWPOSCHANGED: OnWindowPosChanged
    'End Select
End Function

Private Function OnAppBarNotify(ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case wParam
        Case ABN_FULLSCREENAPP: OnABNFullScreenAp CBool(lParam)
        Case ABN_POSCHANGED: OnABNPosChanged
        'Case ABN_WINDOWARRANGE: OnABNWindowArrange CBool(lParam)
    End Select
    OnAppBarNotify = 0
End Function

Private Sub OnActivate(wParam As Long)
    'wParam = 2 -> user gave focus to bar by clicking it
    'wParam = 1 -> system gave focus to bar
    'wParam = 0 -> system took focus from bar
    
    'Use some code to slide the bar into view if
    'I get to code an autohide function ever
End Sub

Private Sub OnABNFullScreenAp(ByVal bOpen As Boolean)
    If bOpen Then
        'Fullscreen app starting, be not always on top
        If bIgnoreFullscreenApps Then Exit Sub
        SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    Else
        'Fullscreen app closing, be always on top if set
        If RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlwaysOnTop", 1) = 1 Then
            DoEvents
            SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
        End If
    End If
End Sub

Private Sub OnABNPosChanged()
    'The taskbar or some other appbar changed it's position
    'and this bar got misplaced, so move back to appbar space
    With rctCurrentBar
        If bAlwaysOnTop Then
            SetWindowPos frmMain.hwnd, HWND_TOPMOST, .Left, .Top, .Right - .Left, .Bottom - .Top, SWP_NOACTIVATE
        Else
            SetWindowPos frmMain.hwnd, HWND_BOTTOM, .Left, .Top, .Right - .Left, .Bottom - .Top, SWP_NOACTIVATE
        End If
        'MoveWindow frmMain.hwnd, .Left, .Top, .Right - .Left, .Bottom - .Top, 1
    End With
End Sub

Private Sub OnWinampHotkey(wParam As Long)
    Select Case wParam
        Case HOTKEY_Z: frmMain.imgWinampC_MouseUp 0, 1, 0, 0, 0
        Case HOTKEY_X: frmMain.imgWinampC_MouseUp 1, 1, 0, 0, 0
        Case HOTKEY_C: frmMain.imgWinampC_MouseUp 2, 1, 0, 0, 0
        Case HOTKEY_V: frmMain.imgWinampC_MouseUp 3, 1, 0, 0, 0
        Case HOTKEY_B: frmMain.imgWinampC_MouseUp 4, 1, 0, 0, 0
        Case HOTKEY_O: frmMain.imgWinampC_MouseUp 5, 1, 0, 0, 0
        Case HOTKEY_L: frmMain.imgWinampC_MouseUp 6, 1, 0, 0, 0
    End Select
End Sub

Public Sub RegisterHotkeys(bOnOff As Boolean, iMode%)
    'iMode: sets which function keys to use:
    '  1 = CTRL + ALT
    '  2 = CTRL + SHIFT
    '  3 = WINDOWS
    
    Dim lID&, lRet&(1 To 7)
    On Error GoTo Error:
    Select Case iMode
        Case 1: lID = MOD_CONTROL + MOD_ALT
        Case 2: lID = MOD_CONTROL + MOD_SHIFT
        Case 3: lID = MOD_WIN
    End Select
    
    If bOnOff Then
        'Register
        lRet(1) = RegisterHotKey(frmMain.hwnd, HOTKEY_Z, lID, VK_Z)
        lRet(2) = RegisterHotKey(frmMain.hwnd, HOTKEY_X, lID, VK_X)
        lRet(3) = RegisterHotKey(frmMain.hwnd, HOTKEY_C, lID, VK_C)
        lRet(4) = RegisterHotKey(frmMain.hwnd, HOTKEY_V, lID, VK_V)
        lRet(5) = RegisterHotKey(frmMain.hwnd, HOTKEY_B, lID, VK_B)
        lRet(6) = RegisterHotKey(frmMain.hwnd, HOTKEY_O, lID, VK_O)
        lRet(7) = RegisterHotKey(frmMain.hwnd, HOTKEY_L, lID, VK_L)
        If lRet(1) = 0 Or lRet(2) = 0 Or _
           lRet(3) = 0 Or lRet(4) = 0 Or _
           lRet(5) = 0 Or lRet(6) = 0 Or _
           lRet(7) = 0 Then GoTo HotkeyErr:
    Else
        'Unregister
        UnregisterHotKey frmMain.hwnd, HOTKEY_Z
        UnregisterHotKey frmMain.hwnd, HOTKEY_X
        UnregisterHotKey frmMain.hwnd, HOTKEY_C
        UnregisterHotKey frmMain.hwnd, HOTKEY_V
        UnregisterHotKey frmMain.hwnd, HOTKEY_B
        UnregisterHotKey frmMain.hwnd, HOTKEY_O
        UnregisterHotKey frmMain.hwnd, HOTKEY_L
    End If
    Exit Sub
    
HotkeyErr:
    Dim sMsg$
    sMsg = "One or more of the Winamp hotkeys failed to register. " & _
         "Maybe some other program already registered it?" & vbCrLf & _
         "The hotkey combination that failed was "
    Select Case iMode
        Case 1: sMsg = sMsg & "Ctrl-Alt"
        Case 2: sMsg = sMsg & "Ctrl-Shift"
        Case 3: sMsg = sMsg & "Windows"
    End Select
    sMsg = sMsg & "-[hotkey], the hotkey(s) that failed to register was/were "
    If lRet(1) = 0 Then sMsg = sMsg & "Z, "
    If lRet(2) = 0 Then sMsg = sMsg & "X, "
    If lRet(3) = 0 Then sMsg = sMsg & "C, "
    If lRet(4) = 0 Then sMsg = sMsg & "V, "
    If lRet(5) = 0 Then sMsg = sMsg & "B, "
    If lRet(6) = 0 Then sMsg = sMsg & "O, "
    If lRet(7) = 0 Then sMsg = sMsg & "L, "
    sMsg = Left(sMsg, Len(sMsg) - 2) & "."
    MsgBox sMsg, vbExclamation, "oops"
    Exit Sub
    
Error:
    ShowError "Winamp controls", "modSubClass_RegisterHotkeys", Err.Number, Err.Description, False
End Sub

Private Sub OnWinSockMsg(ByVal lSourceSocket&, ByVal lEvent&)
    frmMenu.timServerTimeOut.Enabled = False
    frmMenu.timServerTimeOut.Tag = ""
    Select Case lEvent
        Case FD_CONNECT
            'yay, we're connected to the server
        Case FD_WRITE
            'socket ready to send data
            '(no data to send when doing time sync, server
            ' sends time without command - yay!)
        Case FD_READ
            'data is waiting to be read
            frmMain.lblTime.Caption = "read..."
            DoEvents
            Dim lDataLen&, sData As String * 1024
            lDataLen = recv(lSocket, sData, 1024, 0)
            If lDataLen > 0 Then
                OnTimeServerReply Left(sData, 4)
            Else
                closesocket lSocket
            End If
        Case FD_CLOSE
            'connection was closed
            closesocket lSocket
        Case Else
            If (FD_CONNECT And lEvent) Then
                MsgBox "Connection refused. The server is not configured to listen for connections on that port.", vbExclamation, "oops"
                closesocket lSocket
            End If
    End Select
End Sub

Private Sub SendData(ByVal s&, sMsg$)
    Dim uByte() As Byte
    uByte = StrConv(sMsg, vbFromUnicode)
    If UBound(uByte) > -1 Then send s, uByte(0), UBound(uByte) + 1, 0
End Sub

Private Sub OnTimeServerReply(sNTP$)
    nTimeServerDelay = (Timer - nTimeServerDelay) / 2
    If Len(sNTP) <> 4 Then
        MsgBox "The NTP time server returned an invalid response.", vbExclamation, "oops"
        Exit Sub
    End If
    Dim uST As SYSTEMTIME, uDate As Date, nNTPTime As Double
    Dim uTimeZone As TIME_ZONE_INFORMATION
    
    On Error GoTo Error:
    nNTPTime = Asc(Mid(sNTP, 1, 1)) * 256 ^ 3
    nNTPTime = nNTPTime + Asc(Mid(sNTP, 2, 1)) * 256 ^ 2
    nNTPTime = nNTPTime + Asc(Mid(sNTP, 3, 1)) * 256
    nNTPTime = nNTPTime + Asc(Mid(sNTP, 4, 1))
    uDate = DateAdd("s", nNTPTime - CDbl(2840140800#) + CDbl(nTimeServerDelay), #1/1/1990#)
    
    GetTimeZoneInformation uTimeZone
    uDate = DateAdd("n", CDbl(-uTimeZone.Bias), uDate)
    
    nTimeServerDelay = Timer
    If MsgBox("The time server says current time is " & uDate & ". Synchronize clock to this?", vbYesNo + vbQuestion, "blah") = vbNo Then Exit Sub
    nTimeServerDelay = Timer - nTimeServerDelay
    
    uDate = DateAdd("s", CDbl(nTimeServerDelay), uDate)
    With uST
        .wYear = Year(uDate)
        .wMonth = Month(uDate)
        .wDay = Day(uDate)
        .wHour = Hour(uDate)
        .wMinute = Minute(uDate)
        .wSecond = Second(uDate)
    End With
    SetSystemTime uST
    frmMain.lblTime.Caption = Format(Time, "Hh:Mm:Ss")
    frmMain.timTime.Enabled = True
    Exit Sub
    
Error:
    ShowError "Time", "modSubClass", Err.Number, Err.Description, False
End Sub

Public Sub LoadWinKeyHotKeys(bStart As Boolean)
    Dim i%
    If bStart = False Then
        'unregister all winkey hotkeys
        For i = 0 To 99
            UnregisterHotKey frmMain.hwnd, &HB00 + i
        Next i
        Exit Sub
    End If
    
    Dim sDummy$
    'load hotkeys info into array
    For i = 0 To 99
        sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(i)), "0") & CStr(i))
        If sDummy = "" Then Exit For
        sWinKey(i) = sDummy
    Next i
    
    'apply info in hotkeys info array
    ' 0           1      2   3 4 5
    '[c:\prog.exe|-quiet|c:\|Q|1|2]
    ' |           |      |   | | |
    ' |           |      |   | | window state
    ' |           |      |   | modifier (excl Win)
    ' |           |      |   hotkey
    ' |           |      start path
    ' |           parameters
    ' program path
    
    Dim vDummy As Variant, lRet&, sHotKey$
    For i = 0 To 99
        If sWinKey(i) = "" Then Exit For
        vDummy = Split(sWinKey(i), "|")
        Select Case vDummy(4)
            Case 0: lRet = RegisterHotKey(frmMain.hwnd, &HB00 + i, MOD_WIN, Asc(vDummy(3)))
            Case 1: lRet = RegisterHotKey(frmMain.hwnd, &HB00 + i, MOD_WIN + MOD_CONTROL, Asc(vDummy(3)))
            Case 2: lRet = RegisterHotKey(frmMain.hwnd, &HB00 + i, MOD_WIN + MOD_SHIFT, Asc(vDummy(3)))
            Case 3: lRet = RegisterHotKey(frmMain.hwnd, &HB00 + i, MOD_WIN + MOD_ALT, Asc(vDummy(3)))
        End Select
        If lRet = 0 Then
            Select Case vDummy(4)
                Case 0: sHotKey = "Win + " & CStr(vDummy(3))
                Case 1: sHotKey = "Win + Ctrl + " & CStr(vDummy(3))
                Case 2: sHotKey = "Win + Shift + " & CStr(vDummy(3))
                Case 3: sHotKey = "Win + Alt + " & CStr(vDummy(3))
            End Select
            MsgBox "Unable to register the " & sHotKey & " hotkey. Possibly Windows or another program already registered it.", vbCritical, "damn"
        End If
    Next i
End Sub

Private Sub OnWinKeyHotKey(wParam&)
    'use identifier to take appropriate action
    Dim lID&, vDummy As Variant, SEI As SHELLEXECUTEINFO
    lID = wParam - &HB00
    vDummy = Split(sWinKey(lID), "|")
    ' 0           1      2   3 4 5
    '[c:\prog.exe|-quiet|c:\|Q|1|2]
    ' |           |      |   | | |
    ' |           |      |   | | window state
    ' |           |      |   | modifier (excl Win)
    ' |           |      |   hotkey
    ' |           |      start path
    ' |           parameters
    ' program path
    
    With SEI
        .cbSize = Len(SEI)
        .hwnd = frmMain.hwnd
        .fMask = SEE_MASK_FLAG_NO_UI Or SEE_MASK_NOCLOSEPROCESS
        .lpVerb = "open"
        .lpFile = CStr(vDummy(0)) & Chr(0)
        .lpParameters = CStr(vDummy(1)) & Chr(0)
        .lpDirectory = CStr(vDummy(2)) & Chr(0)
        .nShow = IIf(vDummy(5) = 0, SW_SHOWNORMAL, IIf(vDummy(5) = 1, SW_SHOWMINNOACTIVE, SW_SHOWMAXIMIZED))
    End With
    ShellExecuteEx SEI
    If SEI.hInstApp <= 32 Then
        Select Case SEI.hInstApp
            Case SE_ERR_ACCESSDENIED:    MsgBox "An error occurred executing the program: Access denied.", vbCritical, "shit"
            Case SE_ERR_ASSOCINCOMPLETE: MsgBox "An error occurred executing the program: File association incomplete.", vbCritical, "shit"
            Case SE_ERR_DLLNOTFOUND:     MsgBox "An error occurred executing the program: DLL not found.", vbCritical, "shit"
            Case SE_ERR_FILENOTFOUND:    MsgBox "An error occurred executing the program: File not found.", vbCritical, "shit"
            Case SE_ERR_NOASSOC:         MsgBox "An error occurred executing the program: No file association.", vbCritical, "shit"
            Case SE_ERR_OUTOFMEMORY:     MsgBox "An error occurred executing the program: Out of memory.", vbCritical, "shit"
            Case SE_ERR_PATHNOTFOUND:    MsgBox "An error occurred executing the program: Path not found.", vbCritical, "shit"
            Case SE_ERR_SHARE:           MsgBox "An error occurred executing the program: Sharing violation.", vbCritical, "shit"
            Case Else:                   MsgBox "An error occurred executing the program. Can't identify the error.", vbCritical, "shit"
        End Select
    End If
End Sub
