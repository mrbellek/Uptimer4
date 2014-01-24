Attribute VB_Name = "modMisc"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub CopyMemoryGCI Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Any, ByVal Length As Long)
Public Declare Sub CopyMemoryTCP Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageCDS Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As COPYDATASTRUCT) As Long
Public Declare Function SendMessageGet Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function waveOutSetVolume Lib "Winmm" (ByVal wDeviceID As Integer, ByVal dwVolume As Long) As Integer
Public Declare Function waveOutGetVolume Lib "Winmm" (ByVal wDeviceID As Integer, dwVolume As Long) As Integer
Public Declare Function auxGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
'Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, ByVal lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExDwo Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Public Declare Function gethostbyaddr Lib "wsock32" (addr As Long, addrLen As Long, addrType As Long) As Long
Public Declare Function getpeername Lib "wsock32.dll" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
Public Declare Function accept Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, addrLen As Long) As Long
Public Declare Function connect Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
Public Declare Function recv Lib "wsock32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function send Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Public Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function bind Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function listen Lib "wsock32.dll" (ByVal s As Long, ByVal backlog As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORSTRUCT) As Long
'Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTSTRUCT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long

Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

Public Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
Public Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst32 Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext32 Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function GetIfTable Lib "IPhlpAPI" (ByRef pIfRowTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function GetTcpTable Lib "IPhlpAPI" (pTcpTable As MIB_TCPTABLE, pdwSize As Long, bOrder As Long) As Long
Public Declare Function GetUdpTable Lib "IPhlpAPI" (pUdpTable As MIB_UDPTABLE, pdwSize As Long, bOrder As Long) As Long
Public Declare Function GetTcpStatistics Lib "IPhlpAPI" (pStats As MIB_TCPSTATS) As Long
Public Declare Function GetUdpStatistics Lib "IPhlpAPI" (pStats As MIB_UDPSTATS) As Long
Public Declare Function GetIcmpStatistics Lib "IPhlpAPI" (pStats As MIBICMPINFO) As Long
Public Declare Function GetIpStatistics Lib "IPhlpAPI" (pStats As MIB_IPSTATS) As Long

Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long
'Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long
Public Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasGetEntryDialParams Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" (ByVal lpcstr As String, ByRef lprasdialparamsa As RASDIALPARAMS, ByRef lpbool As Long) As Long
Public Declare Function RasDial Lib "rasapi32.dll" Alias "RasDialA" (ByVal lprasdialextensions As Long, ByVal lpcstr As String, ByRef lprasdialparamsa As RASDIALPARAMS, ByVal dword As Long, lpvoid As Any, ByRef lphrasconn As Long) As Long
Public Declare Function RasGetConnectStatus Lib "rasapi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Public Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long

Public Declare Function LockWorkStation Lib "user32.dll" () As Long

Private Type CHOOSEFONTSTRUCT
    lStructSize As Long
    hwndOwner As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To 32) As Byte
End Type

Public Type MIB_IPSTATS
    dwForwarding As Long       ' IP forwarding enabled or disabled
    dwDefaultTTL As Long       ' default time-to-live
    dwInReceives As Long       ' datagrams received
    dwInHdrErrors As Long      ' received header errors
    dwInAddrErrors As Long     ' received address errors
    dwForwDatagrams As Long    ' datagrams forwarded
    dwInUnknownProtos As Long  ' datagrams with unknown protocol
    dwInDiscards As Long       ' received datagrams discarded
    dwInDelivers As Long       ' received datagrams delivered
    dwOutRequests As Long      '
    dwRoutingDiscards As Long  '
    dwOutDiscards As Long      ' sent datagrams discarded
    dwOutNoRoutes As Long      ' datagrams for which no route
    dwReasmTimeout As Long     ' datagrams for which all frags didn't arrive
    dwReasmReqds As Long       ' datagrams requiring reassembly
    dwReasmOks As Long         ' successful reassemblies
    dwReasmFails As Long       ' failed reassemblies
    dwFragOks As Long          ' successful fragmentations
    dwFragFails As Long        ' failed fragmentations
    dwFragCreates As Long      ' datagrams fragmented
    dwNumIf As Long           ' number of interfaces on computer
    dwNumAddr As Long         ' number of IP address on computer
      dwNumRoutes As Long       ' number of routes in routing table
End Type

Public Type MIBICMPSTATS
    dwMsgs As Long            ' number of messages
    dwErrors As Long          ' number of errors
    dwDestUnreachs As Long    ' destination unreachable messages
    dwTimeExcds As Long       ' time-to-live exceeded messages
    dwParmProbs As Long       ' parameter problem messages
    dwSrcQuenchs As Long      ' source quench messages
    dwRedirects As Long       ' redirection messages
    dwEchos As Long           ' echo requests
    dwEchoReps As Long        ' echo replies
    dwTimestamps As Long      ' timestamp requests
    dwTimestampReps As Long   ' timestamp replies
    dwAddrMasks As Long       ' address mask requests
    dwAddrMaskReps As Long    ' address mask replies
End Type

Public Type MIBICMPINFO
    icmpInStats As MIBICMPSTATS        ' stats for incoming messages
    icmpOutStats As MIBICMPSTATS       ' stats for outgoing messages
End Type

Public Type MIB_TCPSTATS
    dwRtoAlgorithm As Long    ' timeout algorithm
    dwRtoMin As Long          ' minimum timeout
    dwRtoMax As Long          ' maximum timeout
    dwMaxConn As Long         ' maximum connections
    dwActiveOpens As Long     ' active opens
    dwPassiveOpens As Long    ' passive opens
    dwAttemptFails As Long    ' failed attempts
    dwEstabResets As Long     ' establised connections reset
    dwCurrEstab As Long       ' established connections
    dwInSegs As Long          ' segments received
    dwOutSegs As Long         ' segment sent
    dwRetransSegs As Long     ' segments retransmitted
    dwInErrs As Long          ' incoming errors
    dwOutRsts As Long         ' outgoing resets
    dwNumConns As Long        ' cumulative connections
End Type

Public Type MIB_UDPSTATS
    dwInDatagrams As Long    ' received datagrams
    dwNoPorts As Long        ' datagrams for which no port
    dwInErrors As Long       ' errors on received datagrams
    dwOutDatagrams As Long   ' sent datagrams
    dwNumAddrs As Long       ' number of entries in UDP listener table
End Type

Public Type RASCONN
    dwSize As Long
    hRasConn As Long
    szEntryName(256) As Byte
    szDeviceType(16) As Byte
    szDeviceName(128) As Byte
End Type

Public Type RASENTRYNAME95
    dwSize As Long
    szEntryName(257) As Byte
End Type

Public Type RASDIALPARAMS
    dwSize As Long 'set to 1052
    szEntryName(256) As Byte
    szPhoneNumber(128) As Byte
    szCallbackNumber(128) As Byte
    szUserName(256) As Byte
    szPassword(256) As Byte
    szDomain(12) As Byte
End Type

Public Type RASCONNSTATUS95
   dwSize As Long 'set to 160 for Win9x, 288 for Win2k[/XP?]
   RasConnState As Long
   dwError As Long
   szDeviceType(16) As Byte
   szDeviceName(32) As Byte
End Type

Public Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer
   dwStrucVersionh As Integer
   dwFileVersionMSl As Integer
   dwFileVersionMSh As Integer
   dwFileVersionLSl As Integer
   dwFileVersionLSh As Integer
   dwProductVersionMSl As Integer
   dwProductVersionMSh As Integer
   dwProductVersionLSl As Integer
   dwProductVersionLSh As Integer
   dwFileFlagsMask As Long
   dwFileFlags As Long
   dwFileOS As Long
   dwFileType As Long
   dwFileSubtype As Long
   dwFileDateMS As Long
   dwFileDateLS As Long
End Type

Public Type MIB_IFROW
    wszName As String * 512
    dwIndex As Long             ' index of the interface
    dwType As Long              ' type of interface
    dwMtu As Long               ' max transmission unit
    dwSpeed As Long             ' speed of the interface
    dwPhysAddrLen As Long       ' length of physical address
    bPhysAddr As String * 8     ' physical address of adapter
    dwAdminStatus As Long       ' administrative status
    dwOperStatus As Long        ' operational status
    dwLastChange As Long        ' last time operational status changed
    dwInOctets As Long          ' octets received
    dwInUcastPkts As Long       ' unicast packets received
    dwInNUcastPkts As Long      ' non-unicast packets received
    dwInDiscards As Long        ' received packets discarded
    dwInErrors As Long          ' erroneous packets received
    dwInUnknownProtos As Long   ' unknown protocol packets received
    dwOutOctets As Long         ' octets sent
    dwOutUcastPkts As Long      ' unicast packets sent
    dwOutNUcastPkts As Long     ' non-unicast packets sent
    dwOutDiscards As Long       ' outgoing packets discarded
    dwOutErrors As Long         ' erroneous packets sent
    dwOutQLen As Long           ' output queue length
    dwDescrLen As Long          ' length of bDescr member
    bDescr As String * 256      ' interface description
End Type

Public Type MIB_TCPROW
    dwState As Long              'state of the connection
    dwLocalAddr As String * 4    'address on local computer
    dwLocalPort As String * 4    'port number on local computer
    dwRemoteAddr As String * 4   'address on remote computer
    dwRemotePort As String * 4   'port number on remote computer
End Type

Public Type MIB_TCPTABLE
    dwNumEntries As Long        'number of entries in the table
    table(100) As MIB_TCPROW    'array of TCP connections
End Type

Public Type MIB_UDPROW
  dwLocalAddr As String * 4 'address on local computer
  dwLocalPort As String * 4 'port number on local computer
End Type

Public Type MIB_UDPTABLE
  dwNumEntries As Long    'number of entries in the table
  table(100) As MIB_UDPROW   'table of MIB_UDPROW structs
End Type

Public Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As String
End Type

Private Type BROWSEINFO
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Type SHELLEXECUTEINFO
        cbSize As Long
        fMask As Long
        hwnd As Long
        lpVerb As String
        lpFile As String
        lpParameters As String
        lpDirectory As String
        nShow As Long
        hInstApp As Long
        '  Optional fields
        lpIDList As Long
        lpClass As String
        hkeyClass As Long
        dwHotKey As Long
        hIcon As Long
        hProcess As Long
End Type

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * 260
End Type

Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uid As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Private Type CHOOSECOLORSTRUCT
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As String
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    mLuid As LUID
    Attributes As Long
End Type

Private Type LARGE_INTEGER
    dwLow As Long
    dwHigh As Long
End Type

Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type

Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As LARGE_INTEGER
    dwSpare(0 To 75) As Long
End Type

Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As LARGE_INTEGER
    liKeSystemTime As LARGE_INTEGER
    liExpTimeZoneBias  As LARGE_INTEGER
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type

Public Type SYSTEM_POWER_STATUS
        ACLineStatus As Byte
        BatteryFlag As Byte
        BatteryLifePercent As Byte
        Reserved1 As Byte
        BatteryLifeTime As Long
        BatteryFullLifeTime As Long
End Type

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Public Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257
    szSystemStatus As String * 129
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrtype As Integer
    hLength As Integer
    hAddrList As Long
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uid As Long
    RECT As RECT
    hinst As Long
    lpszText As String
    lParam As Long
End Type

Private Type DEVMODE
        dmDeviceName As String * 32
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * 32
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Public Type APPBARDATA
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type

Public Type StyleStruct
    dwOld As Long
    dwNew As Long
End Type

Public Const GWL_STYLE = (-16)
Public Const WM_STYLECHANGED = &H7D

Public Const LVS_TYPEMASK = &H3
Public Const LVS_ICON = &H0
'Public Const LVS_REPORT = &H1
Public Const LVS_SMALLICON = &H2
'Public Const LVS_LIST = &H3

Private Const ABE_TOP = 1
Private Const ABE_BOTTOM = 3
Private Const ABM_ACTIVATE = &H6
Private Const ABM_GETAUTOHIDEBAR = &H7
Private Const ABM_GETSTATE = &H4
Private Const ABM_NEW = &H0
Private Const ABM_QUERYPOS = &H2
Public Const ABM_REMOVE = &H1
Private Const ABM_SETAUTOHIDEBAR = &H8
Private Const ABM_SETPOS = &H3
Private Const ABM_WINDOWPOSCHANGED = &H9

Public Const DRIVE_UNKNOWN = 0
Public Const DRIVE_NONE = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_BOTTOM = 1
Public Const HWND_TOP = 0
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_DWORD = 4
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Public Const WM_USER = &H400
Public Const WM_COMMAND = &H111
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONUP = &H205
Public Const WM_MBUTTONUP = &H208
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_WINDOWPOSCHANGED = &H47

'Private Const CW_USEDEFAULT = &H80000000
'Private Const TTF_CENTERTIP = &H2
'Private Const TTM_GETTEXTA = (WM_USER + 11)
Private Const TTF_SUBCLASS = &H10
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_BALLOON = &H40
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800

Public Const WINAMP_BUTTON1 = 40044
Public Const WINAMP_BUTTON2 = 40045
Public Const WINAMP_BUTTON3 = 40046
Public Const WINAMP_BUTTON4 = 40047
Public Const WINAMP_BUTTON5 = 40048
Public Const WINAMP_FILE_PLAY = 40029
Public Const WINAMP_BUTTON4_SHIFT = 40147
Public Const WINAMP_BUTTON1_CTRL = 40154
Public Const WINAMP_BUTTON2_CTRL = 40155
Public Const WINAMP_BUTTON4_CTRL = 40157
Public Const WINAMP_BUTTON5_CTRL = 40158
Public Const WM_WA_IPC = &H400
Public Const WM_COPYDATA = &H4A
Public Const WM_CLOSE = &H10
Public Const IPC_PLAYFILE = 100
Public Const IPC_DELETE = 101
Public Const IPC_GETLISTLENGTH = 124
Public Const IPC_SETPLAYLISTPOS = 121

Private Const BITSPIXEL = 12
Private Const HORZRES = 8
Private Const VERTRES = 10
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const CDS_UPDATEREGISTRY = &H1
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1
Private Const DISP_CHANGE_FAILED = -1
Private Const DISP_CHANGE_BADMODE = -2
Private Const DISP_CHANGE_NOTUPDATED = -3
Private Const DISP_CHANGE_BADFLAGS = -4
Private Const DISP_CHANGE_BADPARAM = -5

Public Const SPI_SCREENSAVERRUNNING = 97
Private Const SPI_GETWORKAREA = 48

Public Const VK_CAPITAL = &H14
Public Const VK_INSERT = &H2D
Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11

Public Const POWER_HIGH = &H1
Public Const POWER_LOW = &H2
Public Const POWER_CRITICAL = &H4
Public Const POWER_CHARGING = &H8
Public Const POWER_NOSYSTEMBATTERY = &H80
Public Const POWER_UNKNOWN = &HFF

Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const EWX_POWEROFF = 8
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2

Private Const CC_RGBINIT = &H1
Private Const CC_ANYCOLOR = &H100
Private Const CC_PREVENTFULLOPEN = &H4

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_RAISEDINNER = &H4
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWMINNOACTIVE = 7

Public Const FD_ACCEPT = &H8
Public Const FD_CONNECT = &H10
Public Const FD_READ = &H1
Public Const FD_CLOSE = &H20
Public Const FD_WRITE = &H2
Public Const AF_INET = 2
Public Const IPPROTO_TCP = 6
Public Const SOCK_STREAM = 1
Public Const SOCKET_ERROR = -1
Public Const INADDR_NONE = &HFFFF
Public Const INVALID_SOCKET = -1
Public Const WINSOCKMSG = 1025

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Public Const TH32CS_SNAPPROCESS As Long = 2&

Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10
Public Const MOUSEEVENTF_ABSOLUTE = &H8000

Public Const MFT_RADIOCHECK = &H200&
Public Const MIIM_TYPE = &H10

Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400

Private Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Const ERROR_NOT_SUPPORTED = 50

Public Const RASCS_Connected = &H2000
Public Const RASCS_Disconnected = &H2001

Public Const MIB_TCP_STATE_CLOSED = 0
Public Const MIB_TCP_STATE_LISTEN = 1
Public Const MIB_TCP_STATE_SYN_SENT = 2
Public Const MIB_TCP_STATE_SYN_RCVD = 3
Public Const MIB_TCP_STATE_ESTAB = 4
Public Const MIB_TCP_STATE_FIN_WAIT1 = 5
Public Const MIB_TCP_STATE_FIN_WAIT2 = 6
Public Const MIB_TCP_STATE_CLOSE_WAIT = 7
Public Const MIB_TCP_STATE_CLOSING = 8
Public Const MIB_TCP_STATE_LAST_ACK = 9
Public Const MIB_TCP_STATE_TIME_WAIT = 10
Public Const MIB_TCP_STATE_DELETE_TCB = 11

Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_SCREENFONTS = &H1
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_USESTYLE = &H80&

Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const REGULAR_FONTTYPE = &H400
Private Const SCREEN_FONTTYPE = &H2000

Private Const LOGPIXELSY = 90

Public Const SND_ASYNC = &H1
Public Const SND_ALIAS = &H10000
Public Const SND_FILENAME = &H20000

Private Const LWA_ALPHA = &H2
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

'User-defined stuff
Public Const MODULE_CDVOLUME = 1
Public Const MODULE_CPUUSAGE = 2
Public Const MODULE_DATE = 3
Public Const MODULE_DISKFREESPACE = 4
Public Const MODULE_EXITWIN = 5
Public Const MODULE_FREEPAGEFILE = 6
Public Const MODULE_FREERAM = 7
Public Const MODULE_IPS = 8
Public Const MODULE_LOCK = 9
Public Const MODULE_MASTERVOLUME = 10
Public Const MODULE_POWER = 11
Public Const MODULE_RESOLUTION = 12
Public Const MODULE_TIME = 13
Public Const MODULE_TOGGLE = 14
Public Const MODULE_UPTIME = 15
Public Const MODULE_WINAMPCONTROLS = 16
Public Const MODULE_WINVERSION = 17

Public Const MODULE_PROCESSES = 18
Public Const MODULE_MOUSEIDLE = 19
Public Const MODULE_TCPMONITOR = 20
Public Const MODULE_MSIEVERSION = 21
Public Const MODULE_DXVERSION = 22
Public Const MODULE_RAS = 23
Public Const MODULE_NETSTAT = 24

Public lBitMask&
Public sDisks$()
Public sCurrentDisk$
Public lOldVolume&, lOldVolume2&
Private lCPUHandle As Long
Public iCurrentIP As Integer
Public bIgnoreLocalHostIP As Boolean
Public sIPs() As String
Public bLocked As Boolean
Public bIsWinNT As Boolean, bIsWin95 As Boolean
Public bIsWinXP As Boolean, bIsWin2000 As Boolean
Private liOldIdleTime As LARGE_INTEGER
Private liOldSystemTime As LARGE_INTEGER
Public rctCurrentBar As RECT
Public sModules$
Public Const sKeySettings$ = "Software\Soeperman Enterprises Ltd.\Uptimer4\Settings"
Public bBarPos As Boolean '1=Top, 0=Bottom
Public bBarDocking As Boolean
Public bBarHidden As Boolean
Public bAlwaysOnTop As Boolean
Public lSockNum&, lReadSockNum&
Public bToggleKeysInit As Boolean, bToggleKeysState(1 To 4) As Boolean
Public sHostname$
Public bReminderSet As Boolean, sReminderTime$
Public lProcessToolTipHwnd&, lNetstatToolTipHwnd&
Public uOldMousePos As POINTAPI
Public iMouseIdle%, iMouseDummy%, iMouseTimeout%
Public lColorFore&, lColorBack&, lColorText&
Public lColorGraphGrid&, lColorGraph1st&, lColorGraph2nd&
Public iProcessTrunc%, iNetstatTrunc%
Public lServed&
Public bEnableMultiRows As Boolean
Public sUptimeLogLocation$
Public bUptimerLogHourly As Boolean
Public bIgnoreFullscreenApps As Boolean
Public nOldTCPUp As Single, nOldTCPDown As Single
Public bSolidGraphs As Boolean
Public bCoolSunkenButtons As Boolean
Public sWinDir$, sWinSysDir$
Public nTimeServerDelay As Single
Public iTransparency%
'Public bMenuOpen As Boolean

Public Function GetCPULoad(Optional bStop As Boolean) As Integer
    GetCPULoad = -1
    On Error GoTo Error:
    If bIsWinNT Then GoTo WinNTCPU:
    If bStop = True Then 'Stop
        Dim hKey As Long
        RegOpenKey HKEY_DYN_DATA, "PerfStats\StopStat", hKey
        If hKey = 0 Then Exit Function
        RegQueryValueEx hKey, "KERNEL\CPUUsage", 0, REG_DWORD, 0, 4
        RegCloseKey hKey
        RegCloseKey lCPUHandle
        GetCPULoad = 0
        lCPUHandle = 0
        Exit Function
    Else 'Start
        If lCPUHandle = 0 Then
            Dim hKeyStartSrv&, hKeyStopSrv&, hKeyStartStat&
            RegOpenKey HKEY_DYN_DATA, "PerfStats\StatData", lCPUHandle
            RegOpenKey HKEY_DYN_DATA, "PerfStats\StartSrv", hKeyStartSrv
            RegOpenKey HKEY_DYN_DATA, "PerfStats\StopSrv", hKeyStopSrv
            If hKeyStartSrv = 0 Or hKeyStopSrv = 0 Or lCPUHandle = 0 Then Exit Function
            RegQueryValueEx hKeyStartSrv, "KERNEL", 0, REG_DWORD, 0, 4
            RegOpenKey HKEY_DYN_DATA, "PerfStats\StartStat", hKeyStartStat
            If hKeyStartStat = 0 Then Exit Function
            RegQueryValueEx hKeyStartStat, "KERNEL\CPUUsage", 0, REG_DWORD, 0, 4
            RegCloseKey hKeyStartStat
            RegQueryValueEx hKeyStopSrv, "KERNEL", 0, REG_DWORD, 0, 4
            RegCloseKey hKeyStartSrv
            RegCloseKey hKeyStopSrv
        End If
    End If
    'Get actual CPU usage
    Dim lData As Long
    RegQueryValueEx lCPUHandle, "KERNEL\CPUUsage", 0, REG_DWORD, lData, 4
    If lData <> 0 Then GetCPULoad = CInt(lData)
    Exit Function
    
WinNTCPU:
    If bStop = True Then GetCPULoad = 0: Exit Function
    
    Dim SysTI As SYSTEM_TIME_INFORMATION
    Dim SysBI As SYSTEM_BASIC_INFORMATION
    Dim SysPI As SYSTEM_PERFORMANCE_INFORMATION
    Dim cuIdleTime As Currency, cuSysTime As Currency
    
    If liOldSystemTime.dwHigh = 0 Or liOldSystemTime.dwLow = 0 Then
        'Initialize first
        If NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SysTI), Len(SysTI), 0&) <> 0 Then
            frmMain.timCPU.Enabled = False
            MsgBox "Error initializing CPU usage monitor!", vbExclamation, "oops"
            Exit Function
        End If
        If NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SysPI), Len(SysPI), 0&) <> 0 Then
            frmMain.timCPU.Enabled = False
            MsgBox "Error initializing CPU usage monitor!", vbExclamation, "oops"
            Exit Function
        End If
        liOldIdleTime = SysPI.liIdleTime
        liOldSystemTime = SysTI.liKeSystemTime
        Exit Function
    End If
    DoEvents
    
    'Get # of processors
    If NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(SysBI), Len(SysBI), 0&) <> 0 Then
        frmMain.timCPU.Enabled = False
        MsgBox "Error getting number of CPU's in system!", vbExclamation, "oops"
        Exit Function
    End If
    
    'Get system time
    If NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(SysTI), Len(SysTI), 0&) <> 0 Then
        frmMain.timCPU.Enabled = False
        MsgBox "Error getting system time!", vbExclamation, "oops"
        Exit Function
    End If
    
    'Get idle time
    If NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(SysPI), Len(SysPI), 0&) <> 0 Then
        frmMain.timCPU.Enabled = False
        MsgBox "Error getting system idle time!", vbExclamation, "oops"
        Exit Function
    End If
    
    Dim sErrStr$
    'Get idle time and system since last update
    cuIdleTime = LI2Currency(SysPI.liIdleTime) - LI2Currency(liOldIdleTime)
    cuSysTime = LI2Currency(SysTI.liKeSystemTime) - LI2Currency(liOldSystemTime)
    
    'Get current CPU usage
    sErrStr = "cuIdleTime=cuIdleTime/cuSysTime"
    sErrStr = "cuIdleTime=" & CStr(cuIdleTime) & " cuSysTime=" & CStr(cuSysTime)
    cuIdleTime = cuIdleTime / cuSysTime
    'sErrStr = "cuIdleTime=100-cuIdleTime"
    cuIdleTime = 100.5 - cuIdleTime * 100 / SysBI.bKeNumberProcessors
    If cuIdleTime > 100 Then cuIdleTime = 100
    If cuIdleTime < 0 Then cuIdleTime = 0
    
    'Store new idle & system time for next query
    liOldIdleTime = SysPI.liIdleTime
    liOldSystemTime = SysTI.liKeSystemTime
    
    'Return CPU load
    GetCPULoad = CInt(cuIdleTime)
    Exit Function
    
Error:
    frmMain.timCPU.Enabled = False
    ShowError "Main", "GetCPULoad(" & CStr(bStop) & ") at " & sErrStr, Err.Number, Err.Description, True
End Function

Public Sub AssignBubbleTip(hwndObj&, sText$, iModule%)
    Dim hwndTip&, uTI As TOOLINFO, rctObj As RECT
    On Error GoTo Error:
    
    'Destroy previous tooltip window
    If iModule = MODULE_PROCESSES Then
        If lProcessToolTipHwnd <> 0 Then DestroyWindow lProcessToolTipHwnd
        lProcessToolTipHwnd = 0
    ElseIf iModule = MODULE_NETSTAT Then
        If lNetstatToolTipHwnd <> 0 Then DestroyWindow lNetstatToolTipHwnd
        lNetstatToolTipHwnd = 0
    End If
    
    'Create new window
    hwndTip = CreateWindowEx(0&, TOOLTIPS_CLASSA, "", TTS_ALWAYSTIP Or TTS_BALLOON, 0, 0, 0, 0, hwndObj, 0&, App.hInstance, 0&)
    If hwndTip = 0 Then
        MsgBox "Unable to create tooltip window.", vbExclamation, "oops"
        Exit Sub
    End If
    If iModule = MODULE_PROCESSES Then
        lProcessToolTipHwnd = hwndTip
    ElseIf iModule = MODULE_NETSTAT Then
        lNetstatToolTipHwnd = hwndTip
    End If
    
    SetWindowPos hwndTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    GetClientRect hwndObj, rctObj
    
    With uTI
        .cbSize = Len(uTI)
        .uFlags = TTF_SUBCLASS '+ TTF_CENTERTIP
        .hwnd = hwndObj
        .hinst = App.hInstance
        .uid = 0
        .lpszText = sText
        .RECT = rctObj
    End With
    SendMessage hwndTip, TTM_ADDTOOLA, 0, uTI
    SendMessage hwndTip, TTM_SETMAXTIPWIDTH, 0, 260
    Exit Sub
    
Error:
    ShowError "Main", "AssignBubbleTip", Err.Number, Err.Description, False
End Sub

Public Function GetFileName(bOpenOrSave As Boolean, sFilter$, Optional sDefExt$, Optional sDialogTitle$) As String
    Dim OFN As OPENFILENAME
    On Error GoTo Error:
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = frmMain.hwnd
        .lpstrFilter = Replace(sFilter, "|", Chr(0)) & Chr(0) & Chr(0)
        .lpstrFile = String(256, 0)
        .nMaxFile = 256
        If sDialogTitle <> "" Then .lpstrTitle = sDialogTitle
        .flags = OFN_HIDEREADONLY
        If bOpenOrSave Then
            'True = Open
            .flags = .flags Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
        Else
            'False = Save
            .flags = .flags Or OFN_CREATEPROMPT Or OFN_OVERWRITEPROMPT
        End If
        If sDefExt <> "" Then .lpstrDefExt = sDefExt
    End With
    If bOpenOrSave Then
        GetOpenFileName OFN
    Else
        GetSaveFileName OFN
    End If
    GetFileName = Left(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr(0)) - 1)
    Exit Function
    
Error:
    ShowError "Main", "GetFileName(" & CStr(bOpenOrSave) & "," & sFilter & "," & sDefExt & "," & sDialogTitle & ")", Err.Number, Err.Description, False
End Function

Public Sub GetWinVersion(bFriendly As Boolean, bSP As Boolean)
    Dim OSI As OSVERSIONINFO
    Dim lBuild&, sExtra$, sWinVer$
    On Error GoTo Error:
    OSI.dwOSVersionInfoSize = Len(OSI)
    GetVersionEx OSI
    sExtra = Left(OSI.szCSDVersion, InStr(OSI.szCSDVersion, Chr(0)) - 1)
    sExtra = Replace(Trim(sExtra), "Service Pack ", "SP", , , vbTextCompare)
    sExtra = Replace(sExtra, "ServicePack ", "SP", , , vbTextCompare)
    lBuild = Val("&H" & Right(Hex(OSI.dwBuildNumber), 3))
    
    sWinVer = CStr(OSI.dwMajorVersion) & "."
    sWinVer = sWinVer & IIf(OSI.dwMinorVersion < 10, "0", "") & CStr(OSI.dwMinorVersion)
    
    Select Case OSI.dwPlatformId
        Case 0        'Win00 0.00.0000
            sWinVer = "Windows Cement!"
            bSP = False
        Case 1
            sWinVer = "Win9x " & sWinVer
            bSP = False
        Case 2
            sWinVer = "WinNT " & sWinVer
            bIsWinNT = True
    End Select
    If bFriendly Then GoTo Friendly:
    frmMain.lblOS.Caption = sWinVer & "." & CStr(lBuild)
    Exit Sub
    
Friendly:
    Select Case sWinVer     'Win00 0.00.0000
        Case "Win9x 4.00"   'Win95C b1995
            sWinVer = "Win95" & OSI.szCSDVersion & " b" & CStr(lBuild)
            bIsWin95 = True
        Case "Win9x 4.10"   'Win98Gold b1998
            sWinVer = IIf(OSI.szCSDVersion = "", "Win98Gold b", "Win98SE b") & CStr(lBuild)
        Case "Win9x 4.90"   'WinME b2300
            sWinVer = "WinME b" & CStr(lBuild)
        'Case Left(sWinVer, 7) = "WinNT 4"
        Case "WinNT 4.00"   'WinNT 4
            If Not bSP Then
                sWinVer = "WinNT 4 b" & CStr(lBuild)
            Else
                sWinVer = "WinNT 4 " & sExtra
            End If
        Case "WinNT 5.00"   'Win2000 b2176
            bIsWin2000 = True
            If Not bSP Then
                sWinVer = "Win2000 b" & CStr(lBuild)
            Else
                sWinVer = "Win2000 " & sExtra
            End If
        Case "WinNT 5.01"   'WinXP b2600
            bIsWinXP = True
            If Not bSP Then
                sWinVer = "WinXP b" & CStr(lBuild)
            Else
                sWinVer = "WinXP " & sExtra
            End If
        Case Else
            sWinVer = "Unknown WinVer?"
    End Select
    frmMain.lblOS.Caption = sWinVer
    Exit Sub
    
Error:
    ShowError "Windows version", "GetWinVersion", Err.Number, Err.Description, False
End Sub

Public Sub GetCurrentDispMode()
    Dim uDC&, lHorRes&, lVerRes&, lBPP&
    On Error GoTo Error:
    uDC = CreateDC("DISPLAY", "", "", ByVal 0&)
    If uDC = 0 Then
        MsgBox "Unable to create device context for current resolution!", vbExclamation, "oops"
        Exit Sub
    End If
    lHorRes = GetDeviceCaps(uDC, HORZRES)
    lVerRes = GetDeviceCaps(uDC, VERTRES)
    lBPP = GetDeviceCaps(uDC, BITSPIXEL)
    DeleteDC uDC
    frmMain.lblResolution.Caption = CStr(lHorRes) & " x " & CStr(lVerRes) & " x " & CStr(lBPP)
    frmMain.imgResolution.ToolTipText = "Resolution: " & frmMain.lblResolution.Caption
    Exit Sub
    
Error:
    ShowError "Main", "GetCurrentDispMode", Err.Number, Err.Description, False
End Sub

Public Sub GetDispModes()
    Dim uDevMode As DEVMODE, i%, j&, sRes$
    On Error GoTo Error:
    With frmMenu
        For i = 0 To 13
            If i <> 6 Then .mnuResRes(i).Visible = False
        Next i
        For i = 0 To 3
            .mnuRes1x(i).Enabled = False
            .mnuRes2x(i).Enabled = False
            .mnuRes3x(i).Enabled = False
            .mnuRes4x(i).Enabled = False
            .mnuRes5x(i).Enabled = False
            .mnuRes6x(i).Enabled = False
            .mnuRes7x(i).Enabled = False
            .mnuRes8x(i).Enabled = False
            .mnuRes9x(i).Enabled = False
            .mnuRes10x(i).Enabled = False
            .mnuRes11x(i).Enabled = False
            .mnuRes12x(i).Enabled = False
            .mnuRes13x(i).Enabled = False
        Next i
    End With
    
    Do While EnumDisplaySettings(0&, j, uDevMode) > 0
        With uDevMode
            sRes = CStr(.dmPelsWidth) & " x " & CStr(.dmPelsHeight) & " x " & CStr(.dmBitsPerPel)
        End With
        With frmMenu
            For i = 0 To 3
                If sRes = .mnuRes1x(i).Caption Then
                    .mnuRes1x(i).Enabled = True
                    .mnuResRes(0).Visible = True
                End If
                If sRes = .mnuRes2x(i).Caption Then
                    .mnuRes2x(i).Enabled = True
                    .mnuResRes(1).Visible = True
                End If
                If sRes = .mnuRes3x(i).Caption Then
                    .mnuRes3x(i).Enabled = True
                    .mnuResRes(2).Visible = True
                End If
                If sRes = .mnuRes4x(i).Caption Then
                    .mnuRes4x(i).Enabled = True
                    .mnuResRes(3).Visible = True
                End If
                If sRes = .mnuRes5x(i).Caption Then
                    .mnuRes5x(i).Enabled = True
                    .mnuResRes(4).Visible = True
                End If
                If sRes = .mnuRes6x(i).Caption Then
                    .mnuRes6x(i).Enabled = True
                    .mnuResRes(5).Visible = True
                End If
                If sRes = .mnuRes7x(i).Caption Then
                    .mnuRes7x(i).Enabled = True
                    .mnuResRes(6).Visible = True
                End If
                If sRes = .mnuRes8x(i).Caption Then
                    .mnuRes8x(i).Enabled = True
                    .mnuResRes(7).Visible = True
                End If
                If sRes = .mnuRes9x(i).Caption Then
                    .mnuRes9x(i).Enabled = True
                    .mnuResRes(8).Visible = True
                End If
                If sRes = .mnuRes10x(i).Caption Then
                    .mnuRes10x(i).Enabled = True
                    .mnuResRes(9).Visible = True
                End If
                If sRes = .mnuRes11x(i).Caption Then
                    .mnuRes11x(i).Enabled = True
                    .mnuResRes(10).Visible = True
                End If
                If sRes = .mnuRes12x(i).Caption Then
                    .mnuRes12x(i).Enabled = True
                    .mnuResRes(11).Visible = True
                End If
                If sRes = .mnuRes13x(i).Caption Then
                    .mnuRes13x(i).Enabled = True
                    .mnuResRes(12).Visible = True
                End If
                If sRes = .mnuRes14x(i).Caption Then
                    .mnuRes14x(i).Enabled = True
                    .mnuResRes(13).Visible = True
                End If
            Next i
        End With
        j = j + 1
    Loop
    Exit Sub
    
Error:
    ShowError "Screen resolution", "GetDispModes", Err.Number, Err.Description, False
End Sub

Public Sub SetDispMode(lHor&, lVer&, lBits&, bConfirm As Boolean)
    Dim uNewDevMode As DEVMODE, uOldDevMode As DEVMODE, lRet&
    On Error GoTo Error:
    EnumDisplaySettings 0&, 0, uNewDevMode
    With uNewDevMode
        .dmPelsWidth = lHor
        .dmPelsHeight = lVer
        .dmBitsPerPel = lBits
    End With
    If bConfirm Then
        Dim uDC As Long
        uDC = CreateDC("DISPLAY", "", "", ByVal 0&)
        If uDC = 0 Then
            MsgBox "Unable to get device context to current display mode!", vbExclamation, "oops"
            Exit Sub
        End If
        EnumDisplaySettings 0&, 0, uOldDevMode
        With uOldDevMode
            .dmPelsWidth = GetDeviceCaps(uDC, HORZRES)
            .dmPelsHeight = GetDeviceCaps(uDC, VERTRES)
            .dmBitsPerPel = GetDeviceCaps(uDC, BITSPIXEL)
        End With
        DeleteDC uDC
    End If
    
    lRet = ChangeDisplaySettings(uNewDevMode, CDS_UPDATEREGISTRY)
    Select Case lRet
        Case DISP_CHANGE_SUCCESSFUL: 'Change successful!
        Case DISP_CHANGE_RESTART:    SHRestartSystemMB frmMain.hwnd, "The display resolution was changed successfully!", EWX_FORCE: Exit Sub
        Case DISP_CHANGE_FAILED:     MsgBox "Failed to set new screen resolution!", vbExclamation, "oops"
        Case DISP_CHANGE_BADMODE:    MsgBox "Your display driver does not support that mode.", vbExclamation, "oops"
        Case DISP_CHANGE_NOTUPDATED: MsgBox "Failed to update the Registry for next boot.", vbExclamation, "oops"
        Case DISP_CHANGE_BADFLAGS:   MsgBox "A bad flag was passed, could not change resolution.", vbExclamation, "oops"
        Case DISP_CHANGE_BADPARAM:   MsgBox "A bad parameter was passed, could not change resolution.", vbExclamation, "oops"
    End Select
    GetCurrentDispMode
    DoEvents
    MakeAppBar False, True
    DoEvents
    MakeAppBar True, bBarPos
    DoEvents
    If lRet <> DISP_CHANGE_SUCCESSFUL Then Exit Sub
    
    If bConfirm = True And MsgBox("Do you want to keep this display setting?", vbOKCancel + vbQuestion, "new resolution") = vbCancel Then
        ChangeDisplaySettings uOldDevMode, CDS_UPDATEREGISTRY
    End If
    Exit Sub
    
Error:
    ShowError "Screen resolution", "SetDispMode", Err.Number, Err.Description, False
End Sub

Public Function ROT13(sText$) As String
    Dim i%, iChar%
    On Error GoTo Error:
    If sText = "" Then Exit Function
    For i = 1 To Len(sText)
        iChar = Asc(Mid(sText, i, 1))
        Select Case iChar
            Case 65 To 90: If iChar + 13 > 90 Then iChar = iChar - 26
            Case 97 To 122: If iChar + 13 > 122 Then iChar = iChar - 26
        End Select
        iChar = iChar + 13
        ROT13 = ROT13 & Chr(iChar)
    Next i
    Exit Function
    
Error:
    ShowError "Lock screen", "ROT13(" & sText & ")", Err.Number, Err.Description, False
End Function

Private Function LI2Currency(uLI As LARGE_INTEGER) As Currency
    CopyMemory LI2Currency, uLI, Len(uLI)
End Function

Public Sub GetShutDownProvilege()
    If Not bIsWinNT Then Exit Sub
    Dim lProcessHandle&, lTokenHandle&, uLUID As LUID
    Dim uTokPrivs As TOKEN_PRIVILEGES
    Dim uNewTokPrivs As TOKEN_PRIVILEGES, lBuffer&
    lProcessHandle = GetCurrentProcess()
    OpenProcessToken lProcessHandle, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lTokenHandle
    LookupPrivilegeValue "", "SeShutdownPrivilege", uLUID
    uTokPrivs.PrivilegeCount = 1
    uTokPrivs.mLuid = uLUID
    uTokPrivs.Attributes = SE_PRIVILEGE_ENABLED
    AdjustTokenPrivileges lTokenHandle, False, uTokPrivs, Len(uNewTokPrivs), uNewTokPrivs, lBuffer
End Sub

Public Sub RegSetString(lRootKey&, sSubkey$, sValueName$, sValue$)
    Dim hKey As Long
    If RegCreateKeyEx(lRootKey, sSubkey, 0&, "REG_SZ", 0&, KEY_CREATE_SUB_KEY Or KEY_SET_VALUE, ByVal 0&, hKey, ByVal 0&) <> 0 Then Exit Sub
    RegSetValueEx hKey, sValueName, 0&, REG_SZ, sValue, Len(sValue) + 1
    RegCloseKey hKey
End Sub

Public Sub RegSetDword(lRootKey&, sSubkey$, sValueName$, lValue&)
    Dim hKey As Long
    If RegCreateKeyEx(lRootKey, sSubkey, 0&, "REG_DWORD", 0&, KEY_CREATE_SUB_KEY Or KEY_SET_VALUE, ByVal 0&, hKey, ByVal 0&) <> 0 Then Exit Sub
    RegSetValueExDwo hKey, sValueName, 0&, REG_DWORD, lValue, 4
    RegCloseKey hKey
End Sub

Public Function RegGetString(lRootKey&, sSubkey$, sValueName$, Optional sDefValue$) As String
    Dim hKey As Long, sBuffer$, lRet&
    If RegOpenKeyEx(lRootKey, sSubkey, 0&, KEY_QUERY_VALUE, hKey) <> 0 Then
        RegGetString = sDefValue
        Exit Function
    End If
    sBuffer = String(255, 0)
    lRet = RegQueryValueEx(hKey, sValueName, 0&, REG_SZ, ByVal sBuffer, 255)
    RegCloseKey hKey
    If lRet = 0 And InStr(sBuffer, Chr(0)) > 1 Then
        RegGetString = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        RegGetString = sDefValue
    End If
End Function

Public Function RegGetDword(lRootKey&, sSubkey$, sValueName$, Optional lDefValue&) As Long
    Dim hKey&, lRet&
    If RegOpenKeyEx(lRootKey, sSubkey, 0&, KEY_QUERY_VALUE, hKey) <> 0 Then
        RegGetDword = lDefValue
        Exit Function
    End If
    lRet = RegQueryValueEx(hKey, sValueName, 0&, REG_DWORD, RegGetDword, 4)
    If lRet <> 0 Then RegGetDword = lDefValue
    RegCloseKey hKey
End Function

Public Sub RegDelValue(lRootKey&, sSubkey$, sValueName$)
    Dim hKey&
    If RegOpenKeyEx(lRootKey, sSubkey, 0&, KEY_WRITE, hKey) <> 0 Then Exit Sub
    RegDeleteValue hKey, sValueName
    RegCloseKey hKey
End Sub

Public Sub RegDelKey(lRootKey&, sSubkey$)
    'Yeah I know, I don't really need this function :)
    RegDeleteKey lRootKey, sSubkey
End Sub

Public Sub MakeAppBar(bStartStop As Boolean, bTop As Boolean)
    Dim APD As APPBARDATA, rctBar As RECT, hwndOldBar&
    On Error GoTo Error:
    If bStartStop = False Then
        APD.cbSize = Len(APD)
        APD.hwnd = frmMain.hwnd
        APD.uCallbackMessage = WM_APPBARNOTIFY
        SHAppBarMessage ABM_REMOVE, APD
        RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "LastAppBar"
        With frmMain
            .Caption = "Uptimer4"
            DoEvents
            .Left = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "FloatX", (Screen.Width - 2970) / 2)
            .Top = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "FloatY", (Screen.Height - 6885) / 2)
            .Width = 2940
            .fraBanner.Width = 1215
        End With
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "BarDocking", 0
        bBarDocking = False
        AlignModules False
        frmMenu.mnuMainDock(0).Visible = True
        frmMenu.mnuMainDock(1).Visible = False
        Exit Sub
    End If
    
    'Destroy old appbar space if present
    hwndOldBar = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "LastAppBar", 0)
    If hwndOldBar <> 0 Then
        APD.cbSize = Len(APD)
        APD.hwnd = hwndOldBar
        SHAppBarMessage ABM_REMOVE, APD
    End If
    With rctCurrentBar
        .Left = 0
        .Right = 0
        .Top = 0
        .Bottom = 0
    End With
    
TryAgain:
    If SystemParametersInfo(SPI_GETWORKAREA, 0, rctBar, 0) = 0 Then
        MsgBox "Unable to get desktop workarea! Could not create appbar space.", vbExclamation, "oops"
        Exit Sub
    End If
    frmMain.Width = (rctBar.Right - rctBar.Left) * Screen.TwipsPerPixelX
    AlignModules True
    
    With APD
        .cbSize = Len(APD)
        .hwnd = frmMain.hwnd
        .uCallbackMessage = WM_APPBARNOTIFY
        .rc = rctBar
        .uEdge = IIf(bTop, ABE_TOP, ABE_BOTTOM)
    End With
    bBarPos = bTop
    If SHAppBarMessage(ABM_NEW, APD) = 0 Then
        If MsgBox("Unable to register new appbar space (ABM_NEW)! Try again?", vbYesNo + vbExclamation, "oops") = vbYes Then
            GoTo TryAgain:
        Else
            Exit Sub
        End If
    End If
    
    If SHAppBarMessage(ABM_QUERYPOS, APD) = 0 Then
        If MsgBox("Unable to register new appbar space (ABM_QUERYPOS)! Try again?", vbYesNo + vbExclamation, "oops") = vbYes Then
            GoTo TryAgain:
        Else
            Exit Sub
        End If
    End If
    
    With APD.rc
        If bTop Then
            .Bottom = .Top + frmMain.Height / Screen.TwipsPerPixelY
        Else
            .Top = .Bottom - frmMain.Height / Screen.TwipsPerPixelY
        End If
        rctCurrentBar.Left = .Left
        rctCurrentBar.Right = .Right
        rctCurrentBar.Top = .Top
        rctCurrentBar.Bottom = .Bottom
    End With
    If SHAppBarMessage(ABM_SETPOS, APD) = 0 Then
        If MsgBox("Unable to register new appbar space (ABM_SETPOS)! Try again?", vbYesNo + vbExclamation, "oops") = vbYes Then
            GoTo TryAgain:
        Else
            Exit Sub
        End If
    End If
    frmMain.Caption = ""
    DoEvents
    With APD.rc
        MoveWindow frmMain.hwnd, .Left, .Top, .Right - .Left, .Bottom - .Top, 1
    End With
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "LastAppBar", frmMain.hwnd
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "BarPos", Abs(CInt(bBarPos))
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "BarDocking", 1
    frmMain.fraBanner.Width = 255
    bBarDocking = True
    frmMenu.mnuMainDock(0).Visible = False
    frmMenu.mnuMainDock(1).Visible = True
    Exit Sub
    
Error:
    ShowError "Main", "MakeAppBar(" & CStr(bStartStop) & "," & CStr(bTop) & ")", Err.Number, Err.Description, False
End Sub

Public Sub AlignModules(bHorVer As Boolean)
    Dim vDisplay As Variant, iHorzSpc%, iVertSpc%, lFrmHeight&
    On Error GoTo Error:
    vDisplay = Split(sModules, ",")
    iHorzSpc = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSpacingHorz", 15 * Screen.TwipsPerPixelX)
    iVertSpc = 255 + CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSpacingVert", 105))
    
    'bHorVer = True  --> align as appbar window
    'bHorVer = False --> align as floating window
    If Not bHorVer Then GoTo AlignFloat:
    
    
    'Align as appbar window
    Dim i&, lTop&, lOldBarHeight&
    With frmMain
        'Modify special modules which look different
        'in docked/floating mode
        .fraBanner.Width = 255
        .fraResolution.Width = 255
        .fraProcesses.Width = 255
        .fraNetstat.Width = 255
        .fraTCPMonitor.Height = 255
        .fraTCPMonitor.Width = 2895 + IIf(frmMenu.mnuTCPMonitorGraph.Checked, 855, 0)
        .imgTCPMonitorArrow(0).Left = 360
        .imgTCPMonitorArrow(1).Left = 1680
        .imgTCPMonitorArrow(1).Top = 0
        .lblTCPMonitorDown.Left = 600
        .lblTCPMonitorUp.Left = 1920
        .lblTCPMonitorUp.Top = 0
        .picGraphTCP.Left = 2880
        .picGraphTCP.Height = 255
        .Caption = ""
        DoEvents
        If bBarDocking Then lOldBarHeight = .Height
        lFrmHeight = 20 * Screen.TwipsPerPixelY
        
        'Start aligning modules as appbar window
        Dim j&
        i = 15
        lTop = 15
        .fraBanner.Top = lTop
        .fraBanner.Left = i
        i = i + .fraBanner.Width + iHorzSpc
        For j = 0 To UBound(vDisplay)
            Select Case vDisplay(j)
                Case 1
                    If bEnableMultiRows And i + .fraVolume2.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraVolume2.Left = i
                    .fraVolume2.Top = lTop
                    i = i + .fraVolume2.Width + iHorzSpc
                Case 2
                    If bEnableMultiRows And i + .fraCPU.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraCPU.Left = i
                    .fraCPU.Top = lTop
                    i = i + .fraCPU.Width + iHorzSpc
                Case 3
                    If bEnableMultiRows And i + .fraDate.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraDate.Left = i
                    .fraDate.Top = lTop
                    i = i + .fraDate.Width + iHorzSpc
                Case 4
                    If bEnableMultiRows And i + .fraDiskFree.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraDiskFree.Left = i
                    .fraDiskFree.Top = lTop
                    i = i + .fraDiskFree.Width + iHorzSpc
                Case 5
                    If bEnableMultiRows And i + .fraExitWin.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraExitWin.Left = i
                    .fraExitWin.Top = lTop
                    i = i + .fraExitWin.Width + iHorzSpc
                Case 6
                    If bEnableMultiRows And i + .fraMemoryPage.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraMemoryPage.Left = i
                    .fraMemoryPage.Top = lTop
                    i = i + .fraMemoryPage.Width + iHorzSpc
                Case 7
                    If bEnableMultiRows And i + .fraMemoryRAM.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraMemoryRAM.Left = i
                    .fraMemoryRAM.Top = lTop
                    i = i + .fraMemoryRAM.Width + iHorzSpc
                Case 8
                    If bEnableMultiRows And i + .fraIPs.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraIPs.Left = i
                    .fraIPs.Top = lTop
                    i = i + .fraIPs.Width + iHorzSpc
                Case 9
                    If bEnableMultiRows And i + .fraLock.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraLock.Left = i
                    .fraLock.Top = lTop
                    i = i + .fraLock.Width + iHorzSpc
                Case 10
                    If bEnableMultiRows And i + .fraVolume.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraVolume.Left = i
                    .fraVolume.Top = lTop
                    i = i + .fraVolume.Width + iHorzSpc
                Case 11
                    If bEnableMultiRows And i + .fraPower.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraPower.Left = i
                    .fraPower.Top = lTop
                    i = i + .fraPower.Width + iHorzSpc
                Case 12
                    If bEnableMultiRows And i + .fraResolution.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraResolution.Left = i
                    .fraResolution.Top = lTop
                    i = i + .fraResolution.Width + iHorzSpc
                Case 13
                    If bEnableMultiRows And i + .fraTime.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraTime.Left = i
                    .fraTime.Top = lTop
                    i = i + .fraTime.Width + iHorzSpc
                Case 14
                    If bEnableMultiRows And i + .fraToggle.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraToggle.Left = i
                    .fraToggle.Top = lTop
                    i = i + .fraToggle.Width + iHorzSpc
                Case 15
                    If bEnableMultiRows And i + .fraUptime.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraUptime.Left = i
                    .fraUptime.Top = lTop
                    i = i + .fraUptime.Width + iHorzSpc
                Case 16
                    If bEnableMultiRows And i + .fraWinamp.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraWinamp.Left = i
                    .fraWinamp.Top = lTop
                    i = i + .fraWinamp.Width + iHorzSpc
                Case 17
                    If bEnableMultiRows And i + .fraOS.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraOS.Left = i
                    .fraOS.Top = lTop
                    i = i + .fraOS.Width + iHorzSpc
                'Add more modules below
                Case 18
                    If bEnableMultiRows And i + .fraProcesses.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraProcesses.Left = i
                    .fraProcesses.Top = lTop
                    i = i + .fraProcesses.Width + iHorzSpc
                Case 19
                    If bEnableMultiRows And i + .fraMouseIdle.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraMouseIdle.Left = i
                    .fraMouseIdle.Top = lTop
                    i = i + .fraMouseIdle.Width + iHorzSpc
                Case 20
                    If bEnableMultiRows And i + .fraTCPMonitor.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraTCPMonitor.Left = i
                    .fraTCPMonitor.Top = lTop
                    i = i + .fraTCPMonitor.Width + iHorzSpc
                Case 21
                    If bEnableMultiRows And i + .fraMSIE.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraMSIE.Left = i
                    .fraMSIE.Top = lTop
                    i = i + .fraMSIE.Width + iHorzSpc
                Case 22
                    If bEnableMultiRows And i + .fraDX.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraDX.Left = i
                    .fraDX.Top = lTop
                    i = i + .fraDX.Width + iHorzSpc
                Case 23
                    If bEnableMultiRows And i + .fraRAS.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraRAS.Left = i
                    .fraRAS.Top = lTop
                    i = i + .fraRAS.Width + iHorzSpc
                Case 24
                    If bEnableMultiRows And i + .fraNetstat.Width > frmMain.Width Then
                        i = 0
                        lTop = lTop + iVertSpc
                    End If
                    .fraNetstat.Left = i
                    .fraNetstat.Top = lTop
                    i = i + .fraNetstat.Width + iHorzSpc
            End Select
            
            lFrmHeight = lTop + 285
        Next j
        .Height = lFrmHeight
    End With
    Exit Sub
    
AlignFloat:
    'Align as floating window
    
    With frmMain
        'Modify special modules that look different
        'in docked/floating state
        .fraBanner.Width = 1215
        .fraResolution.Width = 2295
        .fraProcesses.Width = 1815
        .fraNetstat.Width = 2295
        .fraTCPMonitor.Width = 2415
        .fraTCPMonitor.Height = 495
        .imgTCPMonitorArrow(0).Left = 360
        .imgTCPMonitorArrow(1).Left = 1320
        .imgTCPMonitorArrow(1).Top = 240
        .lblTCPMonitorDown.Left = 600
        .lblTCPMonitorUp.Left = 360
        .lblTCPMonitorUp.Top = 240
        .picGraphTCP.Left = 1560
        .picGraphTCP.Height = 495
        
        'Start aligning modules
        .fraBanner.Left = 120
        .fraCPU.Left = 120
        .fraDate.Left = 120
        .fraDiskFree.Left = 120
        .fraExitWin.Left = 120
        .fraIPs.Left = 120
        .fraLock.Left = 120
        .fraMemoryPage.Left = 120
        .fraMemoryRAM.Left = 120
        .fraOS.Left = 120
        .fraPower.Left = 120
        .fraResolution.Left = 120
        .fraTime.Left = 120
        .fraToggle.Left = 120
        .fraUptime.Left = 120
        .fraVolume.Left = 120
        .fraVolume2.Left = 120
        .fraWinamp.Left = 120
        'Add more modules below
        .fraProcesses.Left = 120
        .fraMouseIdle.Left = 120
        .fraTCPMonitor.Left = 120
        .fraMSIE.Left = 120
        .fraDX.Left = 120
        .fraRAS.Left = 120
        .fraNetstat.Left = 120
        
        j = 120
        .fraBanner.Top = j
        j = j + iVertSpc
        For i = 0 To UBound(vDisplay)
            Select Case vDisplay(i)
                Case 1:  .fraVolume2.Top = j
                Case 2:  .fraCPU.Top = j
                Case 3:  .fraDate.Top = j
                Case 4:  .fraDiskFree.Top = j
                Case 5:  .fraExitWin.Top = j
                Case 6:  .fraMemoryPage.Top = j
                Case 7:  .fraMemoryRAM.Top = j
                Case 8:  .fraIPs.Top = j
                Case 9:  .fraLock.Top = j
                Case 10: .fraVolume.Top = j
                Case 11: .fraPower.Top = j
                Case 12: .fraResolution.Top = j
                Case 13: .fraTime.Top = j
                Case 14: .fraToggle.Top = j
                Case 15: .fraUptime.Top = j
                Case 16: .fraWinamp.Top = j
                Case 17: .fraOS.Top = j
                'Add more modules below
                Case 18: .fraProcesses.Top = j
                Case 19: .fraMouseIdle.Top = j
                Case 20: .fraTCPMonitor.Top = j: j = j + 240
                Case 21: .fraMSIE.Top = j
                Case 22: .fraDX.Top = j
                Case 23: .fraRAS.Top = j
                Case 24: .fraNetstat.Top = j
            End Select
            j = j + iVertSpc
        Next i
        .Height = j + 360
    End With
    Exit Sub
    
Error:
    ShowError "Main", "AlignModules", Err.Number, Err.Description, False
End Sub

Public Function GetColor(lDefColor&) As Long
    Dim CCS As CHOOSECOLORSTRUCT, lColor&, uCusColors() As Byte
    On Error GoTo Error:
    lColor = lDefColor
    ReDim uCusColors(0 To 63)
    With CCS
        .lStructSize = Len(CCS)
        .hwndOwner = frmMain.hwnd 'frmSettings.hWnd
        .rgbResult = lColor
        .flags = CC_ANYCOLOR Or CC_PREVENTFULLOPEN Or CC_RGBINIT
        .lpCustColors = StrConv(uCusColors, vbUnicode)
    End With
    ChooseColor CCS
    GetColor = CCS.rgbResult
    Exit Function
    
Error:
    ShowError "Main", "GetColor(" & CStr(lDefColor) & ")", Err.Number, Err.Description, False
End Function

Public Sub Webserver(bStartStop As Boolean, sIP$, lPort&, hwndConnect&)
    On Error GoTo Error:
    If Not bStartStop Then
        'Shutdown webserver
        closesocket lSockNum
        closesocket lReadSockNum
        Exit Sub
    End If
    
    '===== Startup webserver (yay!) =====
    Dim uSockAddr As sockaddr
    
    'Create socket to use
    lSockNum = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If lSockNum = -1 Then
        MsgBox "Unable to create socket!", vbExclamation, "Uptimer4 webserver"
        lSockNum = 0
        Exit Sub
    End If
    
    'Set server port and listening IP
    With uSockAddr
        .sin_addr = inet_addr(sIP)
        If .sin_addr = -1 Then .sin_addr = 0
        .sin_family = AF_INET
        .sin_port = htons(lPort)
        .sin_zero = String(8, 0)
    End With
    
    'Bind to port
    If bind(lSockNum, uSockAddr, 16) <> 0 Then
        closesocket lSockNum
        lSockNum = 0
        MsgBox "Unable to bind to port " & CStr(lPort) & "! Maybe some other program is already using it?", vbExclamation, "Uptimer4 webserver"
        Exit Sub
    End If
    
    'Listen at port
    If listen(lSockNum, 10) <> 0 Then
        closesocket lSockNum
        lSockNum = 0
        MsgBox "Unable to listen at port " & CStr(lPort) & "!", vbExclamation, "Uptimer4 webserver"
        Exit Sub
    End If
    
    'Set output object for incoming connections
    If WSAAsyncSelect(lSockNum, hwndConnect, ByVal &H202, ByVal FD_CONNECT Or FD_ACCEPT) <> 0 Then
        closesocket lSockNum
        lSockNum = 0
        MsgBox "Unable to direct socket output to hwnd " & CStr(hwndConnect) & "!", vbExclamation, "Uptimer4 webserver"
        Exit Sub
    End If
    
    'Done! Webserver started and listening!
    Exit Sub
    
Error:
    ShowError "Main", "Webserver(" & CStr(bStartStop) & "," & sIP & "," & CStr(lPort) & "," & CStr(hwndConnect) & ")", Err.Number, Err.Description, False
End Sub

Public Sub MakeTrayIcon(bStartStop As Boolean, hwndCallback&)
    Dim NID As NOTIFYICONDATA
    With NID
        .cbSize = Len(NID)
        .hwnd = frmMenu.picTrayIcon.hwnd
        .szTip = "Uptimer4" & Chr(0)
        .hIcon = frmMain.Icon
        .uid = 1
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
    End With
    If bStartStop Then
        Shell_NotifyIcon NIM_ADD, NID
    Else
        Shell_NotifyIcon NIM_DELETE, NID
    End If
    
End Sub

Public Function BrowseForFolder$(hwndOwner&, sPrompt$)
    Dim BI As BROWSEINFO, lIDList&, sPath$
    On Error GoTo Error:
    With BI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    lIDList = SHBrowseForFolder(BI)
    If lIDList <> 0 Then
        sPath = String(MAX_PATH, 0)
        SHGetPathFromIDList lIDList, sPath
        CoTaskMemFree lIDList
        If InStr(sPath, Chr(0)) > 0 Then sPath = Left(sPath, InStr(sPath, Chr(0)) - 1)
        BrowseForFolder = sPath
    Else
        BrowseForFolder = ""
    End If
    Exit Function
    
Error:
    ShowError "Main", "BrowseForFolder(" & CStr(hwndOwner) & "," & sPrompt & ")", Err.Number, Err.Description, False
End Function

Public Sub ShowError(sMod$, sSub$, iErr, sErrDesc$, bDisable As Boolean)
    Dim sMsg$
    sMsg = "An unhandled error occurred in the " & sMod
    sMsg = sMsg & " module in the " & sSub & " sub:" & vbCrLf
    sMsg = sMsg & "Error #" & CStr(iErr) & ": "
    sMsg = sMsg & sErrDesc & vbCrLf
    If bDisable Then sMsg = sMsg & "The " & sMod & " module will be disabled." & vbCrLf
    sMsg = sMsg & vbCrLf & "If you feel like it, email me above info "
    sMsg = sMsg & "what you where doing and how" & vbCrLf
    sMsg = sMsg & "you can reproduce it. My email is klont@windhoos2000.nl"
    
    MsgBox sMsg, vbCritical, "unhandled error"
End Sub

Public Sub LogUptime()
    'Log current uptime to file in ticks
    On Error Resume Next
    If sUptimeLogLocation = "" Or Dir(sUptimeLogLocation) = "" Then
        sUptimeLogLocation = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Uptimer4.log"
    End If
    Open sUptimeLogLocation For Append As #1
        Print #1, Format(Date, "dd/mm/yyyy") & "|" & Format(Time, "Hh:Mm:Ss") & "|" & CStr(GetTickCount())
    Close #1
End Sub

Public Function GetLongestUptime(bReport As Boolean) As Long
    'Retrieve longest uptime from logfile
    Dim nUptime As Single, nLongestUptime As Single
    Dim i%, sMsg$, sLine$, sLongestUptime$, lLongestUptime&
    Dim sLongestUptimeLong$, sUptimeMoment$, sLongestUptimeMoment$
    
    On Error Resume Next
    If sUptimeLogLocation = "" Or Dir(sUptimeLogLocation) = "" Then
        sUptimeLogLocation = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Uptimer4.log"
        If Dir(sUptimeLogLocation) = "" Then GoTo NoLogFile:
    End If
    On Error GoTo Error:
    
    Open sUptimeLogLocation For Input As #1
        Do
            Line Input #1, sLine
            i = InStr(sLine, "|")
            If i <> 0 Then
                i = InStr(i + 1, sLine, "|")
                If i <> 0 Then
                    nUptime = CSng(Val(Mid(sLine, i + 1)))
                    sUptimeMoment = Left(sLine, i - 1)
                    If nUptime < 0 Then nUptime = nUptime + 2 ^ 32
                    If nUptime > nLongestUptime Then
                        nLongestUptime = nUptime
                        lLongestUptime = CLng(Val(Mid(sLine, i + 1)))
                        sLongestUptimeMoment = sUptimeMoment
                    End If
                End If
            End If
        Loop Until EOF(1)
    Close #1
NoLogFile:
    If GetTickCount() > nLongestUptime Then
        lLongestUptime = GetTickCount()
        sLongestUptimeMoment = Format(Date, "dd/mm/yyyy") & "|" & Format(Time, "Hh:Mm:Ss")
    End If
    sLongestUptime = GetDuration(lLongestUptime, False)
    sLongestUptimeLong = GetDuration(lLongestUptime, True)
    
    If bReport Then
        sMsg = "Longest uptime: " & sLongestUptime & ", or " & sLongestUptimeLong
        sMsg = sMsg & "." & vbCrLf & "This uptime was reached on: "
        sMsg = sMsg & Format(Left(sLongestUptimeMoment, InStr(sLongestUptimeMoment, "|") - 1), "Long Date")
        sMsg = sMsg & ", " & Mid(sLongestUptimeMoment, InStr(sLongestUptimeMoment, "|") + 1)
        sMsg = sMsg & "." & vbCrLf & vbCrLf & "Copy to clipboard?"
        If MsgBox(sMsg, vbInformation + vbYesNo, "longest uptime") = vbYes Then
            If RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MircColors", 0) = 0 Then
                sMsg = "Longest uptime: " & sLongestUptimeLong & ", reached on "
                sMsg = sMsg & Format(Left(sLongestUptimeMoment, InStr(sLongestUptimeMoment, "|") - 1), "Long Date")
                sMsg = sMsg & ", " & Mid(sLongestUptimeMoment, InStr(sLongestUptimeMoment, "|") + 1)
            Else
                sMsg = "Longest uptime:" & Chr(3) & "12 " & sLongestUptimeLong & Chr(3) & ", reached on "
                sMsg = sMsg & Format(Left(sLongestUptimeMoment, InStr(sLongestUptimeMoment, "|") - 1), "Long Date")
                sMsg = sMsg & ", " & Mid(sLongestUptimeMoment, InStr(sLongestUptimeMoment, "|") + 1)
            End If
            Clipboard.Clear
            Clipboard.SetText sMsg
        End If
    Else
        GetLongestUptime = lLongestUptime
    End If
    Exit Function
    
Error:
    ShowError "Uptime", "GetLongestUptime", Err.Number, Err.Description, False
End Function

Public Function GetDuration(lTicks&, bLongFormat As Boolean) As String
    Dim nDummy As Single, iDays%, iHours%, iMins%, iSecs%
    On Error GoTo Error:
    'When ticks are over 2^31 (Long limit), a negative
    'value is returned. This is the original DWORD value
    '(with a max of 2^32) minus 2^31 to make it fit into
    'a long variable.
    'So -2^31 + 1 is actually 2^31 + 1,
    'and -1 is actually 2^32-1.
    'Note that under this limitations the max uptime
    'for a Windows machine is 49 days, 17 hours,
    '2 minutes and 48 seconds. After this, the reported
    'uptime will wrap around to 0, even though the system
    'has been up for over 49 days.
    
    nDummy = lTicks \ 1000
    If lTicks < 0 Then nDummy = nDummy + (CSng(2 ^ 32) / 1000)
    
    iDays = nDummy \ (CLng(60 * 60) * 24)
    nDummy = nDummy - CLng(iDays) * 60 * 60 * 24
    
    iHours = nDummy \ CLng(60 * 60)
    nDummy = nDummy - CLng(iHours) * 60 * 60
    
    iMins = nDummy \ CLng(60)
    nDummy = nDummy - CLng(iMins) * 60
    
    iSecs = nDummy
    
    GetDuration = IIf(iDays < 10, "0", "") & CStr(iDays) & ":"
    GetDuration = GetDuration & IIf(iHours < 10, "0", "") & CStr(iHours) & ":"
    GetDuration = GetDuration & IIf(iMins < 10, "0", "") & CStr(iMins) & ":"
    GetDuration = GetDuration & IIf(iSecs < 10, "0", "") & CStr(iSecs)
    
    If bLongFormat Then
        Dim vDummy As Variant
        vDummy = Split(GetDuration, ":")
        GetDuration = CStr(CInt(vDummy(0))) & " days, "
        GetDuration = GetDuration & CStr(CInt(vDummy(1))) & " hours, "
        GetDuration = GetDuration & CStr(CInt(vDummy(2))) & " minutes, "
        GetDuration = GetDuration & CStr(CInt(vDummy(3))) & " seconds"
    End If
    Exit Function
    
Error:
    frmMain.timUptime.Enabled = False
    ShowError "Uptime", "GetDuration", Err.Number, Err.Description, True
End Function

Public Sub StartTCPMonitor()
    Dim uByte() As Byte, lSize&
    On Error GoTo Error:
    If GetIfTable(ByVal 0, lSize, 0) = ERROR_NOT_SUPPORTED Then
        frmMain.timTCPMonitor.Enabled = False
        MsgBox "The GetIfTable() API is not supported by system. The TCP Monitor module will be disabled.", vbExclamation, "oops"
        Exit Sub
    End If
    ReDim uByte(lSize)
    If GetIfTable(uByte(0), lSize, 1) <> 0 Then
        frmMain.timTCPMonitor.Enabled = False
        MsgBox "The GetIfTable() API returned an error for some reason. The TCP Monitor module will be disabled.", vbExclamation, "oops"
        Exit Sub
    End If
    
    Dim IfRowTable As MIB_IFROW, i%, j%
    nOldTCPDown = 0
    nOldTCPUp = 0
    i = 1
    CopyMemoryTCP IfRowTable, uByte(4 + (i - 1) * Len(IfRowTable)), Len(IfRowTable)
    Do
        With frmMenu
            .mnuTCPMonitorAdapter(i).Caption = Trim(Left(IfRowTable.bDescr, InStr(IfRowTable.bDescr, Chr(0)) - 1))
        End With
        nOldTCPDown = nOldTCPDown + IfRowTable.dwInOctets
        nOldTCPUp = nOldTCPUp + IfRowTable.dwOutOctets
        i = i + 1
        CopyMemoryTCP IfRowTable, uByte(4 + (i - 1) * Len(IfRowTable)), Len(IfRowTable)
    Loop Until IfRowTable.dwType = 0
            
    'i is # of adapters in system, if less then 10 (the max),
    'hide other menu items.
    If i < 10 Then
        With frmMenu
            For j = i To 10
                .mnuTCPMonitorAdapter(j).Visible = False
            Next j
        End With
    End If
    'i is # of adapters in system. If adapter to monitor
    'has been set to a number higher than that, set back
    'to 0 (monitor all adapters)
    If i < RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TCPMonitorAdapter", 0) Then
        With frmMenu
            For i = 1 To 10
                .mnuTCPMonitorAdapter(i).Checked = False
            Next i
            .mnuTCPMonitorAll.Checked = True
        End With
    End If
    Exit Sub
    
Error:
    frmMain.timTCPMonitor.Enabled = False
    ShowError "TCP Monitor", "StartTCPMonitor()", Err.Number, Err.Description, True
End Sub

Public Function LongIP2DottedIP$(lIP&)
    Dim uIP(1 To 4) As Byte
    CopyMemory uIP(1), lIP, 4
    LongIP2DottedIP = uIP(1) & "." & uIP(2) & "." & uIP(3) & "." & uIP(4)
End Function

Public Function GetPassword$(sText$, iMode%, Optional sOldPass$)
    'iMode determines what to do:
    ' 0 = two inputboxes, get old+new password, verify old, return new
    ' 1 = one inputbox, get password, verify, return 0/1
    ' 2 = one inputbox, get password, accept any, return password
    Load frmPassword
    With frmPassword
        .lblInfo(0).Caption = sText
        .lblInfo(0).Tag = CStr(iMode)
        .txtPassOld.Tag = sOldPass
        Select Case iMode
            Case 0
                .Caption = "Change password"
            Case 1, 2
                .Caption = "Enter password"
                .txtPassNew.Enabled = False
                .txtPassNew.BackColor = &H8000000F
                .lblInfo(1).Caption = "Password:"
        End Select
        frmPassword.Show 1
        GetPassword = frmMain.imgBanner.Tag
        frmMain.imgBanner.Tag = ""
    End With
End Function

Public Sub StringToBytes(ByVal s$, uBytes() As Byte)
    Dim i%
    For i = 1 To Len(s)
        uBytes(i - 1) = Asc(Mid(s, i, 1))
    Next i
    For i = Len(s) To UBound(uBytes)
        uBytes(i) = 0
    Next i
End Sub

Public Function GetFullPath$(sExeFile$)
    If InStr(sExeFile, "\") > 0 Then
        'Argument is already full path
        GetFullPath = sExeFile
        Exit Function
    End If
    
    'We need to find the full path to some executable
    'in memory. We can do this several ways:
    '1) Check Windows folder
    '2) Check Windows system folder
    '3) Check Program Files, in folder of same name
    '4) Check App Paths regkey
    '5) Search HD for executable (obviously not plausible)

    'If app is self, easy :)
    If LCase(sExeFile) = LCase(App.EXEName) & ".exe" Then
        GetFullPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & sExeFile
        Exit Function
    End If
    
    'Check Windows folder
    If Dir(sWinDir & "\" & sExeFile) <> "" Then
        GetFullPath = sWinDir & "\" & sExeFile
        Exit Function
    End If
    
    'Check Windows system folder
    If Dir(sWinSysDir & "\" & sExeFile) <> "" Then
        GetFullPath = sWinSysDir & "\" & sExeFile
        Exit Function
    End If
    
    'Check App Paths regkey
    Dim sFullPath$
    sFullPath = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\" & sExeFile, "")
    If sFullPath <> "" Then Exit Function
    
    'Check Program Files + folder
    'Maybe get Program Files path from Registry?
    'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\"ProgramFilesDir"/"ProgramFilesPath"
    If Dir(Left(sWinDir, 3) & "Program Files\" & Left(sExeFile, InStr(sExeFile, ".") - 1) & "\" & sExeFile) <> "" Then
        GetFullPath = Left(sWinDir, 3) & "Program Files\" & Left(sExeFile, InStr(sExeFile, ".") - 1) & "\" & sExeFile
        Exit Function
    End If
    
    'Maybe check Program Files folder on other drives?
    
    'Couldn't find it, return bare filename
    GetFullPath = "?:\?\" & sExeFile
End Function

Public Sub GetFont(objBox As Label, sDefFont$, iDefSize%, bDefBold As Boolean, bDefItalic As Boolean, bDefUnderline As Boolean, bDefStrike As Boolean, lDefColor&)
    Dim uCFS As CHOOSEFONTSTRUCT, uLF As LOGFONT, i%
    With uLF
        For i = 1 To Len(sDefFont)
            .lfFaceName(i) = Asc(Mid(sDefFont, i, 1))
        Next i
        .lfHeight = -MulDiv(iDefSize, GetDeviceCaps(frmMain.hdc, LOGPIXELSY), 72)
        .lfUnderline = Abs(CLng(bDefUnderline))
        .lfStrikeOut = Abs(CLng(bDefStrike))
    End With
    With uCFS
        .lStructSize = Len(uCFS)
        .lpLogFont = VarPtr(uLF)
        .flags = CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_NOSCRIPTSEL Or CF_SCREENFONTS Or CF_SELECTSCRIPT Or CF_USESTYLE
        .rgbColors = lDefColor
        .lpszStyle = String(12, 0)
        If bDefBold Then
            If bDefItalic Then
                Mid(.lpszStyle, 1, 11) = "Bold Italic"
            Else
                Mid(.lpszStyle, 1, 4) = "Bold"
            End If
        Else
            If bDefItalic Then
                Mid(.lpszStyle, 1, 6) = "Italic"
            Else
                Mid(.lpszStyle, 1, 7) = "Regular"
            End If
        End If
    End With
    
    If ChooseFont(uCFS) = 0 Then Exit Sub
    
    With objBox
        .FontName = StrConv(uLF.lfFaceName, vbUnicode)
        .FontSize = uCFS.iPointSize / 10
        .ForeColor = uCFS.rgbColors
        .FontBold = IIf(InStr(uCFS.lpszStyle, "Bold") > 0, True, False)
        .FontItalic = IIf(InStr(uCFS.lpszStyle, "Italic") > 0, True, False)
        .FontUnderline = IIf(uLF.lfUnderline = 1, True, False)
        .FontStrikethru = IIf(uLF.lfStrikeOut = 1, True, False)
    End With
End Sub

Public Sub KazaaEnhance(bOn As Boolean, iCloneID%)
    'iCloneID:
    ' 0 = Kazaa v1.4/1.5/1.6/1.7
    ' 1 = Morpheus v1.3.3
    ' 2 = Grokster v1.4/1.5
    ' 3 = RefoSearch
    
    Dim hwndKazaa&
    Dim hwndDummyWindow1&
    Dim hwndDummyWindow2&
    Dim hwndShellEmbedding&
    Dim hwndMsCtlsStatusBar1
    Dim hwndMsCtlsStatusBar2
    Dim hwndMsCtlsStatusBar3
    Dim sTitle$, lShowWindow&, bNewVersion As Boolean
    Select Case iCloneID
        Case 0: sTitle = "KaZaA"
        Case 1: sTitle = "Morpheus"
        Case 2: sTitle = "Grokster"
        Case 3: sTitle = "RefoSearch"
    End Select
    lShowWindow = IIf(bOn, 0, 1)
    
    hwndKazaa = FindWindow("KaZaA", vbNullString)
    If hwndKazaa = 0 Then
        MsgBox "Unable to find " & sTitle & " window. Make sure it's running", vbExclamation, "oops"
        Exit Sub
    End If
    
    Dim sCaption$, i&
    i = GetWindowTextLength(hwndKazaa)
    sCaption = String(i + 1, 0)
    GetWindowText hwndKazaa, sCaption, i + 1
    sCaption = Left(sCaption, InStr(sCaption, Chr(0)) - 1)
    If InStr(sCaption, sTitle) = 0 Then Exit Sub
    
    bNewVersion = IIf(FindWindowEx(hwndKazaa, 0, "#32770", "") <> 0, True, False)
    
    If Not bNewVersion Then
        'Use method for old Kazaa clones
        hwndShellEmbedding = FindWindowEx(hwndKazaa, 0, "Shell Embedding", "")
        hwndMsCtlsStatusBar1 = FindWindowEx(hwndKazaa, 0, "msctls_statusbar32", "")
        hwndMsCtlsStatusBar2 = FindWindowEx(hwndKazaa, hwndMsCtlsStatusBar1, "msctls_statusbar32", "")
        hwndMsCtlsStatusBar3 = FindWindowEx(hwndKazaa, hwndMsCtlsStatusBar2, "msctls_statusbar32", "")
        ShowWindow hwndShellEmbedding, lShowWindow
        ShowWindow hwndMsCtlsStatusBar2, lShowWindow
        ShowWindow hwndMsCtlsStatusBar3, lShowWindow
        SendMessage hwndKazaa, WM_WINDOWPOSCHANGED, 0, 0
    Else
        'Use method for new Kazaa/Grokster client
        hwndDummyWindow1 = FindWindowEx(hwndKazaa, 0, "#32770", "")
        hwndDummyWindow2 = FindWindowEx(hwndKazaa, hwndDummyWindow1, "#32770", "")
        hwndMsCtlsStatusBar1 = FindWindowEx(hwndKazaa, 0, "msctls_statusbar32", "")
        hwndMsCtlsStatusBar2 = FindWindowEx(hwndKazaa, hwndMsCtlsStatusBar1, "msctls_statusbar32", "")
        hwndMsCtlsStatusBar3 = FindWindowEx(hwndKazaa, hwndMsCtlsStatusBar2, "msctls_statusbar32", "")
        ShowWindow hwndDummyWindow1, lShowWindow
        ShowWindow hwndDummyWindow2, lShowWindow
        ShowWindow hwndMsCtlsStatusBar2, lShowWindow
        ShowWindow hwndMsCtlsStatusBar3, lShowWindow
        SendMessage hwndKazaa, WM_WINDOWPOSCHANGED, 0, 0
    End If
End Sub

Public Sub ConnectToServer(sHost$, lPort&, hwndCallback&)
    Dim uSockAddress As sockaddr, lRet&
    Dim lLongIP&, uHostEnt As HOSTENT, lAddrList&
    
    'check if host is IP
    lLongIP = inet_addr(sHost)
    If lLongIP = INADDR_NONE Then
        'nope, do a DNS lookup
        frmMain.lblTime.Caption = "lookup.."
        DoEvents
        lRet = gethostbyname(sHost)
        If lRet <> 0 Then
            'if lookup succeeded, get first long IP
            CopyMemory uHostEnt, lRet, 16
            CopyMemory lAddrList, uHostEnt.hAddrList, 4
            CopyMemory lLongIP, lAddrList, uHostEnt.hLength
        Else
            'lookup failed, hostname not found
            MsgBox "ConnectToServer: Hostname not found: " & sHost & ".", vbExclamation, "oops"
            GoTo CleanUp
        End If
    End If
    
    'create socket
    'doing this BEFORE setting up the socket options
    'prevents a GPF. God knows why.
    frmMain.lblTime.Caption = "socket.."
    DoEvents
    lSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If lSocket = INVALID_SOCKET Then
        MsgBox "ConnectToServer: Unable to create socket.", vbExclamation, "oops"
        GoTo CleanUp
    End If

    'setup socket options
    With uSockAddress
        .sin_addr = lLongIP
        .sin_family = AF_INET
        .sin_port = htons(lPort)
        
        If .sin_port = INVALID_SOCKET Then
            MsgBox "ConnectToServer: Invalid port: " & CStr(lPort) & ".", vbExclamation, "oops"
            GoTo CleanUp
        End If
    End With
    MsgBox "created and setup socket. it works!", vbExclamation, "yay!"
    GoTo CleanUp
        
    'select callback object
    frmMain.lblTime.Caption = "select.."
    DoEvents
    If WSAAsyncSelect(lSocket, hwndCallback, ByVal WINSOCKMSG, ByVal (FD_CONNECT Or FD_READ Or FD_WRITE Or FD_CLOSE)) <> 0 Then
        MsgBox "ConnectToServer: Unable to select callback  object for socket " & CStr(lSocket) & ".", vbExclamation, "oops"
        GoTo CleanUp
    End If
    
    'connect!
    frmMenu.timServerTimeOut.Interval = 2000
    frmMenu.timServerTimeOut.Enabled = True
    nTimeServerDelay = Timer
    frmMain.lblTime.Caption = "connct.."
    DoEvents
    connect lSocket, uSockAddress, 16
    lRet = WSAGetLastError()
    If lRet <> 10035 And lRet <> 0 Then
        MsgBox "ConnectToServer: Unable to connect to " & sHost & ".", vbExclamation, "oops"
        GoTo CleanUp
    End If
    
    Exit Sub
CleanUp:
    closesocket lSocket
    lSocket = 0
    frmMenu.timServerTimeOut.Enabled = False
    frmMenu.timServerTimeOut.Tag = ""
    frmMain.lblTime.Caption = Format(Time, "Hh:Mm:Ss")
    frmMain.timTime.Enabled = True
End Sub

Public Sub SetFormTransparency(hwndForm&)
    If bIsWin2000 = False And bIsWinXP = False Then Exit Sub
    
    Dim lStyle&, iTrans%
    'iTransparency ranges from 0 to 99 and denotes how
    'transparent forms are. 0 is opaque, 99 is barely
    'visible. The SetLayeredWindowAttributes API however
    'takes a number from 1 to 255, 1 being invisible and
    '255 being opaque. So we'll have to convert it.
    
    iTrans = 255 - iTransparency * (255 / 100)
    If iTrans < 2 Then iTrans = 2
    
    lStyle = GetWindowLong(hwndForm, GWL_EXSTYLE)
    lStyle = lStyle Or WS_EX_LAYERED
    SetWindowLong hwndForm, GWL_EXSTYLE, lStyle
    SetLayeredWindowAttributes hwndForm, 0, CLng(iTrans), LWA_ALPHA
End Sub
