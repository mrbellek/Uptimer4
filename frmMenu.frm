VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Uptimer4 popup menus"
   ClientHeight    =   180
   ClientLeft      =   165
   ClientTop       =   2445
   ClientWidth     =   3780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   180
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timServerTimeOut 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   0
   End
   Begin VB.PictureBox picTrayIcon 
      Height          =   135
      Left            =   600
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picWebserverRead 
      Height          =   135
      Left            =   360
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picWebserverAccept 
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuMainModules 
         Caption         =   "Modules..."
      End
      Begin VB.Menu mnuMainSettings 
         Caption         =   "Settings..."
      End
      Begin VB.Menu mnuMainHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuMainAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuUninstall 
         Caption         =   "Uninstall && exit"
      End
      Begin VB.Menu mnuMainStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainReload 
         Caption         =   "Reload custom icons"
      End
      Begin VB.Menu mnuMainUpdate 
         Caption         =   "Update now"
      End
      Begin VB.Menu mnuMainStr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainReset 
         Caption         =   "Reset window"
      End
      Begin VB.Menu mnuMainDock 
         Caption         =   "Dock"
         Index           =   0
      End
      Begin VB.Menu mnuMainDock 
         Caption         =   "Undock"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainHideIcon 
         Caption         =   "Hide icon"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMouseIdle 
      Caption         =   "Mouse idle time"
      Begin VB.Menu mnuMouseIdleUseTimeout 
         Caption         =   "Use idle timeout"
      End
      Begin VB.Menu mnuMouseIdleSetTimeout 
         Caption         =   "Set idle timeout..."
      End
   End
   Begin VB.Menu mnuTime 
      Caption         =   "Time"
      Begin VB.Menu mnuTimeReminder 
         Caption         =   "Alarm clock setup..."
      End
      Begin VB.Menu mnuTimeSync 
         Caption         =   "Sync with timeserver"
      End
      Begin VB.Menu mnuTimeCopy 
         Caption         =   "Copy to clipboard"
      End
   End
   Begin VB.Menu mnuDate 
      Caption         =   "Date"
      Begin VB.Menu mnuDateCopy 
         Caption         =   "Copy to clipboard"
      End
   End
   Begin VB.Menu mnuWinamp 
      Caption         =   "WinAmp"
      Begin VB.Menu mnuWinAmpStart 
         Caption         =   "Start Winamp"
      End
      Begin VB.Menu mnuWinampClose 
         Caption         =   "Close Winamp"
      End
      Begin VB.Menu mnuWinampGetVer 
         Caption         =   "Get Winamp version"
      End
      Begin VB.Menu mnuWinampStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWinAmpMin 
         Caption         =   "Start Winamp minimized"
      End
      Begin VB.Menu mnuWinampStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWinampHotkeys 
         Caption         =   "Enable hotkeys"
      End
      Begin VB.Menu mnuWinampHotkeyMode 
         Caption         =   "Use Ctrl-Alt+hotkey"
         Index           =   1
      End
      Begin VB.Menu mnuWinampHotkeyMode 
         Caption         =   "Use Ctrl-Shift+hotkey"
         Index           =   2
      End
      Begin VB.Menu mnuWinampHotkeyMode 
         Caption         =   "Use Windows+hotkey"
         Index           =   3
      End
   End
   Begin VB.Menu mnuPower 
      Caption         =   "Power"
      Begin VB.Menu mnuPowerBatChargeOnAC 
         Caption         =   "Show battery charge when on AC"
      End
      Begin VB.Menu mnuPowerBar 
         Caption         =   "Show bar"
      End
      Begin VB.Menu mnuPowerInterval 
         Caption         =   "Interval..."
      End
      Begin VB.Menu mnuPowerCopy 
         Caption         =   "Copy to clipboard"
      End
   End
   Begin VB.Menu mnuLock 
      Caption         =   "Lock screen"
      Begin VB.Menu mnuLockSetPass 
         Caption         =   "Set password..."
      End
   End
   Begin VB.Menu mnuExitWin 
      Caption         =   "ExitWin"
      Begin VB.Menu mnuExitWinConfirm 
         Caption         =   "Confirm exit"
      End
      Begin VB.Menu mnuExitWinForce 
         Caption         =   "Force exit"
      End
      Begin VB.Menu mnuExitWinStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitWinButtons 
         Caption         =   "Show logoff"
         Index           =   0
      End
      Begin VB.Menu mnuExitWinButtons 
         Caption         =   "Show reboot"
         Index           =   1
      End
      Begin VB.Menu mnuExitWinButtons 
         Caption         =   "Show suspend"
         Index           =   2
      End
      Begin VB.Menu mnuExitWinButtons 
         Caption         =   "Show shutdown"
         Index           =   3
      End
      Begin VB.Menu mnuExitWinButtons 
         Caption         =   "Show poweroff"
         Index           =   4
      End
   End
   Begin VB.Menu mnuOS 
      Caption         =   "OS"
      Begin VB.Menu mnuOSFriendly 
         Caption         =   "Friendly name"
      End
      Begin VB.Menu mnuOSCopy 
         Caption         =   "Copy to clipboard"
      End
      Begin VB.Menu mnuOSStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOSShowBuildSP 
         Caption         =   "Show build"
         Index           =   0
      End
      Begin VB.Menu mnuOSShowBuildSP 
         Caption         =   "Show Service Pack"
         Index           =   1
      End
   End
   Begin VB.Menu mnuUptime 
      Caption         =   "Uptime"
      Begin VB.Menu mnuUptimeBar 
         Caption         =   "Show bar"
      End
      Begin VB.Menu mnuUptimeGetBoot 
         Caption         =   "Get boot date/time"
      End
      Begin VB.Menu mnuUptimeLoggingGetLongest 
         Caption         =   "Get longest uptime"
      End
      Begin VB.Menu mnuUptimeLogging 
         Caption         =   "Logging"
         Begin VB.Menu mnuUptimeLoggingEnable 
            Caption         =   "Enable logging"
         End
         Begin VB.Menu mnuUptimeLoggingStr2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUptimeLoggingWriteHourly 
            Caption         =   "Write uptime every hour"
         End
         Begin VB.Menu mnuUptimeLoggingWriteNow 
            Caption         =   "Write uptime now"
         End
         Begin VB.Menu mnuUptimeLoggingStr1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUptimeLoggingView 
            Caption         =   "View log"
         End
         Begin VB.Menu mnuUptimeLoggingCleanUp 
            Caption         =   "Clean up log"
         End
         Begin VB.Menu mnuUptimeLoggingClear 
            Caption         =   "Clear log"
         End
      End
      Begin VB.Menu mnuUptimeCopy 
         Caption         =   "Copy to clipboard"
      End
   End
   Begin VB.Menu mnuMemoryRAM 
      Caption         =   "MemoryRAM"
      Begin VB.Menu mnuMemoryRAMDisplay 
         Caption         =   "Display fracture"
         Index           =   0
      End
      Begin VB.Menu mnuMemoryRAMDisplay 
         Caption         =   "Display percentage"
         Index           =   1
      End
      Begin VB.Menu mnuMemoryRAMStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMemoryRAMGraph 
         Caption         =   "Display graph"
      End
      Begin VB.Menu mnuMemoryRAMBar 
         Caption         =   "Show bar"
      End
      Begin VB.Menu mnuMemoryRAMInterval 
         Caption         =   "Interval..."
      End
      Begin VB.Menu mnuMemoryRAMCopy 
         Caption         =   "Copy to clipboard"
      End
      Begin VB.Menu mnuMemoryRAMStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMemoryRamSet 
         Caption         =   "Set amounts..."
      End
      Begin VB.Menu mnuMemoryRamFree 
         Caption         =   "Free up 1 MB"
         Index           =   0
      End
      Begin VB.Menu mnuMemoryRamFree 
         Caption         =   "Free up 5 MB"
         Index           =   1
      End
      Begin VB.Menu mnuMemoryRamFree 
         Caption         =   "Free up 10 MB"
         Index           =   2
      End
      Begin VB.Menu mnuMemoryRamFree 
         Caption         =   "Free up 20 MB"
         Index           =   3
      End
   End
   Begin VB.Menu mnuMemoryPage 
      Caption         =   "MemoryPage"
      Begin VB.Menu mnuMemoryPageDisplay 
         Caption         =   "Display fracture"
         Index           =   0
      End
      Begin VB.Menu mnuMemoryPageDisplay 
         Caption         =   "Display percentage"
         Index           =   1
      End
      Begin VB.Menu mnuMemoryPageStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMemoryPageGraph 
         Caption         =   "Show graph"
      End
      Begin VB.Menu mnuMemoryPageBar 
         Caption         =   "Show bar"
      End
      Begin VB.Menu mnuMemoryPageInterval 
         Caption         =   "Interval..."
      End
      Begin VB.Menu mnuMemoryPageCopy 
         Caption         =   "Copy to clipboard"
      End
   End
   Begin VB.Menu mnuMasterVol 
      Caption         =   "Master Volume"
      Begin VB.Menu mnuMasterVolMute 
         Caption         =   "Mute"
      End
   End
   Begin VB.Menu mnuCDVol 
      Caption         =   "CD Player Volume"
      Begin VB.Menu mnuCDVolMute 
         Caption         =   "Mute"
      End
   End
   Begin VB.Menu mnuIPs 
      Caption         =   "IP Addresses"
      Begin VB.Menu mnuIPsIgnoreLocalhost 
         Caption         =   "Ignore 127.0.0.1"
      End
      Begin VB.Menu mnuIPsCopy 
         Caption         =   "Copy to clipboard"
      End
   End
   Begin VB.Menu mnuCPU 
      Caption         =   "CPU"
      Begin VB.Menu mnuCPUDisplay 
         Caption         =   "Display percentage"
         Index           =   0
      End
      Begin VB.Menu mnuCPUDisplay 
         Caption         =   "Display decimal"
         Index           =   1
      End
      Begin VB.Menu mnuCPUStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCPUGraph 
         Caption         =   "Show graph"
      End
      Begin VB.Menu mnuCPUBar 
         Caption         =   "Show bar"
      End
      Begin VB.Menu mnuCPUInterval 
         Caption         =   "Interval..."
      End
      Begin VB.Menu mnuCPUCopy 
         Caption         =   "Copy to clipboard"
      End
   End
   Begin VB.Menu mnuDiskFreeSpace 
      Caption         =   "DiskFreeSpace"
      Begin VB.Menu mnuDiskFreeSpaceBar 
         Caption         =   "Show bar"
      End
      Begin VB.Menu mnuDiskFreeSpaceQuota 
         Caption         =   "Use disk quotas"
      End
      Begin VB.Menu mnuDiskFreeSpaceInterval 
         Caption         =   "Interval..."
      End
      Begin VB.Menu mnuDiskFreeSpaceCopy 
         Caption         =   "Copy to clipboard"
      End
      Begin VB.Menu mnuDiskFreeSpaceStr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiskFreeSpaceWakeUp 
         Caption         =   "Wake up HD"
      End
      Begin VB.Menu mnuDiskFreeSpaceStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiskFreeSpaceAll 
         Caption         =   "Total"
      End
      Begin VB.Menu mnuDiskFreeSpaceStr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "A:\"
         Index           =   0
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "B:\"
         Index           =   1
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "C:\"
         Index           =   2
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "D:\"
         Index           =   3
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "E:\"
         Index           =   4
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "F:\"
         Index           =   5
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "G:\"
         Index           =   6
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "H:\"
         Index           =   7
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "I:\"
         Index           =   8
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "J:\"
         Index           =   9
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "K:\"
         Index           =   10
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "L:\"
         Index           =   11
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "M:\"
         Index           =   12
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "N:\"
         Index           =   13
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "O:\"
         Index           =   14
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "P:\"
         Index           =   15
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "Q:\"
         Index           =   16
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "R:\"
         Index           =   17
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "S:\"
         Index           =   18
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "T:\"
         Index           =   19
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "U:\"
         Index           =   20
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "V:\"
         Index           =   21
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "W:\"
         Index           =   22
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "X:\"
         Index           =   23
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "Y:\"
         Index           =   24
      End
      Begin VB.Menu mnuDiskFreeSpaceA 
         Caption         =   "Z:\"
         Index           =   25
      End
   End
   Begin VB.Menu mnuRes 
      Caption         =   "Resolution"
      Begin VB.Menu mnuResConfirm 
         Caption         =   "Confirm change"
      End
      Begin VB.Menu mnuResCopy 
         Caption         =   "Copy to clipboard"
      End
      Begin VB.Menu mnuResStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "320 x 200"
         Index           =   0
         Begin VB.Menu mnuRes1x 
            Caption         =   "320 x 200 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes1x 
            Caption         =   "320 x 200 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes1x 
            Caption         =   "320 x 200 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes1x 
            Caption         =   "320 x 200 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "320 x 240"
         Index           =   1
         Begin VB.Menu mnuRes2x 
            Caption         =   "320 x 240 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes2x 
            Caption         =   "320 x 240 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes2x 
            Caption         =   "320 x 240 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes2x 
            Caption         =   "320 x 240 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "400 x 300"
         Index           =   2
         Begin VB.Menu mnuRes3x 
            Caption         =   "400 x 300 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes3x 
            Caption         =   "400 x 300 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes3x 
            Caption         =   "400 x 300 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes3x 
            Caption         =   "400 x 300 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "480 x 360"
         Index           =   3
         Begin VB.Menu mnuRes4x 
            Caption         =   "480 x 360 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes4x 
            Caption         =   "480 x 360 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes4x 
            Caption         =   "480 x 360 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes4x 
            Caption         =   "480 x 360 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "512 x 384"
         Index           =   4
         Begin VB.Menu mnuRes5x 
            Caption         =   "512 x 384 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes5x 
            Caption         =   "512 x 384 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes5x 
            Caption         =   "512 x 384 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes5x 
            Caption         =   "512 x 384 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "640 x 400"
         Index           =   5
         Begin VB.Menu mnuRes6x 
            Caption         =   "640 x 400 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes6x 
            Caption         =   "640 x 400 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes6x 
            Caption         =   "640 x 400 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes6x 
            Caption         =   "640 x 400 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "640 x 480"
         Index           =   6
         Begin VB.Menu mnuRes7x 
            Caption         =   "640 x 480 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes7x 
            Caption         =   "640 x 480 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes7x 
            Caption         =   "640 x 480 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes7x 
            Caption         =   "640 x 480 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "800 x 600"
         Index           =   7
         Begin VB.Menu mnuRes8x 
            Caption         =   "800 x 600 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes8x 
            Caption         =   "800 x 600 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes8x 
            Caption         =   "800 x 600 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes8x 
            Caption         =   "800 x 600 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "1024 x 768"
         Index           =   8
         Begin VB.Menu mnuRes9x 
            Caption         =   "1024 x 768 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes9x 
            Caption         =   "1024 x 768 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes9x 
            Caption         =   "1024 x 768 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes9x 
            Caption         =   "1024 x 768 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "1152 x 864"
         Index           =   9
         Begin VB.Menu mnuRes10x 
            Caption         =   "1152 x 864 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes10x 
            Caption         =   "1152 x 864 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes10x 
            Caption         =   "1152 x 864 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes10x 
            Caption         =   "1152 x 864 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "1280 x 1024"
         Index           =   10
         Begin VB.Menu mnuRes11x 
            Caption         =   "1280 x 1024 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes11x 
            Caption         =   "1280 x 1024 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes11x 
            Caption         =   "1280 x 1024 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes11x 
            Caption         =   "1280 x 1024 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "1600 x 1200"
         Index           =   11
         Begin VB.Menu mnuRes12x 
            Caption         =   "1600 x 1200 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes12x 
            Caption         =   "1600 x 1200 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes12x 
            Caption         =   "1600 x 1200 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes12x 
            Caption         =   "1600 x 1200 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "1920 x 1080"
         Index           =   12
         Begin VB.Menu mnuRes13x 
            Caption         =   "1920 x 1080 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes13x 
            Caption         =   "1920 x 1080 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes13x 
            Caption         =   "1920 x 1080 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes13x 
            Caption         =   "1920 x 1080 x 32"
            Index           =   3
         End
      End
      Begin VB.Menu mnuResRes 
         Caption         =   "2048 x 1536"
         Index           =   13
         Begin VB.Menu mnuRes14x 
            Caption         =   "2048 x 1536 x 8"
            Index           =   0
         End
         Begin VB.Menu mnuRes14x 
            Caption         =   "2048 x 1536 x 16"
            Index           =   1
         End
         Begin VB.Menu mnuRes14x 
            Caption         =   "2048 x 1536 x 24"
            Index           =   2
         End
         Begin VB.Menu mnuRes14x 
            Caption         =   "2048 x 1536 x 32"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuProcesses 
      Caption         =   "Processes"
      Begin VB.Menu mnuProcessesSetTrunc 
         Caption         =   "Set truncate..."
      End
   End
   Begin VB.Menu mnuTCPMonitor 
      Caption         =   "TCP Monitor"
      Begin VB.Menu mnuTCPMonitorInfo 
         Caption         =   "Adapter info..."
      End
      Begin VB.Menu mnuTCPMonitorGraph 
         Caption         =   "Show graph"
      End
      Begin VB.Menu mnuTCPMonitorCopy 
         Caption         =   "Copy to clipboard"
      End
      Begin VB.Menu mnuTCPMonitorStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTCPMonitorAll 
         Caption         =   "All adapters"
      End
      Begin VB.Menu mnuTCPMonitorIgnoreLoopback 
         Caption         =   "Ignore loopback adapter"
      End
      Begin VB.Menu mnuTCPMonitorStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTCPMonitorAdapter 
         Caption         =   "adapter 1"
         Index           =   1
      End
      Begin VB.Menu mnuTCPMonitorAdapter 
         Caption         =   "adapter 2"
         Index           =   2
      End
      Begin VB.Menu mnuTCPMonitorAdapter 
         Caption         =   "adapter 3"
         Index           =   3
      End
      Begin VB.Menu mnuTCPMonitorAdapter 
         Caption         =   "adapter 4"
         Index           =   4
      End
      Begin VB.Menu mnuTCPMonitorAdapter 
         Caption         =   "adapter 5"
         Index           =   5
      End
      Begin VB.Menu mnuTCPMonitorAdapter 
         Caption         =   "adapter 6"
         Index           =   6
      End
      Begin VB.Menu mnuTCPMonitorAdapter 
         Caption         =   "adapter 7"
         Index           =   7
      End
      Begin VB.Menu mnuTCPMonitorAdapter 
         Caption         =   "adapter 8"
         Index           =   8
      End
      Begin VB.Menu mnuTCPMonitorAdapter 
         Caption         =   "adapter 9"
         Index           =   9
      End
      Begin VB.Menu mnuTCPMonitorAdapter 
         Caption         =   "adapter 10"
         Index           =   10
      End
   End
   Begin VB.Menu mnuMSIE 
      Caption         =   "MSIE version"
      Begin VB.Menu mnuMSIEUpdate 
         Caption         =   "Update"
      End
      Begin VB.Menu mnuMSIECopy 
         Caption         =   "Copy to clipboard"
      End
   End
   Begin VB.Menu mnuDX 
      Caption         =   "DirectX"
      Begin VB.Menu mnuDXUpdate 
         Caption         =   "Update"
      End
      Begin VB.Menu mnuDXCopy 
         Caption         =   "Copy to clipboard"
      End
   End
   Begin VB.Menu mnuRAS 
      Caption         =   "RAS Connect"
      Begin VB.Menu mnuRASConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuRASDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuRASStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRASItem 
         Caption         =   "dummy"
         Index           =   0
      End
      Begin VB.Menu mnuRASItem 
         Caption         =   "dummy"
         Index           =   1
      End
      Begin VB.Menu mnuRASItem 
         Caption         =   "dummy"
         Index           =   2
      End
      Begin VB.Menu mnuRASItem 
         Caption         =   "dummy"
         Index           =   3
      End
      Begin VB.Menu mnuRASItem 
         Caption         =   "dummy"
         Index           =   4
      End
      Begin VB.Menu mnuRASItem 
         Caption         =   "dummy"
         Index           =   5
      End
      Begin VB.Menu mnuRASItem 
         Caption         =   "dummy"
         Index           =   6
      End
      Begin VB.Menu mnuRASItem 
         Caption         =   "dummy"
         Index           =   7
      End
      Begin VB.Menu mnuRASItem 
         Caption         =   "dummy"
         Index           =   8
      End
      Begin VB.Menu mnuRASItem 
         Caption         =   "dummy"
         Index           =   9
      End
   End
   Begin VB.Menu mnuNetstat 
      Caption         =   "Netstat"
      Begin VB.Menu mnuNetstatTruncate 
         Caption         =   "Set truncate..."
      End
      Begin VB.Menu mnuNetstatStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNetstatShowUDP 
         Caption         =   "Show UDP ports"
      End
      Begin VB.Menu mnuNetstatShowTimewait 
         Caption         =   "Show TIME_WAIT"
      End
      Begin VB.Menu mnuNetstatShowListening 
         Caption         =   "Show LISTENING"
      End
      Begin VB.Menu mnuNetstatStr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNetstatGetIPStats 
         Caption         =   "Get IP statistics..."
      End
      Begin VB.Menu mnuNetstatGetTCPStats 
         Caption         =   "Get TCP statistics..."
      End
      Begin VB.Menu mnuNetstatGetUDPStats 
         Caption         =   "Get UDP statistics..."
      End
      Begin VB.Menu mnuNetstatGetICMPStats 
         Caption         =   "Get ICMP statistics..."
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sLastServedTime$, sLastServedDate$

Public Sub GetMenuDots()
    Dim lMenu&, lSubMenu& ', lSub2Menu&
    Dim i%, MIO As MENUITEMINFO
    On Error GoTo Error:
    With MIO
        .cbSize = Len(MIO)
        .fMask = MIIM_TYPE
        .fType = MFT_RADIOCHECK
    End With
    
    lMenu = GetMenu(Me.hwnd)
    If lMenu = 0 Then
        MsgBox "Error setting dotted menu items: unable to get handle to main menu.", vbExclamation, "oops"
        Exit Sub
    End If
    
    'Modules and Settings menu
    lSubMenu = GetSubMenu(lMenu, 0)
    MIO.dwTypeData = "Modules" & vbTab & "Shift-Click"
    SetMenuItemInfo lSubMenu, 0, True, MIO
    MIO.dwTypeData = "Settings" & vbTab & "Ctrl-Click"
    SetMenuItemInfo lSubMenu, 1, True, MIO
    
    'Winamp controls
    lSubMenu = GetSubMenu(lMenu, 4)
    MIO.dwTypeData = "Start Winamp" & vbTab & "Click"
    SetMenuItemInfo lSubMenu, 0, True, MIO
    
    'Master & CD volume
    lSubMenu = GetSubMenu(lMenu, 12)
    MIO.dwTypeData = mnuMasterVolMute.Caption & vbTab & "Click"
    SetMenuItemInfo lSubMenu, 0, True, MIO
    lSubMenu = GetSubMenu(lMenu, 13)
    MIO.dwTypeData = mnuCDVolMute.Caption & vbTab & "Click"
    SetMenuItemInfo lSubMenu, 0, True, MIO
    
    'Disk free space
    lSubMenu = GetSubMenu(lMenu, 16)
    If lSubMenu = 0 Then
        MsgBox "Error setting dotted menu items: unable to get handle to Disk free space module menu.", vbExclamation, "oops"
        Exit Sub
    End If
    MIO.dwTypeData = mnuDiskFreeSpaceWakeUp.Caption & vbTab & "Click"
    SetMenuItemInfo lSubMenu, 5, True, MIO
    MIO.dwTypeData = mnuDiskFreeSpaceAll.Caption
    SetMenuItemInfo lSubMenu, 7, True, MIO
    For i = 0 To 25
        mnuDiskFreeSpaceA(i).Visible = True
        MIO.dwTypeData = mnuDiskFreeSpaceA(i).Caption & GetDiskLabel(mnuDiskFreeSpaceA(i).Caption)
        SetMenuItemInfo lSubMenu, i + 9, True, MIO
    Next i
    HideInvalidDisks
    
    'Free RAM
    lSubMenu = GetSubMenu(lMenu, 10)
    MIO.dwTypeData = mnuMemoryRAMDisplay(0).Caption
    SetMenuItemInfo lSubMenu, 0, True, MIO
    MIO.dwTypeData = mnuMemoryRAMDisplay(1).Caption
    SetMenuItemInfo lSubMenu, 1, True, MIO
    
    'Free pagefile
    lSubMenu = GetSubMenu(lMenu, 11)
    MIO.dwTypeData = mnuMemoryPageDisplay(0).Caption
    SetMenuItemInfo lSubMenu, 0, True, MIO
    MIO.dwTypeData = mnuMemoryPageDisplay(1).Caption
    SetMenuItemInfo lSubMenu, 1, True, MIO
    
    'CPU usage
    lSubMenu = GetSubMenu(lMenu, 15)
    MIO.dwTypeData = mnuCPUDisplay(0).Caption
    SetMenuItemInfo lSubMenu, 0, True, MIO
    MIO.dwTypeData = mnuCPUDisplay(1).Caption
    SetMenuItemInfo lSubMenu, 1, True, MIO
    
    'TCP Monitor
    lSubMenu = GetSubMenu(lMenu, 19)
    MIO.dwTypeData = mnuTCPMonitorAll.Caption
    SetMenuItemInfo lSubMenu, 4, True, MIO
    For i = 1 To 10
        MIO.dwTypeData = mnuTCPMonitorAdapter(i).Caption
        SetMenuItemInfo lSubMenu, 6 + i, True, MIO
    Next i
    
    'Winamp controls (hotkeys)
    lSubMenu = GetSubMenu(lMenu, 4)
    For i = 1 To 3
        MIO.dwTypeData = mnuWinampHotkeyMode(i).Caption
        SetMenuItemInfo lSubMenu, i + 6, True, MIO
    Next i
    
    'RAS Connections
    lSubMenu = GetSubMenu(lMenu, 22)
    MIO.dwTypeData = mnuRASConnect.Caption & vbTab & "Click"
    SetMenuItemInfo lSubMenu, 0, True, MIO
    MIO.dwTypeData = mnuRASDisconnect.Caption & vbTab & "Shift-Click"
    SetMenuItemInfo lSubMenu, 1, True, MIO
    For i = 0 To 9
        MIO.dwTypeData = mnuRASItem(i).Caption
        SetMenuItemInfo lSubMenu, i + 3, True, MIO
    Next i
    
    Exit Sub
Error:
    ShowError "Main", "frmMenu.GetMenuDots", Err.Number, Err.Description, False
End Sub

Public Sub HideInvalidDisks()
    Dim i%, sDummy$
    On Error GoTo Error:
    'Reduced this sub from 40 lines to 4.. damn I'm good :)
    
    sDummy = "," & Join(sDisks, ",") & ","
    For i = 0 To 25
        mnuDiskFreeSpaceA(i).Visible = IIf(InStr(sDummy, "," & Chr(i + 65) & ":\,") <> 0, True, False)
    Next i
    Exit Sub
    
Error:
    ShowError "Disk free space", "frmmenu.HideInvalidDisks", Err.Number, Err.Description, False
End Sub

Private Sub Form_Load()
    sLastServedTime = "never"
    sLastServedDate = "never"
End Sub

Private Sub mnuCDVolMute_Click()
    mnuCDVolMute.Checked = Not mnuCDVolMute.Checked
End Sub

Private Sub mnuCPUBar_Click()
    mnuCPUBar.Checked = Not mnuCPUBar.Checked
    frmMain.shpCPUBack.Visible = mnuCPUBar.Checked
    frmMain.shpCPUFore.Visible = mnuCPUBar.Checked
    frmMain.TriggerTimers MODULE_CPUUSAGE
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "CPUShowBar", Abs(CInt(mnuCPUBar.Checked))
End Sub

Private Sub mnuCPUCopy_Click()
    CopyToClipboard frmMain.lblCPU.Caption, 2
End Sub

Private Sub mnuCPUDisplay_Click(Index As Integer)
    mnuCPUDisplay(Abs(Index - 1)).Checked = False
    mnuCPUDisplay(Index).Checked = True
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "CPUDisplay", CLng(Index)
    frmMain.TriggerTimers MODULE_CPUUSAGE
End Sub

Private Sub mnuCPUGraph_Click()
    mnuCPUGraph.Checked = Not mnuCPUGraph.Checked
    frmMain.picGraphCPU.Visible = mnuCPUGraph.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "CPUGraph", Abs(CInt(mnuCPUGraph.Checked))
    If mnuCPUGraph.Checked Then
        frmMain.fraCPU.Width = frmMain.picGraphCPU.Left + frmMain.picGraphCPU.Width
    Else
        frmMain.fraCPU.Width = frmMain.picGraphCPU.Left
    End If
    AlignModules bBarDocking
    frmMain.TriggerTimers MODULE_CPUUSAGE
End Sub

Private Sub mnuCPUInterval_Click()
    Dim iDummy%
    iDummy = GetInterval("CPU usage", 1, frmMain.timCPU.Interval / 1000, 1, 30)
    If iDummy = -1 Then Exit Sub
    frmMain.timCPU.Interval = CLng(iDummy) * 1000
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "CPUInterval", CLng(iDummy) * 1000
End Sub

Private Sub mnuDateCopy_Click()
    CopyToClipboard frmMain.lblDate.Caption, 3
End Sub

Private Sub mnuDiskFreeSpaceA_Click(Index As Integer)
    mnuDiskFreeSpaceAll.Checked = False
    Dim i%
    For i = 0 To 25
        mnuDiskFreeSpaceA(i).Checked = False
    Next i
    mnuDiskFreeSpaceA(Index).Checked = True
    sCurrentDisk = Chr(65 + Index)
    RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "DiskFreeDisk", sCurrentDisk
    frmMain.TriggerTimers MODULE_DISKFREESPACE
End Sub

Private Sub mnuDiskFreeSpaceAll_Click()
    Dim i%
    For i = 0 To 25
        mnuDiskFreeSpaceA(i).Checked = False
    Next i
    mnuDiskFreeSpaceAll.Checked = True
    sCurrentDisk = "."
    RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "DiskFreeDisk", sCurrentDisk
    frmMain.TriggerTimers MODULE_DISKFREESPACE
End Sub

Private Sub mnuDiskFreeSpaceBar_Click()
    mnuDiskFreeSpaceBar.Checked = Not mnuDiskFreeSpaceBar.Checked
    frmMain.shpDiskFreeBack.Visible = mnuDiskFreeSpaceBar.Checked
    frmMain.shpDiskFreeFore.Visible = mnuDiskFreeSpaceBar.Checked
    frmMain.TriggerTimers MODULE_DISKFREESPACE
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "DiskFreeShowBar", Abs(CInt(mnuDiskFreeSpaceBar.Checked))
End Sub

Private Sub mnuDiskFreeSpaceCopy_Click()
    CopyToClipboard frmMain.lblDiskFreeSpace.Caption, 4
End Sub

Private Sub mnuDiskFreeSpaceInterval_Click()
    Dim iDummy%
    iDummy = GetInterval("Disk free space", 5, frmMain.timDiskFreeSpace.Interval / 1000, 1, 60)
    If iDummy = -1 Then Exit Sub
    frmMain.timDiskFreeSpace.Interval = CLng(iDummy) * 1000
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "DiskFreeSpaceInterval", CLng(iDummy) * 1000
End Sub

Private Sub mnuDiskFreeSpaceQuota_Click()
    mnuDiskFreeSpaceQuota.Checked = Not mnuDiskFreeSpaceQuota.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "DiskFreeQuotas", Abs(CInt(mnuDiskFreeSpaceQuota.Checked))
    frmMain.TriggerTimers MODULE_DISKFREESPACE
End Sub

Private Sub mnuDiskFreeSpaceWakeUp_Click()
    frmMain.imgDiskFreeSpace_MouseUp 1, 0, 0, 0
End Sub

Private Sub mnuDXCopy_Click()
    CopyToClipboard frmMain.lblDX.Caption, 22
End Sub

Private Sub mnuDXUpdate_Click()
    frmMain.GetDXVersion
End Sub

Private Sub mnuExitWinButtons_Click(Index As Integer)
    Dim i%, j&, sBlah$
    On Error GoTo Error:
    mnuExitWinButtons(Index).Checked = Not mnuExitWinButtons(Index).Checked
    j = 0
    sBlah = "     "
    For i = 0 To 4
        If Not mnuExitWinButtons(i).Checked Then
            frmMain.imgExitWin(i).Visible = False
            Mid(sBlah, i + 1, 1) = "0"
        Else
            frmMain.imgExitWin(i).Visible = True
            frmMain.imgExitWin(i).Left = j
            Mid(sBlah, i + 1, 1) = "1"
            j = j + 360
        End If
    Next i
    If sBlah = "00000" Then
        sBlah = "11111"
        j = 0
        For i = 0 To 4
            mnuExitWinButtons(i).Checked = True
            frmMain.imgExitWin(i).Visible = True
            frmMain.imgExitWin(i).Left = j
            j = j + 360
        Next i
        frmMain.fraExitWin.Width = j - 120
    End If
    RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "ExitWinButtons", sBlah
    frmMain.fraExitWin.Width = j - 120
    AlignModules bBarDocking
    Exit Sub
    
Error:
    ShowError "Exit Windows", "frmmenu.mnuExitWinButtons(" & CStr(Index) & ")_Click", Err.Number, Err.Description, False
End Sub

Private Sub mnuExitWinConfirm_Click()
    mnuExitWinConfirm.Checked = Not mnuExitWinConfirm.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ExitWinConfirm", Abs(CInt(mnuExitWinConfirm.Checked))
End Sub

Private Sub mnuExitWinForce_Click()
    mnuExitWinForce.Checked = Not mnuExitWinForce.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ExitWinForce", Abs(CInt(mnuExitWinForce.Checked))
End Sub

Private Sub mnuIPsCopy_Click()
    CopyToClipboard frmMain.lblIPs.Caption, 8
End Sub

Private Sub mnuIPsIgnoreLocalhost_Click()
    mnuIPsIgnoreLocalhost.Checked = Not mnuIPsIgnoreLocalhost.Checked
    bIgnoreLocalHostIP = mnuIPsIgnoreLocalhost.Checked
    frmMain.TriggerTimers MODULE_IPS
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "IPsIgnoreLocalhost", Abs(CInt(mnuIPsIgnoreLocalhost.Checked))
End Sub

Public Sub mnuLockSetPass_Click()
    Dim sPass$, sOldPass$, sNewPass$
    On Error GoTo Error:
    sPass = ROT13(RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "LockPassword"))

    If sPass <> "" Then
        'Old password present, needed for verification
        sPass = GetPassword("Please enter your old and new password to change the Lock Screen password:", 0, sPass)
        If sPass <> "0" Then RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "LockPassword", ROT13(sPass)
    Else
        'No old password, either first time or password was deleted
        sPass = GetPassword("Please enter your desired password for the Lock Screen module:", 2)
        If sPass <> "0" Then RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "LockPassword", ROT13(sPass)
    End If
    Exit Sub
    
Error:
    ShowError "Lock screen", "frmmenu.mnuLockSetPass_Click", Err.Number, Err.Description, False
End Sub

Private Sub mnuMainAbout_Click()
    Dim sMsg$, sVersion$
    sVersion = "Uptimer4 - v" & App.Major & "." & App.Minor & "." & App.Revision
    sMsg = "Merijn, merijn@spywareinfo.com" & vbCrLf & "Soeperman Enterprises Ltd."
    ShellAbout Me.hwnd, sVersion & "#Windows powered", sMsg, frmMain.Icon
End Sub

Private Sub mnuMainDock_Click(Index As Integer)
    If Index = 0 Then
        'Dock window
        MakeAppBar True, bBarPos
        mnuMainDock(0).Visible = False
        mnuMainDock(1).Visible = True
    Else
        'Undock window (floating)
        MakeAppBar False, True
        mnuMainDock(0).Visible = True
        mnuMainDock(1).Visible = False
    End If
End Sub

Private Sub mnuMainExit_Click()
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "LastAppBar"
    Unload frmMain
    End
End Sub

Private Sub mnuMainHelp_Click()
    Dim sExePath$
    sExePath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
    If Dir(sExePath & "Uptimer4.hlp") = "" Then
        If MsgBox("Help file not found. Download a complete Uptimer4 install " & _
               "at http://www.geocities.com/merijn_bellekom/new/uptimer.html." & _
               vbCrLf & "Copy this address to clipboard?", vbExclamation + vbYesNo, "not found") = vbYes Then
            Clipboard.Clear
            Clipboard.SetText "http://www.geocities.com/merijn_bellekom/new/uptimer.html"
        End If
    Else
        ShellExecute Me.hwnd, "open", sExePath & "Uptimer4.hlp", "", "", 1
    End If
End Sub

Private Sub mnuMainHideIcon_Click()
    MakeTrayIcon False, 0
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "TrayIcon", 0
End Sub

Private Sub mnuMainModules_Click()
    If Not bBarDocking Then SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    frmModules.Show 1
    If bAlwaysOnTop Then SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub mnuMainReload_Click()
    frmMain.GetCustomIcons
End Sub

Private Sub mnuMainReset_Click()
    If Not bBarDocking Then Exit Sub
    Dim APD As APPBARDATA
    APD.cbSize = Len(APD)
    APD.hwnd = frmMain.hwnd
    SHAppBarMessage 1, APD
    DoEvents
    MakeAppBar True, bBarPos
    DoEvents
    SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub mnuMainSettings_Click()
    If Not bBarDocking Then SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Load frmSettings
    frmSettings.Show 1
    If bAlwaysOnTop Then SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub mnuMainUpdate_Click()
    frmMain.TriggerTimers
End Sub

Private Sub mnuMasterVolMute_Click()
    mnuMasterVolMute.Checked = Not mnuMasterVolMute.Checked
End Sub

Private Sub mnuMemoryPageBar_Click()
    mnuMemoryPageBar.Checked = Not mnuMemoryPageBar.Checked
    frmMain.shpMemoryPageBack.Visible = mnuMemoryPageBar.Checked
    frmMain.shpMemoryPageFore.Visible = mnuMemoryPageBar.Checked
    frmMain.TriggerTimers MODULE_FREEPAGEFILE
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MemoryPageShowBar", Abs(CInt(mnuMemoryPageBar.Checked))
End Sub

Private Sub mnuMemoryPageCopy_Click()
    CopyToClipboard frmMain.lblMemoryPage.Caption, 6
End Sub

Private Sub mnuMemoryPageDisplay_Click(Index As Integer)
    mnuMemoryPageDisplay(Abs(Index - 1)).Checked = False
    mnuMemoryPageDisplay(Index).Checked = True
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MemoryPageDisplay", CLng(Index)
    frmMain.TriggerTimers MODULE_FREEPAGEFILE
End Sub

Private Sub mnuMemoryPageGraph_Click()
    mnuMemoryPageGraph.Checked = Not mnuMemoryPageGraph.Checked
    frmMain.picGraphPage.Visible = mnuMemoryPageGraph.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MemoryPageGraph", Abs(CInt(mnuMemoryPageGraph.Checked))
    If mnuMemoryPageGraph.Checked Then
        frmMain.fraMemoryPage.Width = frmMain.picGraphPage.Left + frmMain.picGraphPage.Width
    Else
        frmMain.fraMemoryPage.Width = frmMain.picGraphPage.Left
    End If
    AlignModules bBarDocking
    frmMain.TriggerTimers MODULE_FREEPAGEFILE
End Sub

Private Sub mnuMemoryPageInterval_Click()
    Dim iDummy%
    iDummy = GetInterval("Free pagefile", 5, frmMain.timMemoryPage.Interval / 1000, 1, 30)
    If iDummy = -1 Then Exit Sub
    frmMain.timMemoryPage.Interval = CLng(iDummy) * 1000
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MemoryPageInterval", CLng(iDummy) * 1000
End Sub

Private Sub mnuMemoryRAMBar_Click()
    mnuMemoryRAMBar.Checked = Not mnuMemoryRAMBar.Checked
    frmMain.shpMemoryRAMBack.Visible = mnuMemoryRAMBar.Checked
    frmMain.shpMemoryRAMFore.Visible = mnuMemoryRAMBar.Checked
    frmMain.TriggerTimers MODULE_FREERAM
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MemoryRAMShowBar", Abs(CInt(mnuMemoryRAMBar.Checked))
End Sub

Private Sub mnuMemoryRAMCopy_Click()
    CopyToClipboard frmMain.lblMemoryRAM.Caption, 7
End Sub

Private Sub mnuMemoryRAMDisplay_Click(Index As Integer)
    mnuMemoryRAMDisplay(Abs(Index - 1)).Checked = False
    mnuMemoryRAMDisplay(Index).Checked = True
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MemoryRAMDisplay", CLng(Index)
    frmMain.TriggerTimers MODULE_FREERAM
End Sub

Private Sub mnuMemoryRamFree_Click(Index As Integer)
    'This doesn't need a API or anything, though it
    'may be possible to use MemAlloc() instead... maybe
    'NB: Blatant code copy from Uptimer3
    
    Dim iToFree As Integer, sDummy As String * 1000
    Dim sBigArray() As String, i As Integer, j As Integer
    On Error GoTo Error:
    'Get amount of MB to free "Free up 14 MB"
    iToFree = Val(Mid(mnuMemoryRamFree(Index).Caption, 9))
    
    With frmMain
        .timMemoryRAM.Enabled = False
        .lblMemoryRAM.Caption = "working..."
        .shpMemoryRAMBack.Visible = True
        .shpMemoryRAMFore.Visible = True
        .shpMemoryRAMFore.Width = 0
    End With
    DoEvents
    
    'Fill array with data
    ReDim sBigArray(iToFree * 1000)
    For i = 1 To iToFree * 1000
        'Make junk a bit random, or it won't work
        For j = 1 To 20
            sDummy = sDummy & String(50, Chr(250 * Rnd))
        Next j
        If i Mod 100 = 0 Then
            frmMain.shpMemoryRAMFore.Width = frmMain.shpMemoryRAMBack.Width * (i / (iToFree * 1000))
            DoEvents
        End If
        sBigArray(i) = sDummy
    Next i
    DoEvents
    'Empty array, thus freeing up memory
    ReDim sBigArray(0)
    sDummy = ""
    With frmMain
        .timMemoryRAM.Enabled = True
        If Not mnuMemoryRAMBar.Checked Then
            .shpMemoryRAMBack.Visible = False
            .shpMemoryRAMFore.Visible = False
        End If
        .TriggerTimers MODULE_FREERAM
    End With
    Exit Sub
    
Error:
    ShowError "Free RAM", "frmmenu.mnuMemoryRAMDisplay_Click", Err.Number, Err.Description, False
End Sub

Private Sub mnuMemoryRAMGraph_Click()
    mnuMemoryRAMGraph.Checked = Not mnuMemoryRAMGraph.Checked
    frmMain.picGraphRAM.Visible = mnuMemoryRAMGraph.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MemoryRAMGraph", Abs(CInt(mnuMemoryRAMGraph.Checked))
    If mnuMemoryRAMGraph.Checked Then
        frmMain.fraMemoryRAM.Width = frmMain.picGraphRAM.Left + frmMain.picGraphRAM.Width
    Else
        frmMain.fraMemoryRAM.Width = frmMain.picGraphRAM.Left
    End If
    AlignModules bBarDocking
    frmMain.TriggerTimers MODULE_FREERAM
End Sub

Private Sub mnuMemoryRAMInterval_Click()
    Dim iDummy%
    iDummy = GetInterval("Free RAM", 2, frmMain.timMemoryRAM.Interval / 1000, 1, 30)
    If iDummy = -1 Then Exit Sub
    frmMain.timMemoryRAM.Interval = CLng(iDummy) * 1000
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MemoryRAMInterval", CLng(iDummy) * 1000
End Sub

Private Sub mnuMemoryRamSet_Click()
    Dim sAmounts$, sDummy$, sMsg$
    On Error GoTo Error:
GetAmount1:
    sMsg = "Enter the amount of memory (in MB) to be freed up with the FIRST menu choice:"
    sDummy = InputBox(sMsg, "Free up memory: Set amounts (1)", "1")
    If sDummy = "" Then Exit Sub
    If Val(sDummy) < 1 Then
        If MsgBox("Invalid amount. Try again?", vbYesNo + vbExclamation, "oops") = vbYes Then GoTo GetAmount1
    End If
    frmMenu.mnuMemoryRamFree(0).Caption = "Free up " & sDummy & " MB"
    sAmounts = sDummy & ","
    sMsg = Replace(sMsg, "FIRST", "SECOND")
GetAmount2:
    sDummy = InputBox(sMsg, "Free up memory: Set amounts (2)", "5")
    If sDummy = "" Then Exit Sub
    If Val(sDummy) < 1 Then
        If MsgBox("Invalid amount. Try again?", vbYesNo + vbExclamation, "oops") = vbYes Then GoTo GetAmount2
    End If
    frmMenu.mnuMemoryRamFree(1).Caption = "Free up " & sDummy & " MB"
    sAmounts = sAmounts & sDummy & ","
    sMsg = Replace(sMsg, "SECOND", "THIRD")
GetAmount3:
    sDummy = InputBox(sMsg, "Free up memory: Set amounts (3)", "10")
    If sDummy = "" Then Exit Sub
    If Val(sDummy) < 1 Then
        If MsgBox("Invalid amount. Try again?", vbYesNo + vbExclamation, "oops") = vbYes Then GoTo GetAmount3
    End If
    frmMenu.mnuMemoryRamFree(2).Caption = "Free up " & sDummy & " MB"
    sAmounts = sAmounts & sDummy & ","
    sMsg = Replace(sMsg, "THIRD", "FOURTH")
GetAmount4:
    sDummy = InputBox(sMsg, "Free up memory: Set amounts (4)", "20")
    If sDummy = "" Then Exit Sub
    If Val(sDummy) < 1 Then
        If MsgBox("Invalid amount. Try again?", vbYesNo + vbExclamation, "oops") = vbYes Then GoTo GetAmount4
    End If
    frmMenu.mnuMemoryRamFree(3).Caption = "Free up " & sDummy & " MB"
    sAmounts = sAmounts & sDummy
    RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "MemoryRAMFreeMemory", sAmounts
    Exit Sub
    
Error:
    ShowError "Free RAM", "frmmenu.mnuMemoryRamSet_Click", Err.Number, Err.Description, False
End Sub

Private Sub mnuMouseIdleSetTimeout_Click()
    Dim sMsg$
    sMsg = "Enter the idle timeout you want in seconds. "
    sMsg = sMsg & "If the mouse is idle longer than this, "
    sMsg = sMsg & "the Mouse idle time will start flashing."
    sMsg = InputBox(sMsg, "Enter mouse idle timeout", CStr(iMouseTimeout))
    If sMsg = "" Then Exit Sub
    iMouseTimeout = Val(sMsg)
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MouseIdleTimeout", CLng(iMouseTimeout)
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MouseIdleUseTimeout", 1
End Sub

Private Sub mnuMouseIdleUseTimeout_Click()
    mnuMouseIdleUseTimeout.Checked = Not mnuMouseIdleUseTimeout.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MouseIdleUseTimeout", Abs(CInt(mnuMouseIdleUseTimeout.Checked))
    mnuMouseIdleSetTimeout.Enabled = mnuMouseIdleUseTimeout.Checked
    If Not mnuMouseIdleUseTimeout.Checked Then iMouseTimeout = 0
End Sub

Private Sub mnuMSIECopy_Click()
    CopyToClipboard frmMain.lblMSIE.Caption, 21
End Sub

Private Sub mnuMSIEUpdate_Click()
    frmMain.GetMSIEVersion
End Sub

Private Sub mnuNetstatGetICMPStats_Click()
    Dim uICMPStats As MIBICMPINFO, sMsg$
    
    If GetIcmpStatistics(uICMPStats) <> 0 Then
        MsgBox "Unable to retrieve ICMP statistics: GetIcmpStatistics() API returned an error.", vbExclamation, "oops"
        Exit Sub
    End If
        
    sMsg = " === Inbound ICMP (received) ===" & vbCrLf
    With uICMPStats.icmpInStats
        sMsg = sMsg & "Number of messages = " & .dwMsgs & vbCrLf
        sMsg = sMsg & "Number of errors = " & .dwErrors & vbCrLf
        sMsg = sMsg & "Destination unreachable messages = " & .dwDestUnreachs & vbCrLf
        sMsg = sMsg & "TTL exceeded messages = " & .dwTimeExcds & vbCrLf
        sMsg = sMsg & "Parameter problems messages = " & .dwParmProbs & vbCrLf
        sMsg = sMsg & "Source quench messages = " & .dwSrcQuenchs & vbCrLf
        sMsg = sMsg & "Redirect messages = " & .dwRedirects & vbCrLf
        sMsg = sMsg & "Echo request messages (ping) = " & .dwEchos & vbCrLf
        sMsg = sMsg & "Echo reply messages (ping reply) = " & .dwEchoReps & vbCrLf
        sMsg = sMsg & "Timestamp request messages = " & .dwTimestamps & vbCrLf
        sMsg = sMsg & "Timestamp reply messages = " & .dwTimestampReps & vbCrLf
        sMsg = sMsg & "Address mask request messages = " & .dwAddrMasks & vbCrLf
        sMsg = sMsg & "Address mask reply messages = " & .dwAddrMaskReps & vbCrLf
    End With
    
    sMsg = sMsg & vbCrLf & " === Outbound ICMP (sent) ===" & vbCrLf
    With uICMPStats.icmpOutStats
        sMsg = sMsg & "Number of messages = " & .dwMsgs & vbCrLf
        sMsg = sMsg & "Number of errors = " & .dwErrors & vbCrLf
        sMsg = sMsg & "Destination unreachable messages = " & .dwDestUnreachs & vbCrLf
        sMsg = sMsg & "TTL exceeded messages = " & .dwTimeExcds & vbCrLf
        sMsg = sMsg & "Parameter problems messages = " & .dwParmProbs & vbCrLf
        sMsg = sMsg & "Source quench messages = " & .dwSrcQuenchs & vbCrLf
        sMsg = sMsg & "Redirect messages = " & .dwRedirects & vbCrLf
        sMsg = sMsg & "Echo request messages (ping) = " & .dwEchos & vbCrLf
        sMsg = sMsg & "Echo reply messages (ping reply) = " & .dwEchoReps & vbCrLf
        sMsg = sMsg & "Timestamp request messages = " & .dwTimestamps & vbCrLf
        sMsg = sMsg & "Timestamp reply messages = " & .dwTimestampReps & vbCrLf
        sMsg = sMsg & "Address mask request messages = " & .dwAddrMasks & vbCrLf
        sMsg = sMsg & "Address mask reply messages = " & .dwAddrMaskReps & vbCrLf
    End With
    
    MsgBox sMsg, vbInformation, "ICMP Statatistics"
End Sub

Private Sub mnuNetstatGetIPStats_Click()
    Dim uIPStats As MIB_IPSTATS, sMsg$
    
    If GetIpStatistics(uIPStats) <> 0 Then
        MsgBox "Unable to retrieve IP statistics: GetIpStatistics() API returned an error.", vbExclamation, "oops"
        Exit Sub
    End If
    
    sMsg = ""
    With uIPStats
        sMsg = sMsg & "IP forwarding enabled = " & .dwForwarding & vbCrLf
        sMsg = sMsg & "Default TTL = " & .dwDefaultTTL & vbCrLf
        sMsg = sMsg & "Datagrams received = " & .dwInReceives & vbCrLf
        sMsg = sMsg & "Received header errors = " & .dwInHdrErrors & vbCrLf
        sMsg = sMsg & "Received address errors = " & .dwInAddrErrors & vbCrLf
        sMsg = sMsg & "Datagrams forwarded = " & .dwForwDatagrams & vbCrLf
        sMsg = sMsg & "Datagrams of unknown protocols received = " & .dwInUnknownProtos & vbCrLf
        sMsg = sMsg & "Received datagrams delivered = " & .dwInDelivers & vbCrLf
        sMsg = sMsg & "Received datagrams discarded = " & .dwInDiscards & vbCrLf
        sMsg = sMsg & "Outgoing datagrams requested = " & .dwOutRequests & vbCrLf
        sMsg = sMsg & "Outgoing datagrams discarded = " & .dwOutDiscards & vbCrLf
        sMsg = sMsg & "Sent datagrams discarded = " & .dwRoutingDiscards & vbCrLf
        sMsg = sMsg & "Datagrams for which no route = " & .dwOutNoRoutes & vbCrLf
        sMsg = sMsg & "Datagrams for which all frags didn't arrive = " & .dwReasmTimeout & vbCrLf
        sMsg = sMsg & "Datagrams requiring reassembly = " & .dwReasmReqds & vbCrLf
        sMsg = sMsg & "Datagrams successfully reassembled = " & .dwReasmOks & vbCrLf
        sMsg = sMsg & "Datagrams failed reassembly = " & .dwReasmFails & vbCrLf
        sMsg = sMsg & "Datagrams fragmented = " & .dwFragCreates & vbCrLf
        sMsg = sMsg & "Successful datagram fragmentations = " & .dwFragOks & vbCrLf
        sMsg = sMsg & "Failed datagram fragmentations = " & .dwFragFails & vbCrLf
        sMsg = sMsg & "Number of interfaces on computer = " & .dwNumIf & vbCrLf
        sMsg = sMsg & "Number of IP address on computer = " & .dwNumAddr & vbCrLf
        sMsg = sMsg & "Number of routes in routing table = " & .dwNumRoutes & vbCrLf
    End With
    
    MsgBox sMsg, vbInformation, "IP Statistics"
End Sub

Private Sub mnuNetstatGetTCPStats_Click()
    Dim uTCPStats As MIB_TCPSTATS, sMsg$
    
    If GetTcpStatistics(uTCPStats) <> 0 Then
        MsgBox "Unable to retrieve TCP statistics: GetTcpStatistics() API returned an error.", vbExclamation, "oops"
        Exit Sub
    End If
    
    sMsg = ""
    With uTCPStats
        sMsg = sMsg & "Timeout algorithm = " & .dwRtoAlgorithm & vbCrLf
        sMsg = sMsg & "Minimum timeout = " & .dwRtoMin & vbCrLf
        sMsg = sMsg & "Maximum timeout = " & .dwRtoMax & vbCrLf
        sMsg = sMsg & "Maximum connections = " & .dwMaxConn & vbCrLf
        sMsg = sMsg & "Active opens = " & .dwActiveOpens & vbCrLf
        sMsg = sMsg & "Passive opens = " & .dwPassiveOpens & vbCrLf
        sMsg = sMsg & "Failed attempts = " & .dwAttemptFails & vbCrLf
        sMsg = sMsg & "Established connections = " & .dwCurrEstab & vbCrLf
        sMsg = sMsg & "Established connections reset = " & .dwEstabResets & vbCrLf
        sMsg = sMsg & "Segments received = " & .dwInSegs & vbCrLf
        sMsg = sMsg & "Segments transmitted = " & .dwOutSegs & vbCrLf
        sMsg = sMsg & "Segments retransmitted = " & .dwRetransSegs & vbCrLf
        sMsg = sMsg & "Incoming errors = " & .dwInErrs & vbCrLf
        sMsg = sMsg & "Outgoing resets = " & .dwOutRsts & vbCrLf
        sMsg = sMsg & "Cumulative connections = " & .dwNumConns & vbCrLf
    End With
    
    MsgBox sMsg, vbInformation, "TCP Statistics"
End Sub

Private Sub mnuNetstatGetUDPStats_Click()
    Dim uUDPStats As MIB_UDPSTATS, sMsg$
    
    If GetUdpStatistics(uUDPStats) <> 0 Then
        MsgBox "Unable to retrieve UDP statistics: GetUdpStatistics() API returned an error.", vbExclamation, "oops"
        Exit Sub
    End If
    
    sMsg = ""
    With uUDPStats
        sMsg = sMsg & "Received datagrams = " & .dwInDatagrams & vbCrLf
        sMsg = sMsg & "Received datagrams for which no port = " & .dwNoPorts & vbCrLf
        sMsg = sMsg & "Received datagram errors = " & .dwInErrors & vbCrLf
        sMsg = sMsg & "Sent datagrams = " & .dwOutDatagrams & vbCrLf
        sMsg = sMsg & "Number of entries in listening table = " & .dwNumAddrs & vbCrLf
    End With
    
    MsgBox sMsg, vbInformation, "UDP Statistics"
End Sub

Private Sub mnuNetstatShowListening_Click()
    mnuNetstatShowListening.Checked = Not mnuNetstatShowListening.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "NetstatShowListening", Abs(CInt(mnuNetstatShowListening.Checked))
    frmMain.TriggerTimers MODULE_NETSTAT
End Sub

Private Sub mnuNetstatShowTimewait_Click()
    mnuNetstatShowTimewait.Checked = Not mnuNetstatShowTimewait.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "NetstatShowTimewait", Abs(CInt(mnuNetstatShowTimewait.Checked))
    frmMain.TriggerTimers MODULE_NETSTAT
End Sub

Private Sub mnuNetstatShowUDP_Click()
    mnuNetstatShowUDP.Checked = Not mnuNetstatShowUDP.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "NetstatShowUDP", Abs(CInt(mnuNetstatShowUDP.Checked))
    frmMain.TriggerTimers MODULE_NETSTAT
End Sub

Private Sub mnuNetstatTruncate_Click()
    Dim sMsg$
GetTrunc:
    sMsg = "When the list of netstat items is too long, "
    sMsg = sMsg & "the bubble tooltip can be too large to fit "
    sMsg = sMsg & "on screen and won't display at all." & vbCrLf
    sMsg = sMsg & "This can be prevented by cutting off the list "
    sMsg = sMsg & "at a certain number of items." & vbCrLf
    sMsg = sMsg & "Enter maximum number of listed items (0 to disable):"
    sMsg = InputBox(sMsg, "Set netstat truncate", iNetstatTrunc)
    If sMsg = "" Then Exit Sub
    If IsNumeric(sMsg) Then
        If Val(sMsg) = 0 Then
            iNetstatTrunc = 0
            RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "NetstatTruncate"
        Else
            iNetstatTrunc = Val(sMsg)
            RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "NetstatTruncate", CLng(Val(sMsg))
        End If
        frmMain.TriggerTimers MODULE_NETSTAT
    Else
        If MsgBox("Invalid number. Try again?", vbYesNo + vbExclamation, "yech") = vbYes Then GoTo GetTrunc
    End If
End Sub

Private Sub mnuOSCopy_Click()
    CopyToClipboard frmMain.lblOS.Caption, 17
End Sub

Private Sub mnuOSFriendly_Click()
    mnuOSFriendly.Checked = Not mnuOSFriendly.Checked
    GetWinVersion mnuOSFriendly.Checked, mnuOSShowBuildSP(1).Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "OSFriendly", Abs(CInt(mnuOSFriendly.Checked))
End Sub

Private Sub mnuOSShowBuildSP_Click(Index As Integer)
    If bIsWinNT Then Exit Sub
    'Index = 0 -> show build
    'Index = 1 -> show service pack (if NT)
    mnuOSShowBuildSP(Abs(Index - 1)).Checked = False
    mnuOSShowBuildSP(Index).Checked = True
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "OSShowSP", CLng(Index)
    If RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "OSFriendly", 0) = 0 Then
        GetWinVersion False, CBool(mnuOSShowBuildSP(Index).Checked)
    Else
        GetWinVersion True, CBool(mnuOSShowBuildSP(Index).Checked)
    End If
End Sub

Private Sub mnuPowerBar_Click()
    mnuPowerBar.Checked = Not mnuPowerBar.Checked
    frmMain.shpPowerBack.Visible = mnuPowerBar.Checked
    frmMain.shpPowerFore.Visible = mnuPowerBar.Checked
    frmMain.TriggerTimers MODULE_POWER
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "PowerShowBar", Abs(CInt(mnuPowerBar.Checked))
End Sub

Private Sub mnuPowerBatChargeOnAC_Click()
    mnuPowerBatChargeOnAC.Checked = Not mnuPowerBatChargeOnAC.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "PowerBattChargeOnAC", Abs(CInt(mnuPowerBatChargeOnAC.Checked))
    frmMain.TriggerTimers MODULE_POWER
End Sub

Private Sub mnuPowerCopy_Click()
    CopyToClipboard frmMain.lblPower.Caption, 11
End Sub

Private Sub mnuPowerInterval_Click()
    Dim iDummy%
    iDummy = GetInterval("Power status", 10, frmMain.timPower.Interval / 1000, 1, 60)
    If iDummy = -1 Then Exit Sub
    frmMain.timPower.Interval = CLng(iDummy) * 1000
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "PowerInterval", CLng(iDummy) * 1000
End Sub

Private Sub mnuProcessesSetTrunc_Click()
    Dim sMsg$
GetTrunc:
    sMsg = "When the list of running processes is too long, "
    sMsg = sMsg & "the bubble tooltip can be too large to fit "
    sMsg = sMsg & "on screen and won't display at all." & vbCrLf
    sMsg = sMsg & "This can be prevented by cutting off the list "
    sMsg = sMsg & "at a certain number of processes." & vbCrLf
    sMsg = sMsg & "Enter maximum number of listed tasks (0 to disable):"
    sMsg = InputBox(sMsg, "Set process list truncate", iProcessTrunc)
    If sMsg = "" Then Exit Sub
    If IsNumeric(sMsg) Then
        If Val(sMsg) = 0 Then
            iProcessTrunc = 0
            RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "ProcessesTruncate"
        Else
            iProcessTrunc = Val(sMsg)
            RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ProcessesTruncate", CLng(Val(sMsg))
        End If
        frmMain.TriggerTimers MODULE_PROCESSES
    Else
        If MsgBox("Invalid number. Try again?", vbYesNo + vbExclamation, "yech") = vbYes Then GoTo GetTrunc
    End If
End Sub

Public Sub mnuRASConnect_Click()
    Dim sConnName$, i%, uBytes() As Byte, lRet&
    Dim uRDP As RASDIALPARAMS, hRasConn&
    uRDP.dwSize = LenB(uRDP) + 4 '1052
    
    'Get name of connection to dial
    For i = 0 To 9
        If mnuRASItem(i).Checked Then
            sConnName = mnuRASItem(i).Caption
            Exit For
        End If
    Next i
    If sConnName = "" Or sConnName = "(empty)" Then
        MsgBox "No connections on DUN system!", vbExclamation, "oops"
        Exit Sub
    End If
    
    'Get username/password for connection
    StringToBytes sConnName, uRDP.szEntryName
    lRet = RasGetEntryDialParams(sConnName, uRDP, 0)
    If lRet <> 0 Then
        MsgBox "Unable to get username/password for '" & sConnName & "' connection! RasGetEntryDialParams: " & CStr(lRet) & ".", vbCritical, "error"
        Exit Sub
    End If
    
    'Fill up rest of structure with stuff
    StringToBytes sConnName, uRDP.szEntryName
    StringToBytes "*", uRDP.szCallbackNumber
    StringToBytes "*", uRDP.szDomain
    If uRDP.szPassword(0) = 0 Then
        Dim sPass$
        sPass = GetPassword("Please enter the password for the " & sConnName & " connection:", 2)
        If sPass = "0" Then Exit Sub
        StringToBytes sPass, uRDP.szPassword
    End If
    
    'Dial connection
    lRet = RasDial(ByVal 0, "", uRDP, -1, ByVal Me.hwnd, hRasConn)
    If lRet <> 0 Then
        MsgBox "Unable to dial '" & sConnName & "' connection! RasDial: " & CStr(lRet) & ".", vbCritical, "error"
        RasHangUp hRasConn
    End If
    frmMain.TriggerTimers MODULE_RAS
End Sub

Public Sub mnuRASDisconnect_Click()
    'Enumerate connections and terminate first one
    Dim uRasConn(255) As RASCONN, i%, lEntries&, j$, lRet&
    uRasConn(0).dwSize = LenB(uRasConn(0)) '412
    lRet = RasEnumConnections(uRasConn(0), 256 * uRasConn(0).dwSize, lEntries)
    If lRet = 0 Then
        If lEntries = 0 Then
            MsgBox "No active RAS connection found.", vbInformation, "done"
        Else
            RasHangUp uRasConn(i).hRasConn
        End If
    Else
        MsgBox "Unable to enumerate active RAS connections! RasEnumConnections: " & CStr(lRet) & ".", vbCritical, "error"
    End If
End Sub

Private Sub mnuRASItem_Click(Index As Integer)
    Dim i%
    For i = 0 To 9
        mnuRASItem(i).Checked = False
    Next i
    mnuRASItem(Index).Checked = True
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "RASConnection", CLng(Index)
End Sub

Private Sub mnuRes10x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 1152, 864, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 1152, 864, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 1152, 864, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 1152, 864, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes11x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 1280, 1024, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 1280, 1024, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 1280, 1024, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 1280, 1024, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes12x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 1600, 1200, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 1600, 1200, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 1600, 1200, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 1600, 1200, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes13x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 1920, 1080, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 1920, 1080, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 1920, 1080, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 1920, 1080, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes14x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 2048, 1536, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 2048, 1536, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 2048, 1536, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 2048, 1536, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes1x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 300, 200, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 300, 200, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 300, 200, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 300, 200, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes2x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 300, 240, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 300, 240, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 300, 240, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 300, 240, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes3x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 400, 300, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 400, 300, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 400, 300, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 400, 300, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes4x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 480, 360, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 480, 360, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 480, 360, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 480, 360, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes5x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 512, 384, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 512, 384, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 512, 384, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 512, 384, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes6x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 640, 400, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 640, 400, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 640, 400, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 640, 400, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes7x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 640, 480, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 640, 480, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 640, 480, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 640, 480, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes8x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 800, 600, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 800, 600, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 800, 600, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 800, 600, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuRes9x_Click(Index As Integer)
    Select Case Index
        Case 0: SetDispMode 1024, 768, 8, mnuResConfirm.Checked
        Case 1: SetDispMode 1024, 768, 16, mnuResConfirm.Checked
        Case 2: SetDispMode 1024, 768, 24, mnuResConfirm.Checked
        Case 3: SetDispMode 1024, 768, 32, mnuResConfirm.Checked
    End Select
End Sub

Private Sub mnuResConfirm_Click()
    mnuResConfirm.Checked = Not mnuResConfirm.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ResolutionConfirm", Abs(CInt(mnuResConfirm.Checked))
End Sub

Private Sub mnuResCopy_Click()
    CopyToClipboard frmMain.lblResolution.Caption, 12
End Sub

Private Sub mnuTCPMonitorAdapter_Click(Index As Integer)
    Dim i%
    If mnuTCPMonitorAdapter(Index).Checked Then Exit Sub
    mnuTCPMonitorAll.Checked = False
    For i = 1 To 10
        mnuTCPMonitorAdapter(i).Checked = False
    Next i
    mnuTCPMonitorAdapter(Index).Checked = True
    mnuTCPMonitorIgnoreLoopback.Checked = False
    mnuTCPMonitorIgnoreLoopback.Enabled = False
    mnuTCPMonitorInfo.Enabled = True
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "TCPMonitorAdapter", CLng(Index)
    ClearTCPMonitorData
    frmMain.TriggerTimers MODULE_TCPMONITOR
End Sub

Private Sub mnuTCPMonitorAll_Click()
    Dim i%
    If mnuTCPMonitorAll.Checked Then Exit Sub
    For i = 1 To 10
        mnuTCPMonitorAdapter(i).Checked = False
    Next i
    mnuTCPMonitorAll.Checked = True
    mnuTCPMonitorIgnoreLoopback.Enabled = True
    mnuTCPMonitorIgnoreLoopback.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TCPMonitorIgnoreLoopback", 1))
    mnuTCPMonitorInfo.Enabled = False
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "TCPMonitorAdapter", 0
    ClearTCPMonitorData
    frmMain.TriggerTimers MODULE_TCPMONITOR
End Sub

Private Sub mnuTCPMonitorCopy_Click()
    CopyToClipboard frmMain.lblTCPMonitorDown.Caption & "|" & frmMain.lblTCPMonitorUp.Caption, 20
End Sub

Private Sub mnuTCPMonitorGraph_Click()
    mnuTCPMonitorGraph.Checked = Not mnuTCPMonitorGraph.Checked
    frmMain.picGraphTCP.Visible = mnuTCPMonitorGraph.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "TCPMonitorGraph", Abs(CInt(mnuTCPMonitorGraph.Checked))
    If mnuTCPMonitorGraph.Checked Then
        frmMain.fraTCPMonitor.Width = frmMain.picGraphTCP.Left + frmMain.picGraphTCP.Width
    Else
        frmMain.fraTCPMonitor.Width = frmMain.picGraphTCP.Left
    End If
    AlignModules bBarDocking
    frmMain.TriggerTimers MODULE_TCPMONITOR
End Sub

Private Sub mnuTCPMonitorIgnoreLoopback_Click()
    mnuTCPMonitorIgnoreLoopback.Checked = Not mnuTCPMonitorIgnoreLoopback.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "TCPMonitorIgnoreLoopback", Abs(CInt(mnuTCPMonitorIgnoreLoopback.Checked))
    frmMain.TriggerTimers MODULE_TCPMONITOR
End Sub

Private Sub mnuTCPMonitorInfo_Click()
    Dim uByte() As Byte, lSize&, sMsg$, sDummy$
    Dim IfRowTable As MIB_IFROW, i%, j%
    If GetIfTable(ByVal 0, lSize, 0) = ERROR_NOT_SUPPORTED Then
        frmMain.timTCPMonitor.Enabled = False
        MsgBox "The GetIfTable() API is not supported by your system. The TCP Monitor module will be disabled.", vbExclamation, "oops"
        Exit Sub
    End If
    ReDim uByte(lSize)
    If GetIfTable(uByte(0), lSize, 1) <> 0 Then
        frmMain.timTCPMonitor.Enabled = False
        MsgBox "The GetIfTable() API returned an error for some reason. The TCP Monitor module will be disabled.", vbExclamation, "oops"
        Exit Sub
    End If
    For i = 1 To 10
        If mnuTCPMonitorAdapter(i).Checked Then Exit For
    Next i
    CopyMemoryTCP IfRowTable, uByte(4 + (i - 1) * Len(IfRowTable)), Len(IfRowTable)
    With IfRowTable
        If .dwType = 0 Then Exit Sub
        sMsg = "========== Adapter #" & CStr(i) & " info ==========" & vbCrLf
        sDummy = Trim(Left(.bDescr, InStr(.bDescr, Chr(0)) - 1))
        sMsg = sMsg & "Description: " & sDummy & vbCrLf
        Select Case .dwType
            Case 1:  sMsg = sMsg & "Type: Other" & vbCrLf
            Case 6:  sMsg = sMsg & "Type: Ethernet" & vbCrLf
            Case 9:  sMsg = sMsg & "Type: Tokenring" & vbCrLf
            Case 15: sMsg = sMsg & "Type: FDDI" & vbCrLf
            Case 23: sMsg = sMsg & "Type: PPP" & vbCrLf
            Case 24: sMsg = sMsg & "Type: Loopback" & vbCrLf
            Case 28: sMsg = sMsg & "Type: Slip" & vbCrLf
        End Select
        If .dwPhysAddrLen <> 0 Then
            sDummy = ""
            For i = 1 To .dwPhysAddrLen
                sDummy = sDummy & IIf(Asc(Mid(.bPhysAddr, i, 1)) < 16, "0", "") & Hex(Asc(Mid(.bPhysAddr, i, 1))) & "-"
            Next i
            sDummy = Left(sDummy, Len(sDummy) - 1)
            sMsg = sMsg & "MAC address: " & sDummy & vbCrLf
        End If
        sMsg = sMsg & "Index: " & LongIP2DottedIP(.dwIndex) & vbCrLf
        Select Case .dwOperStatus
            Case 0: sMsg = sMsg & "Status: Not operational" & vbCrLf
            Case 1: sMsg = sMsg & "Status: Operational" & vbCrLf
            Case 2: sMsg = sMsg & "Status: Disconnected" & vbCrLf
            Case 3: sMsg = sMsg & "Status: Connecting" & vbCrLf
            Case 4: sMsg = sMsg & "Status: Connected" & vbCrLf
            Case 5: sMsg = sMsg & "Status: Unreachable" & vbCrLf
        End Select
        sMsg = sMsg & "Speed: " & Format(.dwSpeed, "###,###,###,##0") & " bits/sec"
        Select Case .dwSpeed
            Case Is < 1000:                sMsg = sMsg & vbCrLf
            Case 1000 To 1000 ^ 2 - 1:     sMsg = sMsg & " (" & CStr(.dwSpeed \ 1000) & " Kbit)" & vbCrLf
            Case 1000 ^ 2 To 1000 ^ 3 - 1: sMsg = sMsg & " (" & CStr(.dwSpeed \ 1000 ^ 2) & " Mbit)" & vbCrLf
            Case Is > 1000 ^ 3:            sMsg = sMsg & " (" & CStr(.dwSpeed \ 1000 ^ 3) & " Gbit)" & vbCrLf
        End Select
        sMsg = sMsg & "MTU: " & CStr(.dwMtu) & vbCrLf
        If .dwLastChange > 0 Then sMsg = sMsg & "Last change in status: " & CStr(.dwLastChange) & vbCrLf
        
        sMsg = sMsg & vbCrLf
        sMsg = sMsg & "Unicast packets received: " & Format(.dwInUcastPkts, "###,###,###,##0") & vbCrLf
        sMsg = sMsg & "Non-Unicast packets received: " & Format(.dwInNUcastPkts, "###,###,###,##0") & vbCrLf
        sMsg = sMsg & "Received packets discarded: " & Format(.dwInDiscards, "###,###,###,##0") & vbCrLf
        sMsg = sMsg & "Erroneous packets received: " & Format(.dwInErrors, "###,###,###,##0") & vbCrLf
        sMsg = sMsg & "Unknown packets received: " & Format(.dwInUnknownProtos, "###,###,###,##0") & vbCrLf
        sMsg = sMsg & "Total bytes received: " & Format(.dwInOctets, "###,###,###,##0")
        Select Case .dwInOctets
            Case Is < 1024:                sMsg = sMsg & vbCrLf
            Case 1024 To 1024 ^ 2 - 1:     sMsg = sMsg & " (" & CStr(.dwInOctets \ 1024) & " KB)" & vbCrLf
            Case 1024 ^ 2 To 1024 ^ 3 - 1: sMsg = sMsg & " (" & Left(CStr(.dwInOctets / 1024 ^ 2), 5) & " MB)" & vbCrLf
            Case Is > 1024 ^ 3:            sMsg = sMsg & " (" & Left(CStr(.dwInOctets / 1024 ^ 3), 5) & " GB)" & vbCrLf
        End Select
        
        sMsg = sMsg & vbCrLf
        sMsg = sMsg & "Unicast packets sent: " & Format(.dwOutUcastPkts, "###,###,###,##0") & vbCrLf
        sMsg = sMsg & "Non-Unicast packets sent: " & Format(.dwOutNUcastPkts, "###,###,###,##0") & vbCrLf
        sMsg = sMsg & "Sent packets discarded: " & Format(.dwOutDiscards, "###,###,###,##0") & vbCrLf
        sMsg = sMsg & "Erroneous packets sent: " & Format(.dwOutErrors, "###,###,###,##0") & vbCrLf
        sMsg = sMsg & "Output queue length: " & Format(.dwOutQLen, "###,###,###,##0") & vbCrLf
        sMsg = sMsg & "Total bytes sent: " & Format(.dwOutOctets, "###,###,###,##0")
        Select Case .dwOutOctets
            Case Is < 1024:                sMsg = sMsg & vbCrLf
            Case 1024 To 1024 ^ 2 - 1:     sMsg = sMsg & " (" & CStr(.dwOutOctets \ 1024) & " KB)" & vbCrLf
            Case 1024 ^ 2 To 1024 ^ 3 - 1: sMsg = sMsg & " (" & Left(CStr(.dwOutOctets / 1024 ^ 2), 5) & " MB)" & vbCrLf
            Case Is > 1024 ^ 3:            sMsg = sMsg & " (" & Left(CStr(.dwOutOctets / 1024 ^ 3), 5) & " GB)" & vbCrLf
        End Select
    End With
    If MsgBox(sMsg & vbCrLf & "Copy this to clipboard?", vbInformation + vbYesNo, "info") = vbYes Then
        Clipboard.Clear
        Clipboard.SetText sMsg
    End If
End Sub

Private Sub mnuTimeCopy_Click()
    CopyToClipboard frmMain.lblTime.Caption, 13
End Sub

Private Sub mnuTimeReminder_Click()
    frmMain.imgBanner.Tag = "setup"
    Load frmAlarm
    frmAlarm.Show 1
    
    'Dim iDummy%, sMsg$, sReminder$
    'On Error GoTo Error:
    'sMsg = "The Reminder function will display an alert messagebox "
    'sMsg = sMsg & "at a preset time with a preset message. "
    'sMsg = sMsg & "The Reminder can be once or daily." & vbCrLf & vbCrLf
    'sMsg = sMsg & "Click Yes to setup a new Reminder, No to "
    'sMsg = sMsg & "delete a previous one, Cancel to abort."
    'iDummy = MsgBox(sMsg, vbYesNoCancel + vbInformation, "reminder setup")
    'If iDummy = vbCancel Then Exit Sub
    'If iDummy = vbNo Then
    '    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "Reminder"
    '    bReminderSet = False
    '    Exit Sub
    'End If
'ReminderGetTime:
    'sMsg = "Enter the time at which the Reminder should run."
    'sMsg = sMsg & vbCrLf & "It should be in hh:mm:ss format."
    'sMsg = InputBox(sMsg, "Enter time", Time)
    'If sMsg = "" Then Exit Sub
    'If Len(sMsg) <> 8 Then
    '    If MsgBox("Invalid time specified. Try again?", vbExclamation + vbYesNo, "oops") = vbYes Then
    '        GoTo ReminderGetTime
    '    Else
    '        Exit Sub
    '    End If
    'End If
    'sReminder = sMsg & ";"
    '
    'sMsg = "Enter the message that should be displayed at " & Left(sReminder, 8) & "."
    'sMsg = sMsg & vbCrLf & "Do not use the semicolon (';') in the message."
    'sMsg = InputBox(sMsg, "Enter message")
    'If sMsg = "" Then Exit Sub
    'sMsg = Replace(sMsg, ";", ",")
    'sReminder = sReminder & sMsg & ";"
    '
    'sMsg = "Should the Reminder be daily or once?" & vbCrLf
    'sMsg = sMsg & "Click Yes for daily or No for once."
    'If MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, "daily") = vbYes Then
    '    sReminder = sReminder & "1"
    'Else
    '    sReminder = sReminder & "0"
    'End If
    'RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "Reminder", sReminder
    'bReminderSet = True
    'MsgBox "Reminder set!", vbInformation, "done"
    'Exit Sub
    '
'Error:
    'ShowError "Time", "frmMenu.mnuTimeReminder_Click", Err.Number, Err.Description, False
End Sub

Private Sub mnuTimeSync_Click()
    Dim iTimeServerIndex%, i%, sTimeServer$, sExtraServer
    
    If MsgBox("The time sync subroutine is not finished yet " & _
              "and causes a GPF I cannot explain yet. To " & _
              "see it, hit OK. To not see it, hit Cancel." & _
              vbCrLf & vbCrLf & "(You need to be connected " & _
              "to the Internet for this to work.)", vbCritical + vbOKCancel, "sure?") = vbCancel Then Exit Sub
    
    iTimeServerIndex = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TimeServerIndex", 0))
    For i = 0 To 99
        sExtraServer = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "ExtraTimeServer" & String(2 - Len(CStr(i)), "0") & CStr(i))
        If sExtraServer <> "" Then
            'extra servers present - fetch right one
            If i = iTimeServerIndex Then
                sTimeServer = sExtraServer
                Exit For
            End If
        Else
            'index goes beyond # of extra servers, use
            'predefined servers
            If iTimeServerIndex = i Then
                sTimeServer = "time.ien.it"
            ElseIf iTimeServerIndex = i + 1 Then
                sTimeServer = "ntps1-0.cs.tu-berlin.de"
            ElseIf iTimeServerIndex = i + 2 Then
                sTimeServer = "ntp-cup.external.hp.com"
            Else
                'index must be invalid - pick one
                sTimeServer = "time.ien.it"
            End If
            Exit For
        End If
    Next i
    frmMain.timTime.Enabled = False
    DoEvents
    
    'sTimeServer should now hold correct time server
    'If MsgBox("sync time with server '" & sTimeServer & "'?", vbYesNo + vbQuestion, "blah") = vbNo Then Exit Sub
    'ConnectToServer "127.0.0.1", 80&, frmMain.hwnd
    timServerTimeOut.Tag = sTimeServer
    ConnectToServer sTimeServer, 37&, frmMain.hwnd
End Sub

Private Sub mnuUninstall_Click()
    If MsgBox("This will delete all Uptimer4 settings from the Registry and close Uptimer4. Are you sure?", vbYesNo + vbExclamation, "uninstall <sob sob>") = vbNo Then Exit Sub
    
    frmMain.Hide
    Unload frmMain
    
    RegDelKey HKEY_LOCAL_MACHINE, sKeySettings
    RegDelKey HKEY_LOCAL_MACHINE, "Software\Soeperman Enterprises Ltd.\Uptimer4"
    'This may not be such a smart idea.
    'Maybe use RegEnum to check for other SEL progs?
    'RegDelKey HKEY_LOCAL_MACHINE, "Software\Soeperman Enterprises Ltd."
    
    If RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Uptimer4") <> "" Then
        RegDelValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Uptimer4"
    End If
    
    If RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\uptimer4.exe", "") <> "" Then
        RegDelKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\uptimer4.exe"
    End If
    
    End
End Sub

Private Sub mnuUptimeBar_Click()
    mnuUptimeBar.Checked = Not mnuUptimeBar.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "UptimeShowBar", Abs(CInt(mnuUptimeBar.Checked))
    frmMain.shpUptimeBack.Visible = mnuUptimeBar.Checked
    frmMain.shpUptimeFore.Visible = mnuUptimeBar.Checked
    frmMain.TriggerTimers MODULE_UPTIME
End Sub

Private Sub mnuUptimeCopy_Click()
    CopyToClipboard frmMain.lblUptime.Caption, 15
End Sub

Private Sub mnuUptimeGetBoot_Click()
    Dim sRawBootTime$, sMsg$
    sMsg = "This computer was booted on:" & vbCrLf
    sRawBootTime = CStr(DateAdd("s", -GetTickCount() \ 1000, Date + Time))
    sMsg = sMsg & Format(sRawBootTime, "Long Date") & ", or " & vbCrLf
    sMsg = sMsg & Format(sRawBootTime, "Short Date") & ", at "
    sMsg = sMsg & Format(sRawBootTime, "Long Time") & "."
    sMsg = sMsg & vbCrLf & vbCrLf & "Copy to clipboard?"
    If MsgBox(sMsg, vbInformation + vbYesNo, "boot date/time") = vbYes Then
        If RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MircColors", 0) = 0 Then
            sMsg = "Boot date/time: " & Format(sRawBootTime, "Short Date")
            sMsg = sMsg & Format(sRawBootTime, "Long Time")
        Else
            sMsg = "Boot date/time:" & Chr(3) & "12 "
            sMsg = sMsg & Format(sRawBootTime, "Short Date")
            sMsg = sMsg & ", " & Format(sRawBootTime, "Long Time") & Chr(3)
        End If
        Clipboard.Clear
        Clipboard.SetText sMsg
    End If
End Sub

Private Sub mnuUptimeLoggingCleanUp_Click()
    Dim sLine$, i%, sLongestLine$
    Dim nLongestUptime!, nUptime!
    
    On Error Resume Next
    If sUptimeLogLocation = "" Or Dir(sUptimeLogLocation) = "" Then
        sUptimeLogLocation = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Uptimer4.log"
        If Dir(sUptimeLogLocation) = "" Then Exit Sub
    End If
    On Error GoTo Error:
    
    Open sUptimeLogLocation For Input As #1
        Do
            Line Input #1, sLine
            i = InStr(sLine, "|")
            If i <> 0 Then
                i = InStr(i + 1, sLine, "|")
                If i <> 0 Then
                    'Get uptime from line
                    nUptime = CSng(Val(Mid(sLine, i + 1)))
                    'Convert negative to positive
                    If nUptime < 0 Then nUptime = nUptime + 2 ^ 32
                    'Higher than last? Record line + uptime
                    If nUptime > nLongestUptime Then
                        sLongestLine = sLine
                        nLongestUptime = nUptime
                    End If
                End If
            End If
        Loop Until EOF(1)
    Close #1
    nLongestUptime = FileLen(sUptimeLogLocation)
    Open sUptimeLogLocation For Output As #1
        Print #1, sLongestLine
    Close #1
    MsgBox "Logfile cleaned up! File size reduced from " & _
           Format(nLongestUptime, "###,###,###") & " bytes " & _
           "to " & CStr(FileLen(sUptimeLogLocation)) & _
           " bytes.", vbInformation, "done"
    Exit Sub
    
Error:
    Close #1
    ShowError "Uptime", "mnuUptimeLoggingCleanUp_Click", Err.Number, Err.Description, False
End Sub

Private Sub mnuUptimeLoggingClear_Click()
    If MsgBox("This will clear your entire uptime log. Continue?", vbQuestion + vbYesNo, "sure?") = vbNo Then Exit Sub
    On Error Resume Next
    If sUptimeLogLocation = "" Or Dir(sUptimeLogLocation) = "" Then
        sUptimeLogLocation = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Uptimer4.log"
    End If
    Open sUptimeLogLocation For Output As #1
    Close #1
End Sub

Private Sub mnuUptimeLoggingEnable_Click()
    mnuUptimeLoggingEnable.Checked = Not mnuUptimeLoggingEnable.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "UptimeEnableLogging", Abs(CInt(mnuUptimeLoggingEnable.Checked))
    If mnuUptimeLoggingEnable.Checked Then
        mnuUptimeLoggingWriteHourly.Enabled = True
        mnuUptimeLoggingWriteNow.Enabled = True
        mnuUptimeLoggingView.Enabled = True
        mnuUptimeLoggingCleanUp.Enabled = True
        mnuUptimeLoggingClear.Enabled = True
        mnuUptimeLoggingGetLongest.Enabled = True
    Else
        mnuUptimeLoggingWriteHourly.Enabled = False
        mnuUptimeLoggingWriteNow.Enabled = False
        mnuUptimeLoggingView.Enabled = False
        mnuUptimeLoggingCleanUp.Enabled = False
        mnuUptimeLoggingClear.Enabled = False
        mnuUptimeLoggingGetLongest.Enabled = False
    End If
End Sub

Private Sub mnuUptimeLoggingGetLongest_Click()
    GetLongestUptime True
End Sub

Private Sub mnuUptimeLoggingView_Click()
    ShellExecute Me.hwnd, "open", sUptimeLogLocation, "", "", 1
End Sub

Private Sub mnuUptimeLoggingWriteHourly_Click()
    mnuUptimeLoggingWriteHourly.Checked = Not mnuUptimeLoggingWriteHourly.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "UptimeLogHourly", Abs(CInt(mnuUptimeLoggingWriteHourly.Checked))
    bUptimerLogHourly = mnuUptimeLoggingWriteHourly.Checked
End Sub

Private Sub mnuUptimeLoggingWriteNow_Click()
    LogUptime
End Sub

Private Sub mnuWinampClose_Click()
    Dim hwndWinamp&
    hwndWinamp = FindWindow("Winamp v1.x", vbNullString)
    If hwndWinamp <> 0 Then PostMessage hwndWinamp, WM_CLOSE, 0, 0
End Sub

Private Sub mnuWinampGetVer_Click()
    Dim sWinampPath$
    sWinampPath = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "WinampPath")
    If sWinampPath = "" Or LCase(Dir(sWinampPath)) <> "winamp.exe" Then
        MsgBox "Unable to find Winamp.exe. Click the Winamp icon to locate it.", vbExclamation, "damn"
        Exit Sub
    End If
    
    Dim lBufferLen&, uBuffer() As Byte, sVersion$
    Dim uVerInfo As VS_FIXEDFILEINFO, lVerPtr&
    lBufferLen = GetFileVersionInfoSize(sWinampPath, ByVal 0)
    If lBufferLen = 0 Then
        MsgBox "Unable to get Winamp.exe version. GetLastError: " & GetLastError()
        Exit Sub
    End If
    ReDim uBuffer(lBufferLen)
    GetFileVersionInfo sWinampPath, 0, lBufferLen, uBuffer(0)
    VerQueryValue uBuffer(0), "\", lVerPtr, ByVal 0
    CopyMemory uVerInfo, ByVal lVerPtr, Len(uVerInfo)
    With uVerInfo
        sVersion = CStr(.dwFileVersionMSh) & "." & CStr(.dwFileVersionMSl) & CStr(.dwFileVersionLSh)
    End With
    MsgBox "Your Winamp version is: " & sVersion & ".", vbInformation, "winamp"
End Sub

Private Sub mnuWinampHotkeyMode_Click(Index As Integer)
    If mnuWinampHotkeyMode(Index).Checked Then Exit Sub
    mnuWinampHotkeyMode(1).Checked = False
    mnuWinampHotkeyMode(2).Checked = False
    mnuWinampHotkeyMode(3).Checked = False
    mnuWinampHotkeyMode(Index).Checked = True
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "WinampHotkeysMode", CLng(Index)
    If Not RunningInIDE Then RegisterHotkeys False, 0
    If Not RunningInIDE Then RegisterHotkeys True, Index
End Sub

Public Sub mnuWinampHotkeys_Click()
    mnuWinampHotkeys.Checked = Not mnuWinampHotkeys.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "WinampHotkeys", Abs(CInt(mnuWinampHotkeys.Checked))
    If mnuWinampHotkeys.Checked Then
        If Not RunningInIDE Then RegisterHotkeys True, CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "WinampHotkeysMode", 3))
        mnuWinampHotkeyMode(1).Enabled = True
        mnuWinampHotkeyMode(2).Enabled = True
        mnuWinampHotkeyMode(3).Enabled = True
        mnuWinampHotkeyMode(CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "WinampHotkeysMode", 3))).Checked = True
    Else
        If Not RunningInIDE Then RegisterHotkeys False, 0
        mnuWinampHotkeyMode(1).Enabled = False
        mnuWinampHotkeyMode(2).Enabled = False
        mnuWinampHotkeyMode(3).Enabled = False
        mnuWinampHotkeyMode(1).Checked = False
        mnuWinampHotkeyMode(2).Checked = False
        mnuWinampHotkeyMode(3).Checked = False
    End If
End Sub

Private Sub mnuWinAmpMin_Click()
    mnuWinAmpMin.Checked = Not mnuWinAmpMin.Checked
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "WinAmpMin", Abs(CInt(mnuWinAmpMin.Checked))
End Sub

Private Sub mnuWinAmpStart_Click()
    frmMain.imgWinamp_MouseUp 1, 0, 0, 0
End Sub

Private Sub picTrayIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case X \ Screen.TwipsPerPixelX
        Case WM_RBUTTONUP
            If CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TrayIcon", 0)) = True Then mnuMainHideIcon.Visible = True
            PopupMenu mnuMain
            mnuMainHideIcon.Visible = False
        Case WM_LBUTTONUP
            If Shift = 1 Then
                If Not bBarDocking Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                frmModules.Show 1
                If bAlwaysOnTop Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            ElseIf Shift = 2 Then
                If Not bBarDocking Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                Load frmSettings
                frmSettings.Show 1
                If bAlwaysOnTop Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            End If
    End Select
End Sub

Private Sub picWebserverAccept_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Output object for webserver incoming connections
    Dim uReadSockAddr As sockaddr
    'Accept connection and get socket
    lReadSockNum = accept(lSockNum, uReadSockAddr, 16)
    
    'Set output object for incoming (and outgoing) data
    If WSAAsyncSelect(lReadSockNum, picWebserverRead.hwnd, ByVal &H202, FD_READ Or FD_CLOSE) <> 0 Then
        closesocket lReadSockNum
        MsgBox "Unable to select output object for incoming data!", vbExclamation, "Uptimer4 webserver"
        Exit Sub
    End If
End Sub

Private Sub picWebserverRead_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lDataLen&, sData As String * 1024, uTheMsg() As Byte
    Dim vCommand As Variant, sCommand$, sHTMLPage$
    On Error GoTo Error:
    
    lDataLen = recv(lReadSockNum, sData, 1024, 0)
    If lDataLen > 0 Then
        sData = Left(sData, lDataLen)
        sCommand = Replace(sData, vbCrLf, " ")
        If Len(sCommand) > 5 Then
            If Left(sCommand, 5) = "GET /" Then
                'Send HTML page back to client
                lServed = lServed + 1
                sHTMLPage = GetHTMLPage(sCommand)
                uTheMsg = ""
                uTheMsg = StrConv(sHTMLPage, vbFromUnicode)
                If UBound(uTheMsg) > -1 Then send lReadSockNum, uTheMsg(0), UBound(uTheMsg) + 1, 0
                closesocket lReadSockNum
                sLastServedTime = CStr(Time)
                sLastServedDate = CStr(Date)
            End If
        End If
    Else
        'Client disconnected
        closesocket lReadSockNum
    End If
    Exit Sub
    
Error:
    ShowError "Main", "frmMenu.picWebserverRead_MouseUp", Err.Number, Err.Description, False
End Sub

Private Function GetHTMLPage(sClientCmd$) As String
    'Function creates dynamic HTML page and returns it
    Dim sHeader$, sPage$, sLine$
    On Error GoTo Error:
    
    'Create HTTP server header
    sHeader = "HTTP/1.1 200 OK" & vbCrLf
    sHeader = sHeader & "Date: " & Format(Date, "Long date") & " "
    sHeader = sHeader & Format(Time, "Long time") & " GMT" & vbCrLf
    sHeader = sHeader & "Server: Uptimer4 Mini Webserver" & vbCrLf
    sHeader = sHeader & "Connection: close" & vbCrLf
    sHeader = sHeader & "Keep-Alive: timeout=3, max=5" & vbCrLf
    sHeader = sHeader & "Content-Length: xXx" & vbCrLf
    sHeader = sHeader & "Content-type: text/html" & vbCrLf
    sHeader = sHeader & "Last-Modified: " & Format(Date, "Long date") & " "
    sHeader = sHeader & Format(Time, "Long time") & " GMT" & vbCrLf & vbCrLf
    
    'Read and send uptimer4.html if present
    If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "uptimer4.html") <> "" Then
        Open App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "uptimer4.html" For Input As #1
            Do
                Line Input #1, sLine
                sPage = sPage & sLine & vbCrLf
            Loop Until EOF(1)
        Close #1
        sPage = ExpandAliases(sPage, sClientCmd)
        sHeader = Replace(sHeader, "xXx", CStr(Len(sPage)))
        GetHTMLPage = sHeader & sPage
        Exit Function
    End If
    
    'Create dynamic page if uptimer4.html is not present
    sPage = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"" ""http://www.w3.org/TR/REC-html40/loose.dtd"">" & vbCrLf
    sPage = sPage & "<HTML><HEAD><TITLE>Uptimer4 webserver</TITLE>" & vbCrLf
    sPage = sPage & "<META HTTP-EQUIV=""content-type"" CONTENT=""text/html; charset=ISO-8859-1"">" & vbCrLf
    sPage = sPage & "<META NAME=""Author"" CONTENT=""Klont"">" & vbCrLf
    sPage = sPage & "<META NAME=""Generator"" CONTENT=""Uptimer4 webserver""></HEAD>" & vbCrLf & vbCrLf
    sPage = sPage & "<BODY><FONT FACE=""Fixedsys"">" & vbCrLf
    sPage = sPage & "<FONT FACE=""Arial"" SIZE=+2><CENTER><B>Welcome to the Uptimer4 webserver!" & vbCrLf
    sPage = sPage & "</B></CENTER></FONT><BLOCKQUOTE>" & vbCrLf
    sPage = sPage & "Create <B>uptimer4.html</B> in the Uptimer4 directory to replace this one.<BR>" & vbCrLf
    sPage = sPage & "To display information about your server, use the following aliases:<UL>" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERNAME = $SERVERNAME" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERVERSION = $SERVERVERSION" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERBOOTTIME = $SERVERBOOTTIME" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERBOOTDATESHORT = $SERVERBOOTDATESHORT" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERBOOTDATELONG = $SERVERBOOTDATELONG" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->PAGESSERVED = $PAGESSERVED" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->LASTSERVEDTIME = $LASTSERVEDTIME" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->LASTSERVEDDATESHORT = $LASTSERVEDDATESHORT" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->LASTSERVEDDATELONG = $LASTSERVEDDATELONG" & vbCrLf
    
    sPage = sPage & "<LI>$<!--.-->HOMEPAGE = $HOMEPAGE" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERIP1 = $SERVERIP1" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERIP2 = $SERVERIP2" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERIP3 = $SERVERIP3" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERIP4 = $SERVERIP4" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERIP5 = $SERVERIP5" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERIP6 = $SERVERIP6" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERIP7 = $SERVERIP7" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERIP8 = $SERVERIP8" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->SERVERIP9 = $SERVERIP9" & vbCrLf
    
    sPage = sPage & "<LI>$<!--.-->TIME = $TIME" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->DATESHORT = $DATESHORT" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->DATELONG = $DATELONG" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->OS = $OS" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->FREERAM = $FREERAM" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->FREEPAGEFILE = $FREEPAGEFILE" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->DISKFREESPACE = $DISKFREESPACE" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->DISKFREESDISK = $DISKFREEDISK" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->RESOLUTION = $RESOLUTION" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->TCPUP = $TCPUP" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->TCPDOWN = $TCPDOWN" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->MSIEVER = $MSIEVER" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->DXVER = $DXVER" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->RASCONNECTION = $RASCONNECTION" & vbCrLf
    
    sPage = sPage & "<LI>$<!--.-->TICKS = $TICKS" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->UPTIMESHORT = $UPTIMESHORT" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->UPTIMELONG = $UPTIMELONG" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->LONGESTUPTIMESHORT = $LONGESTUPTIMESHORT" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->LONGESTUPTIMELONG = $LONGESTUPTIMELONG" & vbCrLf
    
    sPage = sPage & "<LI>$<!--.-->CLIENTNAME = $CLIENTNAME" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->CLIENTIP = $CLIENTIP" & vbCrLf
    sPage = sPage & "<LI>$<!--.-->CLIENTREQUEST = $CLIENTREQUEST" & vbCrLf
    sPage = sPage & "</UL>The latest version of Uptimer4 is always " & vbCrLf
    sPage = sPage & "available at $HOMEPAGE.<BR>" & vbCrLf
    sPage = sPage & "<BLOCKQUOTE><FONT COLOR=red>Klont, <A HREF=""mailto:klont@windhoos2000.nl?subject=Uptimer4"">" & vbCrLf
    sPage = sPage & "klont@windhoos2000.nl</A></FONT></BLOCKQUOTE>" & vbCrLf
    sPage = sPage & "</BLOCKQUOTE></FONT></BODY></HTML>" & vbCrLf
    
    sPage = ExpandAliases(sPage, sClientCmd)
    sHeader = Replace(sHeader, "xXx", CStr(Len(sPage)))
    GetHTMLPage = sHeader & sPage
    Exit Function
    
Error:
    ShowError "Main", "frmMenu.GetHTMLPage", Err.Number, Err.Description, False
End Function

Private Function ExpandAliases(sText$, sRequest$) As String
    'Try to expand the following aliases:
    '$SERVERNAME
    '$SERVERVERSION
    '$SERVERBOOTTIME
    '$SERVERBOOTDATESHORT
    '$SERVERBOOTDATELONG
    '$PAGESSERVED
    '$LASTSERVEDTIME
    '$LASTSERVEDDATESHORT
    '$LASTSERVEDDATELONG
    
    '$HOMEPAGE
    '$SERVERIP1 thru $SERVERIP9
    
    '$TIME
    '$DATESHORT
    '$DATELONG
    '$OS
    '$FREERAM
    '$FREEPAGEFILE
    '$DISKFREESPACE
    '$DISKFREEDISK
    '$RESOLUTION
    '$TCPUP
    '$TCPDOWN
    '$MSIEVER
    '$DXVER
    '$RASCONNECTION
    
    '$TICKS
    '$UPTIMESHORT
    '$UPTIMELONG
    '$LONGESTUPTIMESHORT
    '$LONGESTUPTIMELONG
    
    '$CLIENTNAME
    '$CLIENTIP
    '$CLIENTREQUEST
    
    '========================================
    On Error GoTo Error:
    'Server name and IP addresses
    If frmMain.timIPs.Enabled = False Then
        frmMain.timIPs.Enabled = True
        frmMain.TriggerTimers MODULE_IPS
        frmMain.timIPs.Enabled = False
    End If
    sText = Replace(sText, "$SERVERNAME", sHostname)
    sText = Replace(sText, "$SERVERIP1", sIPs(0))
    If UBound(sIPs) > 0 Then sText = Replace(sText, "$SERVERIP2", sIPs(1)) Else sText = Replace(sText, "$SERVERIP2", "?")
    If UBound(sIPs) > 1 Then sText = Replace(sText, "$SERVERIP3", sIPs(2)) Else sText = Replace(sText, "$SERVERIP3", "?")
    If UBound(sIPs) > 2 Then sText = Replace(sText, "$SERVERIP4", sIPs(3)) Else sText = Replace(sText, "$SERVERIP4", "?")
    If UBound(sIPs) > 3 Then sText = Replace(sText, "$SERVERIP5", sIPs(4)) Else sText = Replace(sText, "$SERVERIP5", "?")
    If UBound(sIPs) > 4 Then sText = Replace(sText, "$SERVERIP6", sIPs(5)) Else sText = Replace(sText, "$SERVERIP6", "?")
    If UBound(sIPs) > 5 Then sText = Replace(sText, "$SERVERIP7", sIPs(6)) Else sText = Replace(sText, "$SERVERIP7", "?")
    If UBound(sIPs) > 6 Then sText = Replace(sText, "$SERVERIP8", sIPs(7)) Else sText = Replace(sText, "$SERVERIP8", "?")
    If UBound(sIPs) > 7 Then sText = Replace(sText, "$SERVERIP9", sIPs(8)) Else sText = Replace(sText, "$SERVERIP9", "?")
    
    'Server system info
    Dim sRawBootTime$
    If InStr(sText, "$SERVERVERSION") > 0 Then sText = Replace(sText, "$SERVERVERSION", CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision))
    If InStr(sText, "$SERVERBOOTTIME") > 0 Then
        sRawBootTime = CStr(DateAdd("s", -GetTickCount() \ 1000, Date + Time))
        sText = Replace(sText, "$SERVERBOOTTIME", Format(sRawBootTime, "Long Time"))
    End If
    If InStr(sText, "$SERVERBOOTDATESHORT") > 0 Or InStr(sText, "$SERVERBOOTDATELONG") > 0 Then
        sRawBootTime = CStr(DateAdd("s", -GetTickCount() \ 1000, Date + Time))
        sText = Replace(sText, "$SERVERBOOTDATESHORT", Format(sRawBootTime, "Short Date"))
        sText = Replace(sText, "$SERVERBOOTDATELONG", Format(sRawBootTime, "Long Date"))
    End If
    If InStr(sText, "$PAGESSERVED") > 0 Then sText = Replace(sText, "$PAGESSERVED", CStr(lServed))
    If InStr(sText, "$LASTSERVEDTIME") > 0 Then sText = Replace(sText, "$LASTSERVEDTIME", sLastServedTime)
    If InStr(sText, "$LASTSERVEDDATESHORT") > 0 Then sText = Replace(sText, "$LASTSERVEDDATESHORT", sLastServedDate)
    If InStr(sText, "$LASTSERVEDDATELONG") > 0 Then sText = Replace(sText, "$LASTSERVEDDATELONG", Format(sLastServedDate, "Long Date"))
    If InStr(sText, "$HOMEPAGE") > 0 Then sText = Replace(sText, "$HOMEPAGE", "<A HREF=""http://www.geocities.com/merijn_bellekom/new/uptimer.html"">http://www.geocities.com/merijn_bellekom/new/uptimer.html</A>")
    If InStr(sText, "$TIME") > 0 Then sText = Replace(sText, "$TIME", CStr(Time))
    If InStr(sText, "$DATESHORT") > 0 Then sText = Replace(sText, "$DATESHORT", Format(Date, "Short date"))
    If InStr(sText, "$DATELONG") > 0 Then sText = Replace(sText, "$DATELONG", Format(Date, "Long date"))
    If InStr(sText, "$OS") > 0 Then sText = Replace(sText, "$OS", frmMain.lblOS.Caption)
    If InStr(sText, "$FREERAM") > 0 Then
        If frmMain.timMemoryRAM.Enabled = False Then
            frmMain.timMemoryRAM.Enabled = True
            frmMain.TriggerTimers MODULE_FREERAM
            frmMain.timMemoryRAM.Enabled = False
        End If
        sText = Replace(sText, "$FREERAM", frmMain.lblMemoryRAM.Caption)
    End If
    If InStr(sText, "$FREEPAGEFILE") > 0 Then
        If frmMain.timMemoryPage.Enabled = False Then
            frmMain.timMemoryPage.Enabled = True
            frmMain.TriggerTimers MODULE_FREEPAGEFILE
            frmMain.timMemoryPage.Enabled = False
        End If
        sText = Replace(sText, "$FREEPAGEFILE", frmMain.lblMemoryPage.Caption)
    End If
    If InStr(sText, "$DISKFREESPACE") > 0 Or InStr(sText, "$DISKFREEDISK") > 0 Then
        If frmMain.timDiskFreeSpace.Enabled = False Then
            frmMain.timDiskFreeSpace.Enabled = True
            frmMain.TriggerTimers MODULE_DISKFREESPACE
            frmMain.timDiskFreeSpace.Enabled = False
        End If
        sText = Replace(sText, "$DISKFREESPACE", frmMain.lblDiskFreeSpace.Caption)
        If sCurrentDisk = "." Then
            sText = Replace(sText, "$DISKFREEDISK", "all")
        Else
            sText = Replace(sText, "$DISKFREEDISK", sCurrentDisk & ":\")
        End If
    End If
    If InStr(sText, "$RESOLUTION") > 0 Then sText = Replace(sText, "$RESOLUTION", frmMain.lblResolution.Caption)
    If InStr(sText, "$TCPUP") > 0 Or InStr(sText, "$TCPDOWN") > 0 Then
        If frmMain.timTCPMonitor.Enabled Then
            sText = Replace(sText, "$TCPUP", frmMain.lblTCPMonitorUp.Caption)
            sText = Replace(sText, "$TCPDOWN", frmMain.lblTCPMonitorDown.Caption)
        Else
            sText = Replace(sText, "$TCPUP", "disabled")
            sText = Replace(sText, "$TCPDOWN", "disabled")
        End If
    End If
    If InStr(sText, "$MSIEVER") > 0 Then sText = Replace(sText, "$MSIEVER", frmMain.lblMSIE.Caption)
    If InStr(sText, "$DXVER") > 0 Then sText = Replace(sText, "$DXVER", frmMain.lblDX.Caption)
    If InStr(sText, "$RASCONNECTION") > 0 Then sText = Replace(sText, "$RASCONNECTION", Mid(frmMain.imgRAS(0).ToolTipText, 18))
    
    'Server uptime info
    If InStr(sText, "$TICKS") > 0 Then sText = Replace(sText, "$TICKS", CStr(GetTickCount()))
    If InStr(sText, "$UPTIMESHORT") > 0 Or InStr(sText, "$UPTIMELONG") > 0 Or _
       InStr(sText, "$LONGESTUPTIMESHORT") > 0 Or InStr(sText, "$LONGESTUPTIMELONG") > 0 Then
        If frmMain.timUptime.Enabled = False Then
            frmMain.timUptime.Enabled = True
            frmMain.TriggerTimers MODULE_UPTIME
            frmMain.timUptime.Enabled = False
        End If
        sText = Replace(sText, "$UPTIMESHORT", frmMain.lblUptime.Caption)
        sText = Replace(sText, "$UPTIMELONG", GetDuration(GetTickCount(), True))
        sText = Replace(sText, "$LONGESTUPTIMESHORT", GetDuration(GetLongestUptime(False), False))
        sText = Replace(sText, "$LONGESTUPTIMELONG", GetDuration(GetLongestUptime(False), True))
    End If
    
    'Client info
    If InStr(sText, "$CLIENTIP") > 0 Then sText = Replace(sText, "$CLIENTIP", GetClientIP(False))
    If InStr(sText, "$CLIENTNAME") > 0 Then sText = Replace(sText, "$CLIENTNAME", GetClientIP(True))
    If InStr(sText, "$CLIENTREQUEST") > 0 Then sText = Replace(sText, "$CLIENTREQUEST", sRequest)
    
    ExpandAliases = sText
    Exit Function
    
Error:
    ShowError "Main", "frmMenu.ExpandAliases", Err.Number, Err.Description, False
End Function

Private Function GetClientIP(bResolveName As Boolean) As String
    Dim uSockAddr As sockaddr
    On Error GoTo Error:
    If getpeername(lReadSockNum, uSockAddr, 16) <> 0 Then
        GetClientIP = "?"
        Exit Function
    End If
    
    'Get dotted IP from long IP
    Dim lpStrIP As Long, lPointerLen As Long, sDottedIP As String
    lpStrIP = inet_ntoa(uSockAddr.sin_addr)
    If lpStrIP = 0 Then
        GetClientIP = "?"
        Exit Function
    End If
    
    lPointerLen = 15 'length of '255.255.255.255', max IP
    sDottedIP = String(32, Chr(0))
    CopyMemoryGCI ByVal sDottedIP, ByVal lpStrIP, lPointerLen
    sDottedIP = Left(sDottedIP, InStr(sDottedIP, Chr(0)) - 1)
    If bResolveName = False Then
        GetClientIP = sDottedIP
        Exit Function
    End If
    
    Dim lRet As Long, uHostEnt As HOSTENT, sMyHostname As String * 255
    lRet = gethostbyaddr(uSockAddr.sin_addr, 4, 2)
    If lRet = 0 Then
        GetClientIP = "?"
        Exit Function
    End If
    CopyMemoryGCI uHostEnt, lRet, Len(uHostEnt)
    CopyMemoryGCI ByVal sMyHostname, uHostEnt.hName, 255
    GetClientIP = Left(sMyHostname, InStr(sMyHostname, Chr(0)) - 1)
    Exit Function
    
Error:
    ShowError "Main", "frmMenu.GetClientIP", Err.Number, Err.Description, False
End Function

Private Function GetInterval(sModuleName$, iDefault%, iCurrent%, iMin%, iMax%) As Integer
    Load frmInterval
    With frmInterval
        .lblInfo(0).Caption = Replace(.lblInfo(0).Caption, "[..]", sModuleName)
        .lblInfo(1).Caption = CStr(iMin)
        .lblInfo(2).Caption = CStr(iMax)
        .hscInterval.Min = iMin
        .hscInterval.Max = iMax
        .hscInterval.Value = iCurrent
        .cmdDefault.Tag = CStr(iDefault)
        .Show 1
        GetInterval = CInt(Val(picTrayIcon.Tag))
        picTrayIcon.Tag = ""
    End With
End Function

Private Sub CopyToClipboard(sText$, iModule%)
    Dim bMirc As Boolean
    On Error GoTo Error:
    bMirc = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MircColors", 0))
    Clipboard.Clear
    Select Case iModule
        Case 2 'CPU usage
            Clipboard.SetText "CPU usage:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 3 'Date
            Clipboard.SetText "Date:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 4 'Disk free space
            Clipboard.SetText "Disk free space:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 6 'Free pagefile
            Clipboard.SetText "Free pagefile:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 7 'Free RAM
            Clipboard.SetText "Free RAM:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 8 'IP addresses
            Clipboard.SetText "IP address:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 11 'Power status
            Clipboard.SetText "Power status:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 12 'Screen resolution
            Clipboard.SetText "Screen resolution:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 13 'Time
            Clipboard.SetText "Time:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 15 'Uptime
            Clipboard.SetText "Uptime:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 17 'Windows version
            Clipboard.SetText "Windows version:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        'Add more modules below
        Case 20 'TCP Monitor
            Clipboard.SetText "TCP throughput: Down" & IIf(bMirc, Chr(3) & "12 ", " ") & Left(sText, InStr(sText, "|") - 1) & IIf(bMirc, Chr(3), "") & " Up" & IIf(bMirc, Chr(3) & "12 ", " ") & Mid(sText, InStr(sText, "|") + 1) & IIf(bMirc, Chr(3), "")
        Case 21 'MSIE version
            Clipboard.SetText "MSIE version:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
        Case 22 'DirectX version
            Clipboard.SetText "DirectX version:" & IIf(bMirc, Chr(3) & "12 " & sText & Chr(3), " " & sText)
    End Select
    Exit Sub
    
Error:
    ShowError "Main", "frmMenu.CopyToClipboard(" & sText & "," & CStr(iModule) & ")", Err.Number, Err.Description, False
End Sub

Private Function GetDiskLabel(sPath$) As String
    Dim lType&
    On Error Resume Next
    GetDiskLabel = ""
    lType = GetDriveType(sPath)
    If lType = DRIVE_FIXED Or lType = DRIVE_RAMDISK Or lType = DRIVE_REMOTE Then
        'GetDiskLabel = LCase(Dir(sPath, vbVolume))
        GetDiskLabel = String(25, 0)
        GetVolumeInformation sPath, GetDiskLabel, 25, ByVal 0, ByVal 0, ByVal 0, vbNullString, ByVal 0
        GetDiskLabel = Trim(Replace(GetDiskLabel, Chr(0), " "))
        If GetDiskLabel <> "" Then GetDiskLabel = "   [" & GetDiskLabel & "]"
    End If
End Function

Public Sub EnumRASConn()
    Dim uREN(255) As RASENTRYNAME95, lEntries&, i%, j$
    On Error GoTo Error:
    uREN(0).dwSize = Len(uREN(0)) + Len(uREN(0)) Mod 4
    'Enumerate all RAS entries, check for rasapi32.dll
    On Error Resume Next
    RasEnumEntries "", "", uREN(0), 255 * uREN(0).dwSize, lEntries
    If Err.Number = 53 Then Exit Sub
    On Error GoTo Error:
    
    'Reset all menu entries
    For i = 0 To 9
        mnuRASItem(i).Caption = "dummy"
        mnuRASItem(i).Enabled = False
        mnuRASItem(i).Visible = False
    Next i
    
    'Set menu entries according to uREN() RAS entries
    For i = 0 To lEntries - 1
        j = StrConv(uREN(i).szEntryName, vbUnicode)
        j = Left(j, InStr(j, Chr(0)) - 1)
        mnuRASItem(i).Visible = True
        mnuRASItem(i).Enabled = True
        mnuRASItem(i).Caption = j
    Next i
    
    'Set first menu to '(empty)' when no RAS entries found
    If lEntries = 0 Then
        mnuRASItem(0).Visible = True
        mnuRASItem(0).Enabled = False
        mnuRASItem(0).Caption = "(empty)"
        mnuRASItem(0).Checked = False
    Else
        'Get default RAS connection from Registry
        Dim sRASDef$, iRasDef%
        iRasDef = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "RASConnection", 0))
        If frmMenu.mnuRASItem(iRasDef).Visible = True Then
            If frmMenu.mnuRASItem(iRasDef).Caption <> "(empty)" Then frmMenu.mnuRASItem(iRasDef).Checked = True
        Else
            sRASDef = RegGetString(HKEY_CURRENT_USER, "RemoteAccess", "Default")
            If sRASDef <> "" Then
                For i = 0 To 9
                    If mnuRASItem(0).Caption = sRASDef Then iRasDef = i
                Next i
                'if default conn from Registry not found, just
                'check first entry if it's a connection
                If mnuRASItem(iRasDef).Caption <> "(empty)" Then
                    mnuRASItem(iRasDef).Checked = True
                    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "RASConnection", CLng(iRasDef)
                End If
            End If
        End If
    End If
    Exit Sub
    
Error:
    ShowError "RAS Connection", "frmMenu_EnumRASConn", Err.Number, Err.Description, False
End Sub

Private Sub timServerTimeOut_Timer()
    MsgBox "Unable to connect to server " & timServerTimeOut.Tag & "! Maybe the server is down?", vbExclamation, "blah"
    timServerTimeOut.Tag = ""
    timServerTimeOut.Enabled = False
End Sub
