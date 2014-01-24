VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uptimer 4"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   HelpContextID   =   1000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraNetstat 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Netstat"
      Height          =   255
      Left            =   120
      TabIndex        =   57
      Top             =   8280
      Width           =   2295
      Begin VB.Timer timNetstat 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2280
         Top             =   0
      End
      Begin VB.PictureBox picNetstat 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":0442
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   58
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox picNetstat 
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   0
         Picture         =   "frmMain.frx":058C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblNetstat 
         BackStyle       =   0  'Transparent
         Caption         =   "00 netstat items"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   59
         Top             =   0
         Width           =   1920
      End
   End
   Begin VB.Frame fraRAS 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "RAS Connections"
      Height          =   255
      Left            =   2400
      TabIndex        =   56
      Top             =   4800
      Width           =   240
      Begin VB.Timer timRAS 
         Interval        =   1000
         Left            =   1080
         Top             =   0
      End
      Begin VB.Image imgRAS 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "frmMain.frx":06D6
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgRAS 
         Height          =   240
         Index           =   2
         Left            =   600
         Picture         =   "frmMain.frx":0820
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgRAS 
         Height          =   240
         Index           =   1
         Left            =   360
         Picture         =   "frmMain.frx":096A
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgRAS 
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":0AB4
         ToolTipText     =   "RAS connection - offline"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraVolume 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Master Volume"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
      Begin VB.CommandButton hscVolume 
         Height          =   255
         Left            =   720
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   135
      End
      Begin VB.Frame fraVolumeBar 
         BackColor       =   &H80000002&
         Height          =   30
         Left            =   720
         TabIndex        =   39
         Top             =   100
         Width           =   1575
      End
      Begin VB.Timer timVolume 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2280
         Top             =   0
      End
      Begin VB.Label lblVolumePerc 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   320
         TabIndex        =   43
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgVolume 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":0BFE
         Stretch         =   -1  'True
         ToolTipText     =   "Master volume"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgVolumeMute 
         Height          =   240
         Index           =   1
         Left            =   0
         Picture         =   "frmMain.frx":0D48
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgVolumeMute 
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":0E92
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame fraMSIE 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Internet Explorer version"
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   6960
      Width           =   2055
      Begin VB.Label lblMSIE 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00.0000.0000"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   52
         Top             =   0
         Width           =   1680
      End
      Begin VB.Image imgMSIE 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":0FDC
         ToolTipText     =   "Internet Explorer version"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraTCPMonitor 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "TCP Monitor"
      Height          =   495
      Left            =   120
      TabIndex        =   45
      Top             =   7680
      Width           =   2415
      Begin VB.PictureBox picGraphTCP 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1560
         ScaleHeight     =   495
         ScaleWidth      =   855
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Timer timTCPMonitor 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2640
         Top             =   0
      End
      Begin VB.Image imgTCPMonitorArrow 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmMain.frx":1566
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgTCPMonitorArrow 
         Height          =   240
         Index           =   0
         Left            =   360
         Picture         =   "frmMain.frx":16B0
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblTCPMonitorUp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0,00 K/s"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   47
         Top             =   240
         Width           =   960
      End
      Begin VB.Label lblTCPMonitorDown 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0,00 K/s"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   600
         TabIndex        =   46
         Top             =   0
         Width           =   960
      End
      Begin VB.Image imgTCPMonitor 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":17FA
         ToolTipText     =   "TCP/IP traffic"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraProcesses 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "List running processes"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   6600
      Width           =   1815
      Begin VB.Timer timProcesses 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1800
         Top             =   0
      End
      Begin VB.PictureBox picProcesses 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":1944
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   37
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox picProcesses 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   0
         Picture         =   "frmMain.frx":19DE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblProcesses 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00 processes"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   38
         Top             =   0
         Width           =   1440
      End
   End
   Begin VB.Frame fraMouseIdle 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Mouse idle time"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   6240
      Width           =   975
      Begin VB.Timer timMouseIdle 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   960
         Top             =   0
      End
      Begin VB.Label lblMouseIdle 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   35
         Top             =   0
         Width           =   600
      End
      Begin VB.Image imgMouseIdle 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":1D80
         ToolTipText     =   "Mouse idle time"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraLock 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Lock"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4800
      Width           =   255
      Begin VB.Image imgLock 
         Height          =   240
         Index           =   1
         Left            =   360
         Picture         =   "frmMain.frx":230A
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgLock 
         Height          =   240
         Index           =   2
         Left            =   720
         Picture         =   "frmMain.frx":2454
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgLock 
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":259E
         Stretch         =   -1  'True
         ToolTipText     =   "Lock screen"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraExitWin 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Exit Windows"
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   4800
      Width           =   1695
      Begin VB.Image imgExitWin 
         Height          =   240
         Index           =   4
         Left            =   1440
         Picture         =   "frmMain.frx":26E8
         Stretch         =   -1  'True
         ToolTipText     =   "Poweroff"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgExitWin 
         Height          =   240
         Index           =   3
         Left            =   1080
         Picture         =   "frmMain.frx":2AA4
         Stretch         =   -1  'True
         ToolTipText     =   "Shutdown"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgExitWin 
         Height          =   240
         Index           =   2
         Left            =   720
         Picture         =   "frmMain.frx":2E65
         Stretch         =   -1  'True
         ToolTipText     =   "Suspend"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgExitWin 
         Height          =   240
         Index           =   1
         Left            =   360
         Picture         =   "frmMain.frx":3232
         Stretch         =   -1  'True
         ToolTipText     =   "Reboot"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgExitWin 
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":3600
         Stretch         =   -1  'True
         ToolTipText     =   "Logoff"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraPower 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Power status"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5880
      Width           =   975
      Begin VB.Timer timPower 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   960
         Top             =   120
      End
      Begin VB.Image imgPower 
         Height          =   240
         Index           =   7
         Left            =   1440
         Picture         =   "frmMain.frx":39B2
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPower 
         Height          =   240
         Index           =   6
         Left            =   1200
         Picture         =   "frmMain.frx":3AFC
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPower 
         Height          =   240
         Index           =   5
         Left            =   960
         Picture         =   "frmMain.frx":3C46
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPower 
         Height          =   240
         Index           =   4
         Left            =   720
         Picture         =   "frmMain.frx":3D90
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPower 
         Height          =   240
         Index           =   3
         Left            =   480
         Picture         =   "frmMain.frx":3EDA
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPower 
         Height          =   240
         Index           =   2
         Left            =   240
         Picture         =   "frmMain.frx":4024
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPower 
         Height          =   240
         Index           =   1
         Left            =   0
         Picture         =   "frmMain.frx":416E
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblPower 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   30
         Top             =   0
         Width           =   600
      End
      Begin VB.Image imgPower 
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":42B8
         Stretch         =   -1  'True
         ToolTipText     =   "Power status"
         Top             =   0
         Width           =   240
      End
      Begin VB.Shape shpPowerFore 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape shpPowerBack 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.Frame fraToggle 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Toggle keys"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5520
      Width           =   2220
      Begin VB.Timer timToggle 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2160
         Top             =   0
      End
      Begin VB.Label lblToggle 
         BackColor       =   &H00800080&
         Caption         =   "Ins"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Index           =   3
         Left            =   1860
         TabIndex        =   28
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lblToggle 
         BackColor       =   &H00800080&
         Caption         =   "Scrl"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Index           =   2
         Left            =   1320
         TabIndex        =   27
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblToggle 
         BackColor       =   &H00800080&
         Caption         =   "Caps"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Index           =   1
         Left            =   780
         TabIndex        =   26
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lblToggle 
         BackColor       =   &H00800080&
         Caption         =   "Num"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   25
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgToggle 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":4680
         ToolTipText     =   "Toggle keys"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraBanner 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Uptimer 3.5"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1215
      Begin VB.Label lblBanner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uptimer"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   23
         Top             =   0
         Width           =   840
      End
      Begin VB.Image imgBanner 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":47CA
         ToolTipText     =   "Uptimer4"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraOS 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Operating system"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5160
      Width           =   2175
      Begin VB.Label lblOS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Win00 0.00.0000"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   21
         Top             =   0
         Width           =   1800
      End
      Begin VB.Image imgOS 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":4914
         ToolTipText     =   "Windows version"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraResolution 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Resolution"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   2295
      Begin VB.Label lblResolution 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0000 x 0000 x 00"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   33
         Top             =   0
         Width           =   1935
      End
      Begin VB.Image imgResolution 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":4A5E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraWinamp 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Winamp controls"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   2055
      Begin VB.Timer timWinamp 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   2040
         Top             =   0
      End
      Begin VB.Image imgWinampC 
         Height          =   240
         Index           =   6
         Left            =   1800
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":4BA8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgWinampC 
         Height          =   240
         Index           =   5
         Left            =   1560
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":4CF2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgWinampC 
         Height          =   240
         Index           =   4
         Left            =   1320
         Picture         =   "frmMain.frx":4E3C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgWinampC 
         Height          =   240
         Index           =   3
         Left            =   1080
         Picture         =   "frmMain.frx":4F86
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgWinampC 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmMain.frx":50D0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgWinampC 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":521A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgWinampC 
         Height          =   240
         Index           =   0
         Left            =   360
         Picture         =   "frmMain.frx":5364
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgWinamp 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":54AE
         Stretch         =   -1  'True
         ToolTipText     =   "WinAmp controls"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraMemoryPage 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Pagefile"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   2655
      Begin VB.Timer timMemoryPage 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   1800
         Top             =   0
      End
      Begin VB.PictureBox picGraphPage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   855
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblMemoryPage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0000/0000 MB"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   17
         Top             =   0
         Width           =   1440
      End
      Begin VB.Image imgMemoryPage 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":585F
         Stretch         =   -1  'True
         ToolTipText     =   "Free pagefile"
         Top             =   0
         Width           =   240
      End
      Begin VB.Shape shpMemoryPageFore 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   615
      End
      Begin VB.Shape shpMemoryPageBack 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame fraTime 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Time"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1335
      Begin VB.Timer timTime 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1320
         Top             =   0
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   15
         Top             =   0
         Width           =   960
      End
      Begin VB.Image imgTime 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":59A9
         Stretch         =   -1  'True
         ToolTipText     =   "Time"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraCPU 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "CPU Usage"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
      Begin VB.Timer timCPU 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   960
         Top             =   0
      End
      Begin VB.PictureBox picGraphCPU 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   855
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblCPU 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   13
         Top             =   0
         Width           =   600
      End
      Begin VB.Image imgCPU 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":5AF3
         Stretch         =   -1  'True
         ToolTipText     =   "CPU usage"
         Top             =   0
         Width           =   240
      End
      Begin VB.Shape shpCPUFore 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape shpCPUBack 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame fraIPs 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "IP Addresses"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   2175
      Begin VB.Timer timIPs 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   2160
         Top             =   0
      End
      Begin VB.Label lblIPs 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "255.255.255.255"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   12
         Top             =   0
         Width           =   1815
      End
      Begin VB.Image imgIPs 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":5EBE
         Stretch         =   -1  'True
         ToolTipText     =   "IP addresses"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraVolume2 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "CD Player Volume"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   2295
      Begin VB.CommandButton hscVolume2 
         Height          =   255
         Left            =   720
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   0
         Width           =   135
      End
      Begin VB.Frame fraVolume2Bar 
         BackColor       =   &H80000002&
         Height          =   30
         Left            =   720
         TabIndex        =   40
         Top             =   120
         Width           =   1575
      End
      Begin VB.Timer timVolume2 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2280
         Top             =   0
      End
      Begin VB.Label lblVolume2Perc 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   320
         TabIndex        =   44
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgVolume2 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":6008
         Stretch         =   -1  'True
         ToolTipText     =   "CD volume"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgVolume2Mute 
         Height          =   240
         Index           =   1
         Left            =   0
         Picture         =   "frmMain.frx":63DD
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgVolume2Mute 
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":6967
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame fraDiskFree 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Disk Free Space"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   2415
      Begin VB.Timer timDiskFreeSpace 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   2400
         Top             =   0
      End
      Begin VB.Label lblDiskFreeSpace 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00,00 GB/00,00 GB"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   7
         Top             =   0
         Width           =   2040
      End
      Begin VB.Image imgDiskFreeSpace 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":6EF1
         Stretch         =   -1  'True
         ToolTipText     =   "Disk free space"
         Top             =   0
         Width           =   240
      End
      Begin VB.Shape shpDiskFreeFore 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   735
      End
      Begin VB.Shape shpDiskFreeBack 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame fraMemoryRAM 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Free RAM && Pagefile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2655
      Begin VB.Timer timMemoryRAM 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   1800
         Top             =   0
      End
      Begin VB.PictureBox picGraphRAM 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   855
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblMemoryRAM 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000/0000 MB"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   5
         Top             =   0
         Width           =   1440
      End
      Begin VB.Image imgMemoryRAM 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":7278
         Stretch         =   -1  'True
         ToolTipText     =   "Free RAM"
         Top             =   0
         Width           =   240
      End
      Begin VB.Shape shpMemoryRAMFore 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   615
      End
      Begin VB.Shape shpMemoryRAMBack 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1680
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   2
         Top             =   0
         Width           =   1200
      End
      Begin VB.Image imgDate 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":7611
         Stretch         =   -1  'True
         ToolTipText     =   "Date"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame fraUptime 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Uptime"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1695
      Begin VB.Timer timUptime 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1680
         Top             =   0
      End
      Begin VB.Label lblUptime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00:00"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   360
         TabIndex        =   3
         Top             =   0
         Width           =   1320
      End
      Begin VB.Image imgUptime 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":79BD
         Stretch         =   -1  'True
         ToolTipText     =   "Uptime"
         Top             =   0
         Width           =   240
      End
      Begin VB.Shape shpUptimeFore 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   615
      End
      Begin VB.Shape shpUptimeBack 
         BackColor       =   &H00800080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   360
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame fraDX 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "DirectX version"
      Height          =   255
      Left            =   120
      TabIndex        =   53
      Top             =   7320
      Width           =   1935
      Begin VB.Label lblDX 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0.00.00.0000"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   225
         Left            =   300
         TabIndex        =   54
         Top             =   0
         Width           =   1560
      End
      Begin VB.Image imgDX 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":7D72
         ToolTipText     =   "DirectX version"
         Top             =   0
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub GetModules(bReadSave As Boolean)
    On Error GoTo Error:
    'bReadSave = True  --> Read modules and apply
    'bReadSave = False --> Save modules to Registry
    
    'Possible modules:
    '1  - CD player volume [disabled, remove ' in frmModules_Load to enable]
    '2  - CPU usage
    '3  - Date
    '4  - Disk free space
    '5  - Exit Windows
    '6  - Free pagefile
    '7  - Free RAM
    '8  - IP addresses
    '9  - Lock screen
    '10 - Master volume
    '11 - Power status
    '12 - Screen resolution
    '13 - Time
    '14 - Toggle keys status
    '15 - Uptime
    '16 - WinAmp controls
    '17 - Windows version
    '[Add more modules below]
    '18 - List running processes
    '19 - Mouse idle time
    '20 - TCP Monitor
    '21 - MSIE version
    '22 - DirectX version
    '23 - RAS connection
    '24 - Netstat
    
    If bReadSave Then GoTo ReadModules:
    
    'Save modules should only be called from frmModules
    Dim i%
    sModules = ""
    With frmModules.lstUsed
        For i = 0 To .ListCount - 1
            Select Case .List(i)
                Case "CD player volume":   sModules = sModules & "1,"
                Case "CPU usage":          sModules = sModules & "2,"
                Case "Date":               sModules = sModules & "3,"
                Case "Disk free space":    sModules = sModules & "4,"
                Case "Exit Windows":       sModules = sModules & "5,"
                Case "Free pagefile":      sModules = sModules & "6,"
                Case "Free RAM":           sModules = sModules & "7,"
                Case "IP addresses":       sModules = sModules & "8,"
                Case "Lock screen":        sModules = sModules & "9,"
                Case "Master volume":      sModules = sModules & "10,"
                Case "Power status":       sModules = sModules & "11,"
                Case "Screen resolution":  sModules = sModules & "12,"
                Case "Time":               sModules = sModules & "13,"
                Case "Toggle keys status": sModules = sModules & "14,"
                Case "Uptime":             sModules = sModules & "15,"
                Case "WinAmp controls":    sModules = sModules & "16,"
                Case "Windows version":    sModules = sModules & "17,"
                'Add more modules below
                Case "List running processes": sModules = sModules & "18,"
                Case "Mouse idle time":    sModules = sModules & "19,"
                Case "TCP Monitor":        sModules = sModules & "20,"
                Case "MSIE version":       sModules = sModules & "21,"
                Case "DirectX version":    sModules = sModules & "22,"
                Case "RAS connection":     sModules = sModules & "23,"
                Case "Netstat":            sModules = sModules & "24,"
            End Select
        Next i
    End With
    sModules = Left(sModules, Len(sModules) - 1)
    RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "Modules", sModules
    Exit Sub
    
ReadModules:
    fraBanner.Visible = True
    fraCPU.Visible = False
    timCPU.Enabled = False
    fraDate.Visible = False
    fraDiskFree.Visible = False
    timDiskFreeSpace.Enabled = False
    fraExitWin.Visible = False
    fraIPs.Visible = False
    timIPs.Enabled = False
    fraLock.Visible = False
    fraMemoryPage.Visible = False
    timMemoryPage.Enabled = False
    fraMemoryRAM.Visible = False
    timMemoryRAM.Enabled = False
    fraOS.Visible = False
    fraPower.Visible = False
    timPower.Enabled = False
    fraResolution.Visible = False
    fraTime.Visible = False
    timTime.Enabled = False
    fraToggle.Visible = False
    timToggle.Enabled = False
    fraUptime.Visible = False
    timUptime.Enabled = False
    fraVolume.Visible = False
    timVolume.Enabled = False
    fraVolume2.Visible = False
    timVolume2.Enabled = False
    fraWinamp.Visible = False
    timWinamp.Enabled = False
    'Add more modules below
    fraProcesses.Visible = False
    fraMouseIdle.Visible = False
    timMouseIdle.Enabled = False
    fraTCPMonitor.Visible = False
    timTCPMonitor.Enabled = False
    fraMSIE.Visible = False
    fraDX.Visible = False
    fraRAS.Visible = False
    timRAS.Enabled = False
    fraNetstat.Visible = False
    
    Dim vDisplay As Variant, sDefModules$
    
    'Default settings, used on first run:
    sModules = "15,13,3,7,6" '(Uptime, Time, Date, Free RAM, Free pagefile)
    
    sModules = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "Modules", sModules)
    vDisplay = Split(sModules, ",")
    For i = 0 To UBound(vDisplay)
        Select Case vDisplay(i)
            Case 1:  fraVolume2.Visible = True:    timVolume2.Enabled = True
            Case 2:  fraCPU.Visible = True:        timCPU.Enabled = True
            Case 3:  fraDate.Visible = True
            Case 4:  fraDiskFree.Visible = True:   timDiskFreeSpace.Enabled = True
            Case 5:  fraExitWin.Visible = True
            Case 6:  fraMemoryPage.Visible = True: timMemoryPage.Enabled = True
            Case 7:  fraMemoryRAM.Visible = True:  timMemoryRAM.Enabled = True
            Case 8:  fraIPs.Visible = True:        timIPs.Enabled = True
            Case 9:  fraLock.Visible = True
            Case 10: fraVolume.Visible = True:     timVolume.Enabled = True
            Case 11: fraPower.Visible = True:      timPower.Enabled = True
            Case 12: fraResolution.Visible = True
            Case 13: fraTime.Visible = True:       timTime.Enabled = True
            Case 14: fraToggle.Visible = True:     timToggle.Enabled = True
            Case 15: fraUptime.Visible = True:     timUptime.Enabled = True
            Case 16: fraWinamp.Visible = True:     timWinamp.Enabled = True
            Case 17: fraOS.Visible = True
            'Add more modules below
            Case 18: fraProcesses.Visible = True
            Case 19: fraMouseIdle.Visible = True:  timMouseIdle.Enabled = True
            Case 20: fraTCPMonitor.Visible = True: timTCPMonitor.Enabled = True
            Case 21: fraMSIE.Visible = True
            Case 22: fraDX.Visible = True
            Case 23: fraRAS.Visible = True:        timRAS.Enabled = True
            Case 24: fraNetstat.Visible = True
        End Select
    Next i
    Exit Sub
    
Error:
    ShowError "Main", "frmMain.GetModules", Err.Number, Err.Description, False
End Sub

Public Sub GetColors()
    lColorText = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ColorText", 16777215)
    lColorFore = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ColorFore", 16711680)
    lColorBack = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ColorBack", 8388736)
    lColorGraphGrid = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ColorGraphGrid", 8421504)
    lColorGraph1st = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ColorGraph1st", 16776960)
    lColorGraph2nd = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ColorGraph2nd", 255)
    
    lblBanner.ForeColor = lColorText
    lblCPU.ForeColor = lColorText
    lblDate.ForeColor = lColorText
    lblDiskFreeSpace.ForeColor = lColorText
    lblIPs.ForeColor = lColorText
    lblMemoryPage.ForeColor = lColorText
    lblMemoryRAM.ForeColor = lColorText
    lblMouseIdle.ForeColor = lColorText
    lblOS.ForeColor = lColorText
    lblPower.ForeColor = lColorText
    lblResolution.ForeColor = lColorText
    lblTime.ForeColor = lColorText
    lblToggle(0).ForeColor = lColorText
    lblToggle(1).ForeColor = lColorText
    lblToggle(2).ForeColor = lColorText
    lblToggle(3).ForeColor = lColorText
    lblUptime.ForeColor = lColorText
    
    shpCPUFore.BackColor = lColorFore
    shpDiskFreeFore.BackColor = lColorFore
    shpMemoryPageFore.BackColor = lColorFore
    shpMemoryRAMFore.BackColor = lColorFore
    shpPowerFore.BackColor = lColorFore
    shpUptimeFore.BackColor = lColorFore
    
    shpCPUBack.BackColor = lColorBack
    shpDiskFreeBack.BackColor = lColorBack
    shpMemoryPageBack.BackColor = lColorBack
    shpMemoryRAMBack.BackColor = lColorBack
    shpPowerBack.BackColor = lColorBack
    shpUptimeBack.BackColor = lColorBack

    'Add more modules below
    lblProcesses.ForeColor = lColorText
    lblMouseIdle.ForeColor = lColorText
    lblTCPMonitorDown.ForeColor = lColorText
    lblTCPMonitorUp.ForeColor = lColorText
    lblMSIE.ForeColor = lColorText
    lblDX.ForeColor = lColorText
    lblNetstat.ForeColor = lColorText
End Sub

Public Sub GetSettings()
    'Retrieves settings for each module and main app
    Dim bDummy As Boolean, iDummy%, lDummy&, sDummy$
    On Error GoTo Error:

    GetWinVersion True, False

    'Main app
    bBarPos = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "BarPos", 1))
    bBarDocking = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "BarDocking", 1))
    bBarHidden = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "BarHidden", 0))
    bAlwaysOnTop = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlwaysOnTop", 1))
    bEnableMultiRows = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "Multirows", 0))
    bIgnoreFullscreenApps = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "IgnoreFullscreenApps", 0))
    bSolidGraphs = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "SolidGraphs", 0))
    bCoolSunkenButtons = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "CoolSunkenButtons", 1))
    iTransparency = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "Transparency", 0))
    
    sWinDir = String(255, 0)
    GetWindowsDirectory sWinDir, 255
    sWinDir = Left(sWinDir, InStr(sWinDir, Chr(0)) - 1)
    If bIsWinNT Then
        sWinSysDir = sWinDir & "\system32"
    Else
        sWinSysDir = sWinDir & "\system"
    End If
    
    If CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TrayIcon", 0)) Then MakeTrayIcon True, frmMenu.picTrayIcon.hwnd
    iDummy = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "WebserverAdapter", 0))
    timIPs_Timer
    lDummy = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "Webserver", 0)
    If lDummy Then Webserver True, sIPs(iDummy), lDummy, frmMenu.picWebserverAccept.hwnd
    
    
    '1 - CD player volume
    
    
    '2 - CPU usage
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "CPUShowBar", 1))
    frmMenu.mnuCPUBar.Checked = bDummy
    shpCPUBack.Visible = bDummy
    shpCPUFore.Visible = bDummy
    timCPU.Interval = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "CPUInterval", 1000)
    iDummy = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "CPUDisplay", 0))
    frmMenu.mnuCPUDisplay(iDummy).Checked = True
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "CPUGraph", 0))
    frmMenu.mnuCPUGraph.Checked = bDummy
    picGraphCPU.Visible = bDummy
    fraCPU.Width = picGraphCPU.Left + IIf(bDummy, picGraphCPU.Width, 0)
    
    
    '3 - Date
    lblDate.Caption = Format(Date, "dd/mm/yyyy")
    lblDate.ToolTipText = Format(Date, "Long date")
    
    
    '4 - Disk free space
    ReDim sDisks(0)
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "DiskFreeShowBar", 1))
    frmMenu.mnuDiskFreeSpaceBar.Checked = bDummy
    shpDiskFreeBack.Visible = bDummy
    shpDiskFreeFore.Visible = bDummy
    timDiskFreeSpace.Interval = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "DiskFreeSpaceInterval", 5000)
    sCurrentDisk = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "DiskFreeDisk", ".")
    If sCurrentDisk = "." Then
        frmMenu.mnuDiskFreeSpaceAll.Checked = True
    Else
        frmMenu.mnuDiskFreeSpaceA(Asc(sCurrentDisk) - 65).Checked = True
    End If
    If Not bIsWinNT Then frmMenu.mnuDiskFreeSpaceQuota.Enabled = False
    
    
    '5 - Exit Windows
    frmMenu.mnuExitWinConfirm.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ExitWinConfirm", 1))
    frmMenu.mnuExitWinForce.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ExitWinForce", 0))
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "ExitWinButtons", "11111")
    If sDummy = "00000" Then sDummy = "11111"
    lDummy = 0
    For iDummy = 0 To 4
        If Mid(sDummy, iDummy + 1, 1) = "0" Then
            frmMenu.mnuExitWinButtons(iDummy).Checked = False
            imgExitWin(iDummy).Visible = False
        Else
            frmMenu.mnuExitWinButtons(iDummy).Checked = True
            imgExitWin(iDummy).Visible = True
            imgExitWin(iDummy).Left = lDummy
            lDummy = lDummy + 360
        End If
        fraExitWin.Width = lDummy - 120
    Next iDummy
    
    
    '6 - Free pagefile
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MemoryPageShowBar", 1))
    frmMenu.mnuMemoryPageBar.Checked = bDummy
    shpMemoryPageBack.Visible = bDummy
    shpMemoryPageFore.Visible = bDummy
    timMemoryPage.Interval = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MemoryPageInterval", 5000)
    iDummy = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MemoryPageDisplay", 0))
    frmMenu.mnuMemoryPageDisplay(iDummy).Checked = True
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MemoryPageGraph", 0))
    frmMenu.mnuMemoryPageGraph.Checked = bDummy
    picGraphPage.Visible = bDummy
    fraMemoryPage.Width = picGraphPage.Left + IIf(bDummy, picGraphPage.Width, 0)
    
    
    '7 - Free RAM
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MemoryRAMShowBar", 1))
    frmMenu.mnuMemoryRAMBar.Checked = bDummy
    shpMemoryRAMBack.Visible = bDummy
    shpMemoryRAMFore.Visible = bDummy
    timMemoryRAM.Interval = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MemoryRAMInterval", 2000)
    iDummy = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MemoryRAMDisplay", 0))
    frmMenu.mnuMemoryRAMDisplay(iDummy).Checked = True
    Dim vDummy As Variant
    vDummy = Split(RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "MemoryRAMFreeMemory", "1,5,10,20"), ",")
    With frmMenu
        .mnuMemoryRamFree(0).Caption = "Free up " & CStr(vDummy(0)) & " MB"
        .mnuMemoryRamFree(1).Caption = "Free up " & CStr(vDummy(1)) & " MB"
        .mnuMemoryRamFree(2).Caption = "Free up " & CStr(vDummy(2)) & " MB"
        .mnuMemoryRamFree(3).Caption = "Free up " & CStr(vDummy(3)) & " MB"
    End With
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MemoryRAMGraph", 0))
    frmMenu.mnuMemoryRAMGraph.Checked = bDummy
    picGraphRAM.Visible = bDummy
    fraMemoryRAM.Width = picGraphRAM.Left + IIf(bDummy, picGraphRAM.Width, 0)
    
    
    '8 - IP addresses
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "IPsIgnoreLocalhost", 1))
    bIgnoreLocalHostIP = bDummy
    frmMenu.mnuIPsIgnoreLocalhost.Checked = bDummy
    
    
    '9 - Lock screen
    
    
    '10 - Master volume
    
    
    '11 - Power status
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "PowerShowBar", 1))
    frmMenu.mnuPowerBar.Checked = bDummy
    shpPowerBack.Visible = bDummy
    shpPowerFore.Visible = bDummy
    timPower.Interval = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "PowerInterval", 10000)
    frmMenu.mnuPowerBatChargeOnAC.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "PowerBattChargeOnAC", 0))
    
    
    '12 - Screen resolution
    GetCurrentDispMode
    GetDispModes
    frmMenu.mnuResConfirm.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ResolutionConfirm", 1))
    
    
    '13 - Time
    If RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmIsSet", 0) = 1 Then
        bReminderSet = True
        sReminderTime = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTime")
        If sReminderTime = "" Then bReminderSet = False
    End If

    
    '14 - Toggle keys status
    
    
    '15 - Uptime
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "UptimeShowBar", 0))
    frmMenu.mnuUptimeBar.Checked = bDummy
    shpUptimeBack.Visible = bDummy
    shpUptimeFore.Visible = bDummy
    frmMenu.mnuUptimeLoggingWriteHourly.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "UptimeLogHourly", 0))
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "UptimeEnableLogging", 0))
    With frmMenu
        If bDummy Then
            .mnuUptimeLoggingEnable.Checked = True
            sUptimeLogLocation = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Uptimer4.log"
        Else
            .mnuUptimeLoggingWriteHourly.Enabled = False
            .mnuUptimeLoggingWriteNow.Enabled = False
            .mnuUptimeLoggingView.Enabled = False
            .mnuUptimeLoggingCleanUp.Enabled = False
            .mnuUptimeLoggingClear.Enabled = False
        End If
    End With
    
    
    '16 - WinAmp controls
    frmMenu.mnuWinAmpMin.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "WinAmpMin", 0))
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "WinampHotkeys", 0))
    iDummy = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "WinampHotkeysMode", 3))
    frmMenu.mnuWinampHotkeys.Checked = bDummy
    If bDummy And Not RunningInIDE Then RegisterHotkeys True, iDummy
    If bDummy Then
        frmMenu.mnuWinampHotkeyMode(1).Enabled = True
        frmMenu.mnuWinampHotkeyMode(2).Enabled = True
        frmMenu.mnuWinampHotkeyMode(3).Enabled = True
        frmMenu.mnuWinampHotkeyMode(iDummy).Checked = True
    Else
        frmMenu.mnuWinampHotkeyMode(1).Enabled = False
        frmMenu.mnuWinampHotkeyMode(2).Enabled = False
        frmMenu.mnuWinampHotkeyMode(3).Enabled = False
    End If
    timWinamp_Timer
    
    
    '17 - Windows version
    If Not bIsWinNT Then frmMenu.mnuOSShowBuildSP(1).Enabled = False
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "OSFriendly", 0))
    frmMenu.mnuOSFriendly.Checked = bDummy
    If bIsWinNT Then
        GetWinVersion bDummy, CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "OSShowSP", 0))
    Else
        GetWinVersion bDummy, False
    End If
    
    '18 - List running processes
    iProcessTrunc = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ProcessesTruncate", 0))
    
    
    '19 - Mouse idle time
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MouseIdleUseTimeout", 0))
    frmMenu.mnuMouseIdleUseTimeout.Checked = bDummy
    frmMenu.mnuMouseIdleSetTimeout.Enabled = bDummy
    If bDummy Then iMouseTimeout = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MouseIdleTimeout", 180))
    
    
    '20 - TCP Monitor
    iDummy = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TCPMonitorAdapter", 0))
    If iDummy = 0 Then
        frmMenu.mnuTCPMonitorAll.Checked = True
        frmMenu.mnuTCPMonitorIgnoreLoopback.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TCPMonitorIgnoreLoopback", 1))
        frmMenu.mnuTCPMonitorInfo.Enabled = False
    Else
        frmMenu.mnuTCPMonitorAdapter(iDummy).Checked = True
        frmMenu.mnuTCPMonitorIgnoreLoopback.Enabled = False
    End If
    bDummy = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TCPMonitorGraph", 0))
    frmMenu.mnuTCPMonitorGraph.Checked = bDummy
    picGraphTCP.Visible = bDummy
    fraTCPMonitor.Width = picGraphTCP.Left + IIf(bDummy, picGraphTCP.Width, 0)
    If bDummy Then DrawGrid picGraphTCP

        
    '21 - MSIE version
    GetMSIEVersion
    
    
    '22 - DirectX version
    GetDXVersion
    
    
    '23 - RAS connection
    frmMenu.EnumRASConn
    
    
    '24 - Netstat
    frmMenu.mnuNetstatShowUDP.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "NetstatShowUDP", 1))
    frmMenu.mnuNetstatShowTimewait.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "NetstatShowTimewait", 0))
    frmMenu.mnuNetstatShowListening.Checked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "NetstatShowListening", 1))
    iNetstatTrunc = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "NetstatTruncate", 0))
    
    
    Exit Sub
Error:
    ShowError "Main", "frmMain.GetSettings", Err.Number, Err.Description, False
End Sub

Public Sub GetDXVersion()
    Dim sDX$
    sDX = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\DirectX", "Version")
    If sDX = "" Then sDX = "not installed"
    lblDX.Caption = sDX
End Sub

Public Sub GetMSIEVersion()
    Dim sMSIE$, sMinor$
    sMSIE = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer", "Version")
    sMinor = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MinorVersion")
    If sMSIE = "" Then
        lblMSIE.Caption = "not installed"
        lblMSIE.ToolTipText = ""
    Else
        lblMSIE.Caption = sMSIE
        sMinor = "Internet Explorer " & Left(sMSIE, 4) & sMinor
        lblMSIE.ToolTipText = IIf(Right(sMinor, 1) = ";", Left(sMinor, Len(sMinor) - 1), sMinor)
    End If
End Sub

Public Sub GetCustomIcons()
    Dim sPath$, sCustomImgs$(50)
    On Error GoTo Error:
    sPath = App.Path & IIf(Right(App.Path, 1) = "\", "icons\", "\icons\")
    On Error Resume Next
    
    If Dir(sPath & "cdvolume.ico") <> "" Then sCustomImgs(0) = sPath & "cdvolume.ico"
    If Dir(sPath & "cdvolume.gif") <> "" Then sCustomImgs(0) = sPath & "cdvolume.gif"
    If Dir(sPath & "cpuusage.ico") <> "" Then sCustomImgs(1) = sPath & "cpuusage.ico"
    If Dir(sPath & "cpuusage.gif") <> "" Then sCustomImgs(1) = sPath & "cpuusage.gif"
    If Dir(sPath & "date.ico") <> "" Then sCustomImgs(2) = sPath & "date.ico"
    If Dir(sPath & "date.gif") <> "" Then sCustomImgs(2) = sPath & "date.gif"
    If Dir(sPath & "diskfree.ico") <> "" Then sCustomImgs(3) = sPath & "diskfree.ico"
    If Dir(sPath & "diskfree.gif") <> "" Then sCustomImgs(3) = sPath & "diskfree.gif"
    If Dir(sPath & "exitwin1.ico") <> "" Then sCustomImgs(4) = sPath & "exitwin1.ico"
    If Dir(sPath & "exitwin1.gif") <> "" Then sCustomImgs(4) = sPath & "exitwin1.gif"
    If Dir(sPath & "exitwin2.ico") <> "" Then sCustomImgs(5) = sPath & "exitwin2.ico"
    If Dir(sPath & "exitwin2.gif") <> "" Then sCustomImgs(5) = sPath & "exitwin2.gif"
    If Dir(sPath & "exitwin3.ico") <> "" Then sCustomImgs(6) = sPath & "exitwin3.ico"
    If Dir(sPath & "exitwin3.gif") <> "" Then sCustomImgs(6) = sPath & "exitwin3.gif"
    If Dir(sPath & "exitwin4.ico") <> "" Then sCustomImgs(7) = sPath & "exitwin4.ico"
    If Dir(sPath & "exitwin4.gif") <> "" Then sCustomImgs(7) = sPath & "exitwin4.gif"
    If Dir(sPath & "exitwin5.ico") <> "" Then sCustomImgs(8) = sPath & "exitwin5.ico"
    If Dir(sPath & "exitwin5.gif") <> "" Then sCustomImgs(8) = sPath & "exitwin5.gif"
    If Dir(sPath & "freepagefile.ico") <> "" Then sCustomImgs(9) = sPath & "freepagefile.ico"
    If Dir(sPath & "freepagefile.gif") <> "" Then sCustomImgs(9) = sPath & "freepagefile.gif"
    If Dir(sPath & "freeram.ico") <> "" Then sCustomImgs(10) = sPath & "freeram.ico"
    If Dir(sPath & "freeram.gif") <> "" Then sCustomImgs(10) = sPath & "freeram.gif"
    If Dir(sPath & "ips.ico") <> "" Then sCustomImgs(11) = sPath & "ips.ico"
    If Dir(sPath & "ips.gif") <> "" Then sCustomImgs(11) = sPath & "ips.gif"
    If Dir(sPath & "lock1.ico") <> "" Then sCustomImgs(12) = sPath & "lock1.ico"
    If Dir(sPath & "lock1.gif") <> "" Then sCustomImgs(12) = sPath & "lock1.gif"
    If Dir(sPath & "lock2.ico") <> "" Then sCustomImgs(13) = sPath & "lock2.ico"
    If Dir(sPath & "lock2.gif") <> "" Then sCustomImgs(13) = sPath & "lock2.gif"
    If Dir(sPath & "mastervolume.ico") <> "" Then sCustomImgs(14) = sPath & "mastervolume.ico"
    If Dir(sPath & "mastervolume.gif") <> "" Then sCustomImgs(14) = sPath & "mastervolume.gif"
    If Dir(sPath & "power1.ico") <> "" Then sCustomImgs(15) = sPath & "power1.ico"
    If Dir(sPath & "power1.gif") <> "" Then sCustomImgs(15) = sPath & "power1.gif"
    If Dir(sPath & "power2.ico") <> "" Then sCustomImgs(16) = sPath & "power2.ico"
    If Dir(sPath & "power2.gif") <> "" Then sCustomImgs(16) = sPath & "power2.gif"
    If Dir(sPath & "power3.ico") <> "" Then sCustomImgs(17) = sPath & "power3.ico"
    If Dir(sPath & "power3.gif") <> "" Then sCustomImgs(17) = sPath & "power3.gif"
    If Dir(sPath & "power4.ico") <> "" Then sCustomImgs(18) = sPath & "power4.ico"
    If Dir(sPath & "power4.gif") <> "" Then sCustomImgs(18) = sPath & "power4.gif"
    If Dir(sPath & "power5.ico") <> "" Then sCustomImgs(19) = sPath & "power5.ico"
    If Dir(sPath & "power5.gif") <> "" Then sCustomImgs(19) = sPath & "power5.gif"
    If Dir(sPath & "power6.ico") <> "" Then sCustomImgs(20) = sPath & "power6.ico"
    If Dir(sPath & "power6.gif") <> "" Then sCustomImgs(20) = sPath & "power6.gif"
    If Dir(sPath & "power7.ico") <> "" Then sCustomImgs(21) = sPath & "power7.ico"
    If Dir(sPath & "power7.gif") <> "" Then sCustomImgs(21) = sPath & "power7.gif"
    If Dir(sPath & "power8.ico") <> "" Then sCustomImgs(22) = sPath & "power8.ico"
    If Dir(sPath & "power8.gif") <> "" Then sCustomImgs(22) = sPath & "power8.gif"
    If Dir(sPath & "resolution.ico") <> "" Then sCustomImgs(23) = sPath & "resolution.ico"
    If Dir(sPath & "resolution.gif") <> "" Then sCustomImgs(23) = sPath & "resolution.gif"
    If Dir(sPath & "time.ico") <> "" Then sCustomImgs(24) = sPath & "time.ico"
    If Dir(sPath & "time.gif") <> "" Then sCustomImgs(24) = sPath & "time.gif"
    If Dir(sPath & "togglekeys.ico") <> "" Then sCustomImgs(25) = sPath & "togglekeys.ico"
    If Dir(sPath & "togglekeys.gif") <> "" Then sCustomImgs(25) = sPath & "togglekeys.gif"
    If Dir(sPath & "uptime.ico") <> "" Then sCustomImgs(26) = sPath & "uptime.ico"
    If Dir(sPath & "uptime.gif") <> "" Then sCustomImgs(26) = sPath & "uptime.gif"
    If Dir(sPath & "winamp.ico") <> "" Then sCustomImgs(27) = sPath & "winamp.ico"
    If Dir(sPath & "winamp.gif") <> "" Then sCustomImgs(27) = sPath & "winamp.gif"
    If Dir(sPath & "winampc1.ico") <> "" Then sCustomImgs(28) = sPath & "winampc1.ico"
    If Dir(sPath & "winampc1.gif") <> "" Then sCustomImgs(28) = sPath & "winampc1.gif"
    If Dir(sPath & "winampc2.ico") <> "" Then sCustomImgs(29) = sPath & "winampc2.ico"
    If Dir(sPath & "winampc2.gif") <> "" Then sCustomImgs(29) = sPath & "winampc2.gif"
    If Dir(sPath & "winampc3.ico") <> "" Then sCustomImgs(30) = sPath & "winampc3.ico"
    If Dir(sPath & "winampc3.gif") <> "" Then sCustomImgs(30) = sPath & "winampc3.gif"
    If Dir(sPath & "winampc4.ico") <> "" Then sCustomImgs(31) = sPath & "winampc4.ico"
    If Dir(sPath & "winampc4.gif") <> "" Then sCustomImgs(31) = sPath & "winampc4.gif"
    If Dir(sPath & "winampc5.ico") <> "" Then sCustomImgs(32) = sPath & "winampc5.ico"
    If Dir(sPath & "winampc5.gif") <> "" Then sCustomImgs(32) = sPath & "winampc5.gif"
    If Dir(sPath & "winampc6.ico") <> "" Then sCustomImgs(33) = sPath & "winampc6.ico"
    If Dir(sPath & "winampc6.gif") <> "" Then sCustomImgs(33) = sPath & "winampc6.gif"
    If Dir(sPath & "winampc7.ico") <> "" Then sCustomImgs(34) = sPath & "winampc7.ico"
    If Dir(sPath & "winampc7.gif") <> "" Then sCustomImgs(34) = sPath & "winampc7.gif"
    If Dir(sPath & "winversion.ico") <> "" Then sCustomImgs(35) = sPath & "winversion.ico"
    If Dir(sPath & "winversion.gif") <> "" Then sCustomImgs(35) = sPath & "winversion.gif"
    'Add more modules below
    If Dir(sPath & "mouseidle.ico") <> "" Then sCustomImgs(37) = sPath & "mouseidle.ico"
    If Dir(sPath & "mouseidle.gif") <> "" Then sCustomImgs(37) = sPath & "mouseidle.gif"
    If Dir(sPath & "tcpmonitor1.ico") <> "" Then sCustomImgs(38) = sPath & "tcpmonitor1.ico"
    If Dir(sPath & "tcpmonitor1.gif") <> "" Then sCustomImgs(38) = sPath & "tcpmonitor1.gif"
    If Dir(sPath & "tcpmonitor2.ico") <> "" Then sCustomImgs(39) = sPath & "tcpmonitor2.ico"
    If Dir(sPath & "tcpmonitor2.gif") <> "" Then sCustomImgs(39) = sPath & "tcpmonitor2.gif"
    If Dir(sPath & "tcpmonitor3.ico") <> "" Then sCustomImgs(40) = sPath & "tcpmonitor3.ico"
    If Dir(sPath & "tcpmonitor3.gif") <> "" Then sCustomImgs(40) = sPath & "tcpmonitor3.gif"
    If Dir(sPath & "msieversion.ico") <> "" Then sCustomImgs(41) = sPath & "msieversion.ico"
    If Dir(sPath & "msieversion.gif") <> "" Then sCustomImgs(41) = sPath & "msieversion.gif"
    If Dir(sPath & "dxversion.ico") <> "" Then sCustomImgs(42) = sPath & "dxversion.ico"
    If Dir(sPath & "dxversion.gif") <> "" Then sCustomImgs(42) = sPath & "dxversion.gif"
    'Me moron, me forgot mute icons on CDPlayer/MasterVolume
    If Dir(sPath & "cdplayermute.ico") <> "" Then sCustomImgs(43) = sPath & "cdplayermute.ico"
    If Dir(sPath & "cdplayermute.gif") <> "" Then sCustomImgs(43) = sPath & "cdplayermute.gif"
    If Dir(sPath & "mastervolumemute.ico") <> "" Then sCustomImgs(44) = sPath & "mastervolumemute.ico"
    If Dir(sPath & "mastervolumemute.gif") <> "" Then sCustomImgs(44) = sPath & "mastervolumemute.gif"
    
    If Dir(sPath & "rasconnection1.ico") <> "" Then sCustomImgs(45) = sPath & "rasconnection1.ico"
    If Dir(sPath & "rasconnection1.gif") <> "" Then sCustomImgs(45) = sPath & "rasconnection1.gif"
    If Dir(sPath & "rasconnection2.ico") <> "" Then sCustomImgs(46) = sPath & "rasconnection2.ico"
    If Dir(sPath & "rasconnection2.gif") <> "" Then sCustomImgs(46) = sPath & "rasconnection2.gif"
    If Dir(sPath & "rasconnection3.ico") <> "" Then sCustomImgs(47) = sPath & "rasconnection3.ico"
    If Dir(sPath & "rasconnection3.gif") <> "" Then sCustomImgs(47) = sPath & "rasconnection4.gif"
    If Dir(sPath & "processes1.ico") <> "" Then sCustomImgs(36) = sPath & "processes1.ico"
    If Dir(sPath & "processes1.gif") <> "" Then sCustomImgs(36) = sPath & "processes1.gif"
    If Dir(sPath & "processes2.ico") <> "" Then sCustomImgs(50) = sPath & "processes2.ico"
    If Dir(sPath & "processes2.gif") <> "" Then sCustomImgs(50) = sPath & "processes2.gif"
    If Dir(sPath & "netstat1.ico") <> "" Then sCustomImgs(48) = sPath & "netstat1.ico"
    If Dir(sPath & "netstat1.gif") <> "" Then sCustomImgs(48) = sPath & "netstat1.gif"
    If Dir(sPath & "netstat2.ico") <> "" Then sCustomImgs(49) = sPath & "netstat2.ico"
    If Dir(sPath & "netstat2.gif") <> "" Then sCustomImgs(49) = sPath & "netstat2.gif"
    
    If sCustomImgs(0) <> "" Then imgVolume2.Picture = LoadPicture(sCustomImgs(0))
    If sCustomImgs(0) <> "" Then imgVolumeMute(1).Picture = LoadPicture(sCustomImgs(0))
    If sCustomImgs(1) <> "" Then imgCPU.Picture = LoadPicture(sCustomImgs(1))
    If sCustomImgs(2) <> "" Then imgDate.Picture = LoadPicture(sCustomImgs(2))
    If sCustomImgs(3) <> "" Then imgDiskFreeSpace.Picture = LoadPicture(sCustomImgs(3))
    If sCustomImgs(4) <> "" Then imgExitWin(0).Picture = LoadPicture(sCustomImgs(4))
    If sCustomImgs(5) <> "" Then imgExitWin(1).Picture = LoadPicture(sCustomImgs(5))
    If sCustomImgs(6) <> "" Then imgExitWin(2).Picture = LoadPicture(sCustomImgs(6))
    If sCustomImgs(7) <> "" Then imgExitWin(3).Picture = LoadPicture(sCustomImgs(7))
    If sCustomImgs(8) <> "" Then imgExitWin(4).Picture = LoadPicture(sCustomImgs(8))
    If sCustomImgs(9) <> "" Then imgMemoryPage.Picture = LoadPicture(sCustomImgs(9))
    If sCustomImgs(10) <> "" Then imgMemoryRAM.Picture = LoadPicture(sCustomImgs(10))
    If sCustomImgs(11) <> "" Then imgIPs.Picture = LoadPicture(sCustomImgs(11))
    If sCustomImgs(12) <> "" Then imgLock(0).Picture = LoadPicture(sCustomImgs(12)): imgLock(2).Picture = LoadPicture(sCustomImgs(12))
    If sCustomImgs(13) <> "" Then imgLock(1).Picture = LoadPicture(sCustomImgs(13))
    If sCustomImgs(14) <> "" Then imgVolume.Picture = LoadPicture(sCustomImgs(14))
    If sCustomImgs(14) <> "" Then imgVolumeMute(1).Picture = LoadPicture(sCustomImgs(14))
    If sCustomImgs(15) <> "" Then imgPower(1).Picture = LoadPicture(sCustomImgs(15))
    If sCustomImgs(16) <> "" Then imgPower(2).Picture = LoadPicture(sCustomImgs(16))
    If sCustomImgs(17) <> "" Then imgPower(3).Picture = LoadPicture(sCustomImgs(17))
    If sCustomImgs(18) <> "" Then imgPower(4).Picture = LoadPicture(sCustomImgs(18))
    If sCustomImgs(19) <> "" Then imgPower(5).Picture = LoadPicture(sCustomImgs(19))
    If sCustomImgs(20) <> "" Then imgPower(6).Picture = LoadPicture(sCustomImgs(20))
    If sCustomImgs(21) <> "" Then imgPower(7).Picture = LoadPicture(sCustomImgs(21))
    If sCustomImgs(22) <> "" Then imgPower(8).Picture = LoadPicture(sCustomImgs(22))
    If sCustomImgs(23) <> "" Then imgResolution.Picture = LoadPicture(sCustomImgs(23))
    If sCustomImgs(24) <> "" Then imgTime.Picture = LoadPicture(sCustomImgs(24))
    If sCustomImgs(25) <> "" Then imgToggle.Picture = LoadPicture(sCustomImgs(25))
    If sCustomImgs(26) <> "" Then imgUptime.Picture = LoadPicture(sCustomImgs(26))
    If sCustomImgs(27) <> "" Then imgWinamp.Picture = LoadPicture(sCustomImgs(27))
    If sCustomImgs(28) <> "" Then imgWinampC(0).Picture = LoadPicture(sCustomImgs(28))
    If sCustomImgs(29) <> "" Then imgWinampC(1).Picture = LoadPicture(sCustomImgs(29))
    If sCustomImgs(30) <> "" Then imgWinampC(2).Picture = LoadPicture(sCustomImgs(30))
    If sCustomImgs(31) <> "" Then imgWinampC(3).Picture = LoadPicture(sCustomImgs(31))
    If sCustomImgs(32) <> "" Then imgWinampC(4).Picture = LoadPicture(sCustomImgs(32))
    If sCustomImgs(33) <> "" Then imgWinampC(5).Picture = LoadPicture(sCustomImgs(33))
    If sCustomImgs(34) <> "" Then imgWinampC(6).Picture = LoadPicture(sCustomImgs(34))
    If sCustomImgs(35) <> "" Then imgOS.Picture = LoadPicture(sCustomImgs(35))
    'Add more modules below
    If sCustomImgs(36) <> "" Then picProcesses(0).Picture = LoadPicture(sCustomImgs(36))
    If sCustomImgs(37) <> "" Then imgMouseIdle.Picture = LoadPicture(sCustomImgs(37))
    If sCustomImgs(38) <> "" Then imgTCPMonitor.Picture = LoadPicture(sCustomImgs(38))
    If sCustomImgs(39) <> "" Then imgTCPMonitorArrow(0).Picture = LoadPicture(sCustomImgs(39))
    If sCustomImgs(40) <> "" Then imgTCPMonitorArrow(1).Picture = LoadPicture(sCustomImgs(40))
    If sCustomImgs(41) <> "" Then imgMSIE.Picture = LoadPicture(sCustomImgs(41))
    If sCustomImgs(42) <> "" Then imgDX.Picture = LoadPicture(sCustomImgs(42))
    
    If sCustomImgs(43) <> "" Then imgVolumeMute(0).Picture = LoadPicture(sCustomImgs(43))
    If sCustomImgs(44) <> "" Then imgVolume2Mute(0).Picture = LoadPicture(sCustomImgs(44))
    
    If sCustomImgs(45) <> "" Then imgRAS(0).Picture = LoadPicture(sCustomImgs(45))
    If sCustomImgs(45) <> "" Then imgRAS(1).Picture = LoadPicture(sCustomImgs(45))
    If sCustomImgs(46) <> "" Then imgRAS(2).Picture = LoadPicture(sCustomImgs(46))
    If sCustomImgs(47) <> "" Then imgRAS(3).Picture = LoadPicture(sCustomImgs(47))
    If sCustomImgs(48) <> "" Then picNetstat(0).Picture = LoadPicture(sCustomImgs(48))
    
    If sCustomImgs(49) <> "" Then picNetstat(1).Picture = LoadPicture(sCustomImgs(49))
    If sCustomImgs(50) <> "" Then picProcesses(1).Picture = LoadPicture(sCustomImgs(50))
    
    '1  - CD player volume
    '2  - CPU usage
    '3  - Date
    '4  - Disk free space
    '5  - Exit Windows
    '6  - Free pagefile
    '7  - Free RAM
    '8  - IP addresses
    '9  - Lock screen
    '10 - Master volume
    '11 - Power status
    '12 - Screen resolution
    '13 - Time
    '14 - Toggle keys status
    '15 - Uptime
    '16 - WinAmp controls
    '17 - Windows version
    '<Add more modules below>
    '18 - List running processes
    '19 - Mouse idle time
    '20 - TCP Monitor
    '21 - MSIE version
    '22 - DirectX version
    '23 - RAS connection
    '24 - Netstat
    Exit Sub
    
Error:
    ShowError "Main", "frmMain.GetCustomIcons", Err.Number, Err.Description, False
End Sub

Public Sub TriggerTimers(Optional iModule%)
    On Error GoTo Error:
    If iModule = 0 Then
        If timUptime.Enabled Then timUptime_Timer
        If timTime.Enabled Then timTime_Timer
        If timMemoryRAM.Enabled Then timMemoryRAM_Timer
        If timMemoryPage.Enabled Then timMemoryPage_Timer
        If timDiskFreeSpace.Enabled Then timDiskFreeSpace_Timer
        If timVolume.Enabled Then timVolume_Timer
        If timVolume2.Enabled Then timVolume2_Timer
        If timCPU.Enabled Then timCPU_Timer
        If timIPs.Enabled Then timIPs_Timer
        If timPower.Enabled Then timPower_Timer
        If timToggle.Enabled Then timToggle_Timer
        If fraProcesses.Visible Then picProcesses_MouseUp 0, 1, 0, 0, 0
        If timTCPMonitor.Enabled Then
            DrawGrid picGraphTCP
            StartTCPMonitor
            nOldTCPDown = 0
            nOldTCPUp = 0
            timTCPMonitor_Timer
        End If
        If fraMSIE.Visible Then GetMSIEVersion
        If fraDX.Visible Then GetDXVersion
        If fraRAS.Visible Then timRAS_Timer
        If fraNetstat.Visible Then picNetstat_MouseUp 0, 1, 0, 0, 0
    Else
        Select Case iModule
            Case 0:  Exit Sub
            Case 1:  If timVolume2.Enabled Then timVolume2_Timer
            Case 2:  If timCPU.Enabled Then timCPU_Timer
            Case 3:  'nothing
            Case 4:  If timDiskFreeSpace.Enabled Then timDiskFreeSpace_Timer
            Case 5:  'nothing
            Case 6:  If timMemoryPage.Enabled Then timMemoryPage_Timer
            Case 7:  If timMemoryRAM.Enabled Then timMemoryRAM_Timer
            Case 8:  If timIPs.Enabled Then timIPs_Timer
            Case 9:  'nothing
            Case 10: If timVolume.Enabled Then timVolume_Timer
            Case 11: If timPower.Enabled Then timPower_Timer
            Case 12: 'nothing
            Case 13: If timTime.Enabled Then timTime_Timer
            Case 14: If timToggle.Enabled Then timToggle_Timer
            Case 15: If timUptime.Enabled Then timUptime_Timer
            Case 16: 'nothing
            Case 17: 'nothing
            'Add new modules below
            Case 18: If fraProcesses.Visible Then picProcesses_MouseUp 0, 1, 0, 0, 0
            Case 19: If timMouseIdle.Enabled Then timMouseIdle_Timer
            Case 20
                If timTCPMonitor.Enabled Then
                    StartTCPMonitor
                    frmMenu.GetMenuDots
                    nOldTCPDown = 0
                    nOldTCPUp = 0
                    timTCPMonitor_Timer
                End If
            Case 21: GetMSIEVersion
            Case 22: GetDXVersion
            Case 23: If timRAS.Enabled Then timRAS_Timer
            Case 24: If fraNetstat.Visible Then picNetstat_MouseUp 0, 1, 0, 0, 0
        End Select
    End If
    Exit Sub
    
Error:
    ShowError "Main", "frmmain.TriggerTimers", Err.Number, Err.Description, False
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "Uptimer4 is already running.", vbInformation, "uptimer4"
        End
    End If
    
    Load frmMenu
    Dim uWSAData As WSAData
    If WSAStartup(&H101, uWSAData) <> 0 Then MsgBox "Unable to start Winsock!", vbExclamation, "oops"
    ReDim sIPs(0)
    sIPs(0) = "127.0.0.1"
    GetSettings
    GetColors
    GetCustomIcons
    
    GetModules True
    
    'Register program in App Paths regkey
    Dim sRegPath$, sMyPath$
    sRegPath = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\uptimer4.exe", "")
    sMyPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Uptimer4.exe"
    If sRegPath = "" Or LCase(sRegPath) <> LCase(sMyPath) Then
        RegSetString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\uptimer4.exe", "", sMyPath
    End If
    
    'Check for IPHLPAPI.DLL problems
    On Error Resume Next
    If GetIfTable(ByVal 0, ByVal 0, 1) = ERROR_NOT_SUPPORTED Then
        MsgBox "The API functions from Iphlpapi.dll are not supported " & _
               "on your system. The TCP monitor and Netstat modules " & _
               "will be unavailable on your system.", vbExclamation, "damned"
    End If
    If Err.Number = 53 Then
        MsgBox "The Iphlpapi.dll file was not found. The TCP monitor " & _
               "and Netstat modules will be unavailable on your system.", vbExclamation, "damned"
    End If
    
    'Destroy previous appbar space if present
    If RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "LastAppBar", 0) <> 0 Then
        Dim sMsg$
        sMsg = "Uptimer4 either crashed or was not shut down "
        sMsg = sMsg & "properly last time. Would you like to "
        sMsg = sMsg & "open the Modules window to disable a module "
        sMsg = sMsg & "that caused Uptimer4 to crash?"
        If MsgBox(sMsg, vbQuestion + vbYesNo, "Uptimer4") = vbYes Then
            sModules = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "Modules", "15,13,3,7,6")
            bBarDocking = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "BarDocking", 1))
            If bBarDocking Then
                fraBanner.Width = 255
                fraResolution.Width = 255
            End If
            imgBanner.Tag = "recover"
            frmModules.Show 1
            imgBanner.Tag = ""
        End If
        Dim APD As APPBARDATA
        APD.cbSize = Len(APD)
        APD.hwnd = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "LastAppBar", 0)
        SHAppBarMessage ABM_REMOVE, APD
        RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "LastAppBar"
    End If
    
    'Subclass/hook
    If Not RunningInIDE Then HookForm True
    
    'Make appbar if set
    If bBarDocking Then
        MakeAppBar True, bBarPos
    Else
        AlignModules False
        Me.Left = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "FloatX", (Screen.Width - Me.Width) / 2)
        Me.Top = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "FloatY", (Screen.Height - Me.Height) / 2)
        If bBarHidden Then frmMain.Visible = False
    End If
    
    'Hide if set
    If Not bBarHidden Then Me.Show
    
    'Make transparent if not docked
    If Not bBarDocking Then SetFormTransparency Me.hwnd
        
    'Set always on top if set
    If bAlwaysOnTop Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    DoEvents
    
    'Update all visible modules
    TriggerTimers
    
    'Get cool dot options on menus
    frmMenu.GetMenuDots
    
    'Register WinKey hotkeys
    If Not RunningInIDE Then LoadWinKeyHotKeys True
    
    'Show help file if this is first run
    If RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "FirstTimeRun", 1) = 1 Then
        ShellExecute Me.hwnd, "open", App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Uptimer4.hlp", "", "", 1
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "FirstTimeRun", 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Me.Hide
    If bBarDocking Then
        MakeAppBar False, bBarPos
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "BarDocking", 1
    Else
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "BarDocking", 0
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "FloatX", Me.Left
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "FloatY", Me.Top
    End If
    closesocket lSocket
    Webserver False, "", 0, 0
    GetCPULoad True
    MakeTrayIcon False, 0
    If frmMenu.mnuUptimeLoggingEnable.Checked Then LogUptime
    Unload frmMenu
    Do
    Loop Until WSACleanup() = -1
    If Not RunningInIDE Then HookForm False
    If Not RunningInIDE Then RegisterHotkeys False, 0
    If Not RunningInIDE Then LoadWinKeyHotKeys False
End Sub

Private Sub fraBanner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu frmMenu.mnuMain
End Sub

Private Sub fraVolume_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > fraVolumeBar.Left And X < fraVolumeBar.Left + fraVolumeBar.Width Then
        hscVolume.Left = X - 68
        hscVolume_MouseMove 1, 0, 100, 0
    End If
End Sub

Private Sub fraVolume2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > fraVolume2Bar.Left And X < fraVolume2Bar.Left + fraVolume2Bar.Width Then
        hscVolume2.Left = X - 68
        hscVolume2_MouseMove 1, 0, 100, 0
    End If
End Sub

Private Sub fraVolume2Bar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    hscVolume2.Left = fraVolume2Bar.Left + X - 68
    hscVolume2_MouseMove 1, 0, 100, 0
End Sub

Private Sub fraVolumeBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    hscVolume.Left = fraVolumeBar.Left + X - 68
    hscVolume_MouseMove 1, 0, 100, 0
End Sub

Private Sub hscVolume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bLocked Or Button <> 1 Then Exit Sub
    On Error GoTo Error:
    'Move hscVolume to stay under mouse cursor
    Dim lNewPos&
    lNewPos = hscVolume.Left + X - 68
    If lNewPos > fraVolumeBar.Left And _
       lNewPos < fraVolumeBar.Left + fraVolumeBar.Width - hscVolume.Width Then
        hscVolume.Left = lNewPos
    End If
    If hscVolume.Left < fraVolumeBar.Left Then hscVolume.Left = fraVolumeBar.Left
    If hscVolume.Left > fraVolumeBar.Left + fraVolumeBar.Width - hscVolume.Width Then hscVolume.Left = fraVolumeBar.Left + fraVolumeBar.Width - hscVolume.Width
    
    'Map new pos of hscVolume to number between 0-65535
    Dim lVolume&, lMax&, sTmp$
    lMax = fraVolumeBar.Width - hscVolume.Width - 15
    lVolume = 65535 * (hscVolume.Left - fraVolumeBar.Left) / lMax
    If lVolume > 65535 Then lVolume = 65535
    lblVolumePerc.Caption = CStr(Int(lVolume * 100 / 65535)) & "%"
    If lblVolumePerc.Caption = "100%" Then lblVolumePerc.Caption = "100"
    sTmp = Hex(lVolume)
    sTmp = String(4 - Len(sTmp), "0") & sTmp
    sTmp = "&H" & sTmp & sTmp
    lVolume = CLng(sTmp)
    waveOutSetVolume 0, lVolume
    Exit Sub
    
Error:
    timVolume.Enabled = False
    ShowError "Master volume", "hscVolume_MouseMove ", Err.Number, Err.Description, True
End Sub

Private Sub hscVolume2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bLocked Or Button <> 1 Then Exit Sub
    On Error GoTo Error:
    'Move hscVolume to stay under mouse cursor
    Dim lNewPos&
    lNewPos = hscVolume2.Left + X - 68
    If lNewPos > fraVolume2Bar.Left And _
       lNewPos < fraVolume2Bar.Left + fraVolume2Bar.Width - hscVolume2.Width Then
        hscVolume2.Left = lNewPos
    End If
    If hscVolume2.Left < fraVolume2Bar.Left Then hscVolume2.Left = fraVolume2Bar.Left
    If hscVolume2.Left > fraVolume2Bar.Left + fraVolume2Bar.Width - hscVolume2.Width Then hscVolume2.Left = fraVolume2Bar.Left + fraVolume2Bar.Width - hscVolume2.Width
    
    'Map new pos of hscVolume to number between 0-65535
    Dim lVolume&, lMax&, sTmp$
    lMax = fraVolume2Bar.Width - hscVolume2.Width - 15
    lVolume = 65535 * (hscVolume2.Left - fraVolume2Bar.Left) / lMax
    If lVolume > 65535 Then lVolume = 65535
    lblVolume2Perc.Caption = CStr(Int(lVolume * 100 / 65535)) & "%"
    If lblVolume2Perc.Caption = "100%" Then lblVolume2Perc.Caption = "100"
    sTmp = Hex(lVolume)
    sTmp = String(4 - Len(sTmp), "0") & sTmp
    sTmp = "&H" & sTmp & sTmp
    lVolume = CLng(sTmp)
    auxSetVolume 0, lVolume
    Exit Sub

Error:
    timVolume2.Enabled = False
    ShowError "CD player volume", "hscVolume2_MouseMove", Err.Number, Err.Description, True
End Sub

Private Sub imgBanner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bLocked Then Exit Sub
    If Button = 2 Then
        PopupMenu frmMenu.mnuMain
    End If
    If Button = 1 Then
        If Shift = 1 Then
            If Not bBarDocking Then SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            frmModules.Show 1
            If bAlwaysOnTop Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        ElseIf Shift = 2 Then
            If Not bBarDocking Then SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            Load frmSettings
            frmSettings.Show 1
            If bAlwaysOnTop Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        End If
    End If
End Sub

Private Sub imgCPU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuCPU
End Sub

Private Sub imgDate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuDate
End Sub

Private Sub imgDiskFreeSpace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 1 Then imgDiskFreeSpace.BorderStyle = 1
End Sub

Private Sub imgDX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuDX
End Sub

Private Sub imgExitWin_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 1 Then imgExitWin(Index).BorderStyle = 1
End Sub

Private Sub imgExitWin_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Error:
    If bLocked Then Exit Sub
    If Button = 1 Then
        Dim lFlag As Long, lRet&
        imgExitWin(Index).BorderStyle = 0
        If frmMenu.mnuExitWinConfirm.Checked Then
            Select Case Index
                Case 0: lRet = MsgBox("Are you sure you want to LOGOFF?", vbQuestion + vbYesNo, "Uptimer4 exitwin")
                Case 1: lRet = MsgBox("Are you sure you want to REBOOT?", vbQuestion + vbYesNo, "Uptimer4 exitwin")
                Case 2: lRet = MsgBox("Are you sure you want to SUSPEND?", vbQuestion + vbYesNo, "Uptimer4 exitwin")
                Case 3: lRet = MsgBox("Are you sure you want to SHUTDOWN?", vbQuestion + vbYesNo, "Uptimer4 exitwin")
                Case 4: lRet = MsgBox("Are you sure you want to POWEROFF?", vbQuestion + vbYesNo, "Uptimer4 exitwin")
            End Select
            If lRet = vbNo Then Exit Sub
        End If
        If frmMenu.mnuExitWinForce.Checked Then lFlag = EWX_FORCE
        GetShutDownProvilege
        Unload Me
        Select Case Index
            Case 0: ExitWindowsEx EWX_LOGOFF Or lFlag, 0&
            Case 1: ExitWindowsEx EWX_REBOOT Or lFlag, 0&
            Case 2: SetSystemPowerState 1, 0
            Case 3: ExitWindowsEx EWX_SHUTDOWN Or lFlag, 0&
            Case 4: ExitWindowsEx EWX_POWEROFF Or lFlag, 0&
        End Select
        End
    ElseIf Button = 2 Then
        PopupMenu frmMenu.mnuExitWin
    End If
    Exit Sub
    
Error:
    ShowError "Exit Windows", "imgExitWin(" & CStr(Index) & ")_MouseUp", Err.Number, Err.Description, False
End Sub

Private Sub imgLock_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgLock(Index).BorderStyle = 1
End Sub

Private Sub imgLock_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> 0 Then Exit Sub
    On Error GoTo Error:
    If Button = 2 Then
        If Not bIsWinNT And Not bLocked Then PopupMenu frmMenu.mnuLock
    Else
        imgLock(Index).BorderStyle = 0
        'If bIsWinNT Then Exit Sub
        
        Dim hwndTaskBar&, hwndDesktop&, hwndToolbar&
        hwndTaskBar = FindWindow("Shell_TrayWnd", "")
        hwndDesktop = FindWindow("ProgMan", "Program Manager")
        
        If bLocked Then
            'Screen is locked, so unlock
            Dim sPass$
            sPass = GetPassword("Please type the magic word to unlock your screen:", 1, ROT13(RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "LockPassword")))
            If sPass = "1" Then
                ShowWindow hwndDesktop, 1
                ShowWindow hwndTaskBar, 1
                If imgLock(0).Tag <> "" Then
                    ShowWindow CLng(imgLock(0).Tag), 1
                    imgLock(0).Tag = ""
                End If
                SystemParametersInfo SPI_SCREENSAVERRUNNING, 0, 0, 0
                imgLock(0).Picture = imgLock(2).Picture
                imgLock(0).ToolTipText = "Lock screen"
                bLocked = False
                
                If RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TooltipsStayWhenLocked", 0) = 0 Then
                    TriggerTimers MODULE_PROCESSES
                    TriggerTimers MODULE_NETSTAT
                End If
            End If
        Else
            'Screen is unlocked, so lock
            If bIsWinNT Then
                'lock NT workstation, no turning back
                If LockWorkStation() = 0 Then
                    MsgBox "Failed to call LockWorkStation(). Only Windows 2000 and higher support this function.", vbCritical, "oops"
                End If
                Exit Sub
            End If
            
            If RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "LockPassword", "0") = "0" Then
                If MsgBox("You have not set a password to unlock your screen yet. Do you want to set it now?", vbQuestion + vbYesNo, "set lock password") = vbYes Then
                    frmMenu.mnuLockSetPass_Click
                End If
            End If
            ShowWindow hwndDesktop, 0
            hwndToolbar = FindWindow("BaseBar", "")
            If hwndToolbar <> 0 Then
                ShowWindow hwndToolbar, 0
                imgLock(0).Tag = CStr(hwndToolbar)
            End If
            SystemParametersInfo SPI_SCREENSAVERRUNNING, 1, 0, 0
            ShowWindow hwndTaskBar, 0
            
            imgLock(0).Picture = imgLock(1).Picture
            imgLock(0).ToolTipText = "Unlock screen"
            bLocked = True
            
            If RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TooltipsStayWhenLocked", 0) = 0 Then
                DestroyWindow lProcessToolTipHwnd
                DestroyWindow lNetstatToolTipHwnd
            End If
        End If
    End If
    Exit Sub
    
'NoWinNT:
'    MsgBox "This module does not work in Windows NT, i.e. " & _
'      "WinNT 4.x, Win2000, WinXP Pro.", vbCritical, "duh"
    Exit Sub
    
Error:
    ShowError "Lock screen", "imgLock_MouseUp", Err.Number, Err.Description, False
End Sub

Private Sub imgMemoryPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuMemoryPage
End Sub

Private Sub imgMouseIdle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuMouseIdle
End Sub

Private Sub imgMSIE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuMSIE
End Sub

Private Sub imgOS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuOS
End Sub

Private Sub imgPower_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuPower
End Sub

Private Sub imgRAS_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 1 Then imgRAS(Index).BorderStyle = 1
End Sub

Private Sub imgRAS_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bLocked Then Exit Sub
    If Button = 1 Then
        imgRAS(Index).BorderStyle = 0
        If Shift = 0 Then
            If frmMenu.mnuRASConnect.Enabled Then frmMenu.mnuRASConnect_Click
        Else
            If frmMenu.mnuRASDisconnect.Enabled Then frmMenu.mnuRASDisconnect_Click
        End If
    ElseIf Button = 2 Then
        PopupMenu frmMenu.mnuRAS
    End If
End Sub

Private Sub imgResolution_Click()
    If Not bLocked Then PopupMenu frmMenu.mnuRes
End Sub

Public Sub imgDiskFreeSpace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bLocked Then Exit Sub
    If Button = 2 Then PopupMenu frmMenu.mnuDiskFreeSpace
    If Button = 1 Then
        On Error Resume Next
        Dim sFile$
        Randomize
        sFile = "c:\bla" & CStr(Int(Rnd * 100000)) & ".tmp"
        Open sFile For Output As #1
            Print #1, "Uptimer4 temp file, safe to delete"
        Close #1
        DoEvents
        Kill sFile
        imgDiskFreeSpace.BorderStyle = 0
    End If
End Sub

Private Sub imgIPs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuIPs
End Sub

Private Sub imgMemoryRAM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuMemoryRAM
End Sub

Private Sub imgTCPMonitor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuTCPMonitor
End Sub

Private Sub imgTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuTime
End Sub

Private Sub imgUptime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuUptime
End Sub

Private Sub imgVolume_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bLocked Then Exit Sub
    If Button = 2 Then PopupMenu frmMenu.mnuMasterVol
    On Error GoTo Error:
    If Button = 1 Then
        'Mute/unmute
        Dim lVolume&
        waveOutGetVolume 0, lVolume
        If lVolume <> 0 Then
            'mute
            lOldVolume = lVolume
            waveOutSetVolume 0, 0
            imgVolume.Picture = imgVolumeMute(0).Picture
            hscVolume.Visible = False
            timVolume_Timer
        Else
            'unmute
            If lOldVolume = 0 Then Exit Sub
            waveOutSetVolume 0, lOldVolume
            imgVolume.Picture = imgVolumeMute(1).Picture
            timVolume_Timer
            hscVolume.Visible = True
        End If
    End If
    Exit Sub
    
Error:
    ShowError "Master volume", "imgVolume_MouseUp", Err.Number, Err.Description, False
End Sub

Private Sub imgVolume2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bLocked Then Exit Sub
    If Button = 2 Then PopupMenu frmMenu.mnuCDVol
    On Error GoTo Error:
    If Button = 1 Then
        'Mute/unmute
        Dim lVolume&
        auxGetVolume 0, lVolume
        If lVolume <> 0 Then
            'mute
            lOldVolume2 = lVolume
            auxSetVolume 0, 0
            imgVolume2.Picture = imgVolume2Mute(0).Picture
            hscVolume2.Visible = False
            timVolume2_Timer
        Else
            'unmute
            If lOldVolume2 = 0 Then Exit Sub
            auxSetVolume 0, lOldVolume2
            imgVolume2.Picture = imgVolume2Mute(1).Picture
            timVolume2_Timer
            hscVolume2.Visible = True
        End If
    End If
    Exit Sub
    
Error:
    ShowError "CD player volume", "imgVolume2_MouseUp", Err.Number, Err.Description, False
End Sub

Private Sub imgWinamp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 1 Then imgWinamp.BorderStyle = 1
End Sub

Public Sub imgWinamp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Error:
    If bLocked Then Exit Sub
    If Button = 2 Then
        PopupMenu frmMenu.mnuWinamp
        Exit Sub
    End If
    If Button <> 1 Then Exit Sub
    imgWinamp.BorderStyle = 0
    
    'If Winamp is already started running it using
    'ShellExecuteEx will show it - perfect! :)
    Dim lRet&, SEI As SHELLEXECUTEINFO, sWinampPath$
    'Attempt to get winamp.exe path from:
    '* Own settings
    sWinampPath = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "WinampPath", "blaat")
    
    If sWinampPath = "blaat" Then
        '* App paths
        sWinampPath = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\Winamp.exe", sWinampPath)
        '* Program files
        If Dir("C:\Program Files\Winamp\winamp.exe") <> "" Then sWinampPath = "C:\Program Files\Winamp\winamp.exe"
        If Dir("D:\Program Files\Winamp\winamp.exe") <> "" Then sWinampPath = "D:\Program Files\Winamp\winamp.exe"
        If Dir("E:\Program Files\Winamp\winamp.exe") <> "" Then sWinampPath = "E:\Program Files\Winamp\winamp.exe"
        '* Program files\Utils
        If Dir("C:\Program Files\Utils\Winamp\winamp.exe") <> "" Then sWinampPath = "C:\Program Files\Utils\Winamp\winamp.exe"
        If Dir("D:\Program Files\Utils\Winamp\winamp.exe") <> "" Then sWinampPath = "D:\Program Files\Utils\Winamp\winamp.exe"
        '* Root folder
        If Dir("C:\Winamp\winamp.exe") <> "" Then sWinampPath = "C:\Winamp\winamp.exe"
        If Dir("D:\Winamp\winamp.exe") <> "" Then sWinampPath = "D:\Winamp\winamp.exe"
        If Dir("E:\Winamp\winamp.exe") <> "" Then sWinampPath = "E:\Winamp\winamp.exe"
        'Winamp.exe not found? Let user locate it manually
GetWinampPath:
        If MsgBox("Unable to locate winamp.exe. Would you like to locate it yourself?", vbYesNo + vbQuestion, "winamp") = vbYes Then
            sWinampPath = BrowseForFolder(Me.hwnd, "Select the Winamp folder:")
            sWinampPath = sWinampPath & IIf(Right(sWinampPath, 1) = "\", "", "\") & "winamp.exe"
            If Dir(sWinampPath) = "" Then GoTo GetWinampPath:
                
            'If found, save path
            If sWinampPath <> "" And LCase(Dir(sWinampPath)) = "winamp.exe" Then RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "WinampPath", sWinampPath
        Else
            Exit Sub
        End If
    End If
        
    'Winamp.exe (should be) found by now, so run it
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_FLAG_NO_UI
        .hwnd = Me.hwnd
        .lpVerb = "open"
        .nShow = IIf(frmMenu.mnuWinAmpMin.Checked, SW_SHOWMINNOACTIVE, SW_SHOWNOACTIVATE)
        .lpFile = sWinampPath
    End With
    ShellExecuteEx SEI
    Exit Sub
    
Error:
    ShowError "Winamp controls", "imgWinamp_MouseUp", Err.Number, Err.Description, False
End Sub

Private Sub imgWinampC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 1 Then imgWinampC(Index).BorderStyle = 1
End Sub

Public Sub imgWinampC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bLocked Then Exit Sub
    On Error GoTo Error:
    imgWinampC(Index).BorderStyle = 0
    Dim hwndWinamp As Long
    hwndWinamp = FindWindow("Winamp v1.x", vbNullString)
    If hwndWinamp = 0 Then Exit Sub
    'Shift = 1
    'Ctrl = 2
    
    Select Case Index
        Case 0 'Previous
            If (Shift And 2) Then   'Go to first song
                PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON1_CTRL, 0
            Else                    'Go back 1 song
                PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON1, 0
            End If
        Case 1 'Play
            If (Shift And 2) Then   'Open location
                PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON2_CTRL, 0
            Else                    'Just play
                PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON2, 0
            End If
        Case 2 'Pause
            PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON3, 0
        Case 3 'Stop
            If (Shift And 1) Then   'Fade & stop
                PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON4_SHIFT, 0
            ElseIf (Shift And 2) Then 'Stop after current song
                PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON4_CTRL, 0
            Else                    ' Just stop
                PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON4, 0
            End If
        Case 4 'Next
            If (Shift And 2) Then   'Go to last song
                PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON5_CTRL, 0
            Else                    'Go to next song
                PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON5, 0
            End If
        Case 5                      'Open file
            PostMessage hwndWinamp, WM_COMMAND, WINAMP_FILE_PLAY, 0
        Case 6
            Dim sFile$, uCDS As COPYDATASTRUCT
            sFile = GetFileName(True, "MP3 files (*.mp3)|*.mp3|All files (*.*)|*.*", , "Add file to playlist...")
            If sFile = "" Then Exit Sub
            uCDS.cbData = Len(sFile) + 1
            uCDS.lpData = sFile
            uCDS.dwData = IPC_PLAYFILE
            SendMessageCDS hwndWinamp, WM_COPYDATA, 0, uCDS
    End Select
    DoEvents
    timWinamp_Timer
    Exit Sub
    
Error:
    ShowError "Winamp controls", "imgWinampC_MouseUp", Err.Number, Err.Description, False
End Sub

Private Sub imgWinampC_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bLocked Then Exit Sub
    If Index <> 5 And Index <> 6 Then Exit Sub
    On Error Resume Next
    Dim i%, uCDS As COPYDATASTRUCT, hwndWinamp&, lListLen&
    hwndWinamp = FindWindow("Winamp v1.x", vbNullString)
    If hwndWinamp = 0 Then Exit Sub
    
    If Index = 5 Then
        PostMessage hwndWinamp, WM_WA_IPC, 0, IPC_DELETE
    Else
        lListLen = SendMessageGet(hwndWinamp, WM_WA_IPC, 0, IPC_GETLISTLENGTH)
    End If
    
    'Add files in Data.Files() to playlist
    i = 1
    Do
        uCDS.lpData = Data.Files(i)
        If Err Then Exit Do
        uCDS.cbData = Len(Data.Files(i)) + 1
        uCDS.dwData = IPC_PLAYFILE
        SendMessageCDS hwndWinamp, WM_COPYDATA, 0, uCDS
        i = i + 1
    Loop
    If i = 1 Then Exit Sub 'no files
   
    If Index = 5 Then
        'Set winamp to play first one of added files NOW
        SendMessageGet hwndWinamp, WM_WA_IPC, 0, IPC_SETPLAYLISTPOS
        PostMessage hwndWinamp, WM_COMMAND, WINAMP_BUTTON2, 0
    Else
        'Set winamp to play first one of added files next
        SendMessageGet hwndWinamp, WM_WA_IPC, lListLen, IPC_SETPLAYLISTPOS
    End If
End Sub

Private Sub lblBanner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 2 Then PopupMenu frmMenu.mnuMain
End Sub

Private Sub picNetstat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bLocked And Button = 0 And Index = 0 Then timProcesses.Enabled = True
End Sub

Private Sub picNetstat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Error:
    If bLocked Then Exit Sub
    If Button = 2 Then
        PopupMenu frmMenu.mnuNetstat
        Exit Sub
    End If
    
    Dim uTcpTable As MIB_TCPTABLE, uUdpTable As MIB_UDPTABLE
    Dim lSize&, i%, sNetstat$, sTruncated$, j%
    Dim sLocalIP$, sLocalPort$, sRemoteIP$, sRemotePort$, sState$
    Dim bShowTimewait As Boolean, bShowListening As Boolean
    bShowTimewait = frmMenu.mnuNetstatShowTimewait.Checked
    bShowListening = frmMenu.mnuNetstatShowListening.Checked
    
    picNetstat(0).Visible = False
    picNetstat(1).Visible = True
    DoEvents
    
    'Get TCP netstat
    GetTcpTable uTcpTable, lSize, 1
    GetTcpTable uTcpTable, lSize, 1
    
    'Get UDP netstat
    If frmMenu.mnuNetstatShowUDP.Checked Then
        GetUdpTable uUdpTable, lSize, 1
        GetUdpTable uUdpTable, lSize, 1
    End If
    
    'Write header for table
    '[If anyone know how the hell I can align the items
    'right without using tabs (which don't display in the
    'tooltip) PLEASE EMAIL ME!]
    sNetstat = "== xXx netstat items at " & Format(Time, "Long time") & " (click image to refresh) ==" & vbCrLf
    'sNetstat = sNetstat & "Local IP" & String(3, vbTab) & "Local port" & vbTab & "Remote IP" & String(2, vbTab) & "Remote port" & vbTab & "State" & vbCrLf
    sNetstat = sNetstat & "Protocol   Local IP" & String(21, " ") & "Local port" & "       " & "Remote IP" & String(12, " ") & "Remote port" & "      " & "State" & vbCrLf
    sNetstat = sNetstat & String(125, "-") & vbCrLf
    
    'Enumerate TCP items
    For i = 0 To uTcpTable.dwNumEntries - 1
        With uTcpTable.table(i)
            sLocalIP = CStr(Asc(Mid(.dwLocalAddr, 1, 1))) & "." & CStr(Asc(Mid(.dwLocalAddr, 2, 1))) & "." & CStr(Asc(Mid(.dwLocalAddr, 3, 1))) & "." & CStr(Asc(Mid(.dwLocalAddr, 4, 1)))
            'sLocalIP = sLocalIP & String((26 - Len(sLocalIP)) \ 6.5, vbTab)
            sLocalIP = sLocalIP & String(30 - Len(sLocalIP), " ")
            
            sLocalPort = CStr(CLng(Asc(Mid(.dwLocalPort, 1, 1))) * 256 + Asc(Mid(.dwLocalPort, 2, 1)))
            'sLocalPort = sLocalPort & vbTab & vbTab
            sLocalPort = sLocalPort & String(18 - Len(sLocalPort), " ")
            
            Select Case .dwState - 1
                Case MIB_TCP_STATE_SYN_RCVD:   sState = "SYN_RECEIVED"
                Case MIB_TCP_STATE_LISTEN:     sState = "LISTENING"
                Case MIB_TCP_STATE_SYN_SENT:   sState = "SYN_SENT"
                Case MIB_TCP_STATE_FIN_WAIT1:  sState = "FIN_WAIT1"
                Case MIB_TCP_STATE_FIN_WAIT2:  sState = "FIN_WAIT2"
                Case MIB_TCP_STATE_LAST_ACK:   sState = "LAST_ACK"
                Case MIB_TCP_STATE_ESTAB:      sState = "ESTABLISHED"
                Case MIB_TCP_STATE_CLOSING:    sState = "CLOSING"
                Case MIB_TCP_STATE_DELETE_TCB: sState = "DELETE_TCB"
                Case MIB_TCP_STATE_CLOSE_WAIT: sState = "CLOSE_WAIT"
                Case MIB_TCP_STATE_TIME_WAIT:  sState = "TIME_WAIT"
                Case MIB_TCP_STATE_CLOSED:     sState = "STATE_CLOSED"
            End Select
            
            sRemoteIP = CStr(Asc(Mid(.dwRemoteAddr, 1, 1))) & "." & CStr(Asc(Mid(.dwRemoteAddr, 2, 1))) & "." & CStr(Asc(Mid(.dwRemoteAddr, 3, 1))) & "." & CStr(Asc(Mid(.dwRemoteAddr, 4, 1)))
            'sRemoteIP = sRemoteIP & String((26 - Len(sRemoteIP)) \ 6.5, vbTab)
            sRemoteIP = sRemoteIP & String(26 - Len(sRemoteIP), " ")
            
            If .dwState - 1 <> MIB_TCP_STATE_LISTEN Then
                'If item is listening port, set remote port
                'to 0 or junk will show up
                sRemotePort = CStr(CLng(Asc(Mid(.dwRemotePort, 1, 1))) * 256 + Asc(Mid(.dwRemotePort, 2, 1)))
                'sRemotePort = sRemotePort & vbTab & vbTab
                sRemotePort = sRemotePort & String(12, " ")
            Else
                'sRemotePort = "0" & vbTab & vbTab
                sRemotePort = "0" & String(20, " ")
            End If
            
            'Apply TIME_WAIT and LISTENING filters
            If bShowTimewait = False And sState = "TIME_WAIT" Or _
               bShowListening = False And sState = "LISTENING" Then
                'don't show the line
            Else
                j = j + 1
                If iNetstatTrunc > 0 And j = iNetstatTrunc + 1 Then sTruncated = sNetstat & "** truncated **"
                sNetstat = sNetstat & "TCP       " & sLocalIP & sLocalPort & sRemoteIP & sRemotePort & sState & vbCrLf
            End If
        End With
    Next i
    
    'Get UDP netstat
    For i = 0 To uUdpTable.dwNumEntries - 1
        With uUdpTable.table(i)
            sLocalIP = CStr(Asc(Mid(.dwLocalAddr, 1, 1))) & "." & CStr(Asc(Mid(.dwLocalAddr, 2, 1))) & "." & CStr(Asc(Mid(.dwLocalAddr, 3, 1))) & "." & CStr(Asc(Mid(.dwLocalAddr, 4, 1)))
            sLocalIP = sLocalIP & String(30 - Len(sLocalIP), " ")
            
            sLocalPort = CStr(CLng(Asc(Mid(.dwLocalPort, 1, 1))) * 256 + Asc(Mid(.dwLocalPort, 2, 1)))
            sLocalPort = sLocalPort & String(18 - Len(sLocalPort), " ")
            
            j = j + 1
            If iNetstatTrunc > 0 And j = iNetstatTrunc + 1 Then sTruncated = sNetstat & "** truncated **"
            sNetstat = sNetstat & "UDP       " & sLocalIP & sLocalPort & vbCrLf
        End With
    Next i
    
    'Truncated? If not, remove last vbCrLf
    If sTruncated = "" Then
        sNetstat = Left(sNetstat, Len(sNetstat) - 2)
    Else
        'Else, set table to truncated version
        sNetstat = sTruncated
    End If
    
    'Set # of items in header
    sNetstat = Replace(sNetstat, "xXx", CStr(j))
    If j = 0 Then sNetstat = sNetstat & vbCrLf & "(no items)"
    
    'Set label caption and create bubble tooltip
    lblNetstat.Caption = CStr(uTcpTable.dwNumEntries) & " netstat items"
    AssignBubbleTip picNetstat(0).hwnd, sNetstat, MODULE_NETSTAT
    DoEvents
    picNetstat(1).Visible = False
    picNetstat(0).Visible = True
    Exit Sub
    
Error:
    picNetstat(1).Visible = False
    picNetstat(0).Visible = True
    ShowError "Netstat", "picNetstat_MouseUp(" & Button & ")", Err.Number, Err.Description, False
End Sub

Private Sub picProcesses_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 And Index = 0 Then timProcesses.Enabled = True
End Sub

Private Sub picProcesses_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Error:
    If bLocked Then Exit Sub
    If Button = 2 Then
        PopupMenu frmMenu.mnuProcesses
        Exit Sub
    End If
    Dim sProcessList$, hSnap&, uProcess As PROCESSENTRY32
    Dim sTruncated$, i%, sProcess$
    picProcesses(0).Visible = False
    picProcesses(1).Visible = True
    DoEvents
    On Error Resume Next
    hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
    If Err Or hSnap < 1 Then
        MsgBox "Unable to create process list.", vbExclamation, "oops"
        picProcesses(1).Visible = False
        picProcesses(0).Visible = True
        Exit Sub
    End If
    On Error GoTo Error:
    uProcess.dwSize = Len(uProcess)
    If ProcessFirst32(hSnap, uProcess) = 0 Then
        MsgBox "No processes found.", vbExclamation, "wtf?"
        CloseHandle hSnap
        picProcesses(1).Visible = False
        picProcesses(0).Visible = True
        Exit Sub
    End If
    i = 1
    Do
        sProcess = Left(uProcess.szExeFile, InStr(uProcess.szExeFile, Chr(0)) - 1)
        If sProcess <> "[System Process]" And sProcess <> "System" Then
            sProcessList = sProcessList & GetFullPath(sProcess) & vbCrLf
            i = i + 1
            If iProcessTrunc > 0 And i = iProcessTrunc + 1 Then sTruncated = sProcessList
        End If
    Loop Until ProcessNext32(hSnap, uProcess) = 0
    CloseHandle hSnap
    i = i - 1
    If sTruncated = "" Then
        sProcessList = Left(sProcessList, Len(sProcessList) - 2)
    Else
        sProcessList = Left(sTruncated, Len(sTruncated) - 2)
        sProcessList = sTruncated & "** truncated **"
    End If
    sProcessList = "== " & CStr(i) & " running processes at " & CStr(Time) & " (click image to refresh) ==" & vbCrLf & sProcessList
    AssignBubbleTip picProcesses(0).hwnd, sProcessList, MODULE_PROCESSES
    lblProcesses.Caption = CStr(i) & " processes"
    DoEvents
    picProcesses(1).Visible = False
    picProcesses(0).Visible = True
    Exit Sub
    
Error:
    picProcesses(1).Visible = False
    picProcesses(0).Visible = True
    ShowError "List running processes", "picProcesses_MouseUp", Err.Number, Err.Description, False
End Sub

Private Sub timCPU_Timer()
    Dim lCPU&
    On Error GoTo Error:
    lCPU = GetCPULoad
    If lCPU > 100 Then
        lCPU = 100
    End If
        
    If lCPU <> -1 Then
        If frmMenu.mnuCPUDisplay(0).Checked Then
            lblCPU.Caption = CStr(lCPU) & " %"
        Else
            Select Case lCPU
                Case 100:      lblCPU.Caption = "1.00"
                Case 10 To 99: lblCPU.Caption = Left(CStr(lCPU / 100), 4) & IIf(Len(CStr(lCPU \ 100)) = 3, "0", "")
                Case 1 To 9:   lblCPU.Caption = Left(CStr(lCPU / 100), 4)
                Case 0:        lblCPU.Caption = "0.00"
            End Select
            lblCPU.Caption = Replace(lblCPU.Caption, ",", ".")
        End If
        If shpCPUBack.Visible Then shpCPUFore.Width = shpCPUBack.Width * (lCPU / 100)
        If picGraphCPU.Visible Then GraphPoint lCPU, 100, MODULE_CPUUSAGE, bSolidGraphs, picGraphCPU
    End If
    Exit Sub
    
Error:
    timCPU.Enabled = False
    ShowError "CPU usage", "timCPU_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timMemoryPage_Timer()
    Dim uMemStatus As MEMORYSTATUS, lTotPage&, lFreePage&
    On Error GoTo Error:
    uMemStatus.dwLength = Len(uMemStatus)
    GlobalMemoryStatus uMemStatus
    lTotPage = uMemStatus.dwTotalPageFile
    lFreePage = uMemStatus.dwAvailPageFile
    lTotPage = Int(1 + lTotPage \ 1024 ^ 2)
    lFreePage = Int(1 + lFreePage \ 1024 ^ 2)
    
    If frmMenu.mnuMemoryPageDisplay(0).Checked Then
        lblMemoryPage.Caption = CStr(lFreePage) & "/" & CStr(lTotPage) & " MB"
    Else
        lblMemoryPage.Caption = CStr(Int(100 * (lFreePage / lTotPage))) & " %"
    End If
    If shpMemoryPageBack.Visible Then shpMemoryPageFore.Width = shpMemoryPageBack.Width * (lFreePage / lTotPage)
    If picGraphPage.Visible Then GraphPoint lFreePage, lTotPage, MODULE_FREEPAGEFILE, bSolidGraphs, picGraphPage
    Exit Sub
    
Error:
    timMemoryPage.Enabled = False
    ShowError "Free pagefile", "timMemoryPage_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timMouseIdle_Timer()
    Dim uMousePos As POINTAPI, iMins%, iSecs%
    On Error GoTo Error:
    GetCursorPos uMousePos
    If uMousePos.X <> uOldMousePos.X Or uMousePos.Y <> uOldMousePos.Y Then
        iMouseIdle = 0
        iMouseDummy = 0
        lblMouseIdle.Caption = "00:00"
    Else
        iMouseDummy = iMouseDummy + 1
        If iMouseDummy = 4 Then
            iMouseIdle = iMouseIdle + 1
            iMouseDummy = 0
            iMins = iMouseIdle \ 60
            iSecs = iMouseIdle - iMins * 60
            lblMouseIdle.Caption = IIf(iMins < 10, "0", "") & CStr(iMins) & ":" & IIf(iSecs < 10, "0", "") & CStr(iSecs)
            If iMouseIdle > iMouseTimeout Then
                ' flashy flashy
                If lblMouseIdle.ForeColor = &H80000009 Then
                    lblMouseIdle.ForeColor = &HFF&
                Else
                    lblMouseIdle.ForeColor = &H80000009
                End If
            End If
        End If
    End If
        
    uOldMousePos.X = uMousePos.X
    uOldMousePos.Y = uMousePos.Y
    Exit Sub

Error:
    timMouseIdle.Enabled = False
    ShowError "Mouse idle time", "timMouseIdle_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timNetstat_Timer()
    Dim uMousePos As POINTAPI
    GetCursorPos uMousePos
    SetCursorPos uMousePos.X, uMousePos.Y
    timNetstat.Enabled = False
End Sub

Private Sub timPower_Timer()
    Dim SPS As SYSTEM_POWER_STATUS
    On Error GoTo Error:
    GetSystemPowerStatus SPS
    Select Case SPS.BatteryFlag
        'Case POWER_CHARGING:        imgPower(0).Picture = imgPower(2).Picture
        Case POWER_HIGH:            imgPower(0).Picture = imgPower(3).Picture
        Case POWER_LOW:             imgPower(0).Picture = imgPower(4).Picture
        Case POWER_CRITICAL:        imgPower(0).Picture = imgPower(5).Picture
        Case POWER_NOSYSTEMBATTERY: imgPower(0).Picture = imgPower(6).Picture
        Case POWER_UNKNOWN:         imgPower(0).Picture = imgPower(7).Picture
    End Select
    If SPS.ACLineStatus = 1 Then
        If SPS.BatteryLifePercent <= 100 And frmMenu.mnuPowerBatChargeOnAC.Checked Then
            lblPower.Caption = CStr(SPS.BatteryLifePercent) & " %"
        Else
            lblPower.Caption = "AC"
        End If
        If Not (SPS.BatteryFlag And POWER_CHARGING) And SPS.BatteryFlag < 128 Then imgPower(0).Picture = imgPower(1).Picture
    Else
        lblPower.Caption = CStr(SPS.BatteryLifePercent) & " %"
    End If
    imgPower(0).ToolTipText = "Battery: " & CStr(SPS.BatteryLifePercent) & " %"
    If (SPS.BatteryFlag And POWER_CHARGING) And SPS.BatteryLifePercent < 100 Then imgPower(0).Picture = imgPower(2).Picture
    If shpPowerBack.Visible Then
        If SPS.BatteryLifePercent <= 100 Then
            shpPowerFore.Width = shpPowerBack.Width * (100 * SPS.BatteryLifePercent / 100) / 100
        Else
            shpPowerFore.Width = shpPowerBack.Width
            lblPower.Caption = "AC"
        End If
    End If
    Exit Sub
    
Error:
    timPower.Enabled = False
    ShowError "Power status", "timPower_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timProcesses_Timer()
    Dim uMousePos As POINTAPI
    GetCursorPos uMousePos
    SetCursorPos uMousePos.X, uMousePos.Y
    timProcesses.Enabled = False
End Sub

Private Sub timRAS_Timer()
    'Purpose: enum RAS connections, when one is found
    'that's not disconnected, display status accordingly
    
    'Will this work on Win2000/WinXP? uRCS.dwSize
    'needs to 288 maybe?
    
    Dim uRC(255) As RASCONN, uRCS As RASCONNSTATUS95
    Dim lEntries&
    uRC(0).dwSize = LenB(uRC(0))
    uRCS.dwSize = 160 'structure is actually 64 bytes
    
    If RasEnumConnections(uRC(0), 256 * uRC(0).dwSize, lEntries) = 0 Then
        'found a connection, which can be online/offline/busy
        If RasGetConnectStatus(uRC(0).hRasConn, uRCS) = 0 Then
            'not offline, so online or busy connecting
            Select Case uRCS.RasConnState
                Case RASCS_Connected
                    'online
                    If imgRAS(0).Picture <> imgRAS(1).Picture Then
                        imgRAS(0).Picture = imgRAS(1).Picture
                        imgRAS(0).ToolTipText = "RAS connection - online"
                        frmMenu.mnuRASConnect.Enabled = False
                        'frmMenu.mnuRASDisconnect.Enabled = True
                    End If
                Case RASCS_Disconnected
                    'offline (busy disconnecting)
                    If imgRAS(0).Picture <> imgRAS(2).Picture Then
                        imgRAS(0).Picture = imgRAS(2).Picture
                        imgRAS(0).ToolTipText = "RAS connection - disconnecting"
                        frmMenu.mnuRASConnect.Enabled = True
                        'frmMenu.mnuRASDisconnect.Enabled = False
                    End If
                Case Else
                    'busy connecting
                    If imgRAS(0).Picture <> imgRAS(3).Picture Then
                        imgRAS(0).Picture = imgRAS(3).Picture
                        imgRAS(0).ToolTipText = "RAS connection - connecting"
                        frmMenu.mnuRASConnect.Enabled = False
                        'frmMenu.mnuRASDisconnect.Enabled = True
                    End If
            End Select
        Else
            'offline
            If imgRAS(0).Picture <> imgRAS(2).Picture Then
                imgRAS(0).Picture = imgRAS(2).Picture
                imgRAS(0).ToolTipText = "RAS connection - offline"
                frmMenu.mnuRASConnect.Enabled = True
                'frmMenu.mnuRASDisconnect.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub timTCPMonitor_Timer()
    Dim uByte() As Byte, lSize&
    Dim IfRowTable As MIB_IFROW, i%
    Dim nTCPDown As Single, nTCPUp As Single
    Dim nTCPSpeedDown As Single, nTCPSpeedUp As Single
    On Error GoTo Error:
    
    GetIfTable ByVal 0, lSize, 0
    ReDim uByte(lSize)
    GetIfTable uByte(0), lSize, 1
    i = 1
    CopyMemoryTCP IfRowTable, uByte(4 + (i - 1) * Len(IfRowTable)), Len(IfRowTable)
    Do
        With IfRowTable
            If frmMenu.mnuTCPMonitorAll.Checked Then
                'Count all adapters if set to 'All'
                If Not (InStr(LCase(.bDescr), "loopback") > 0 And frmMenu.mnuTCPMonitorIgnoreLoopback.Checked) Then
                    'But exclude loopback adapter if set
                    nTCPDown = nTCPDown + .dwInOctets
                    nTCPUp = nTCPUp + .dwOutOctets
                End If
            Else
                'Count just one adapter
                If frmMenu.mnuTCPMonitorAdapter(i).Checked Then
                    nTCPDown = nTCPDown + .dwInOctets
                    nTCPUp = nTCPUp + .dwOutOctets
                End If
            End If
        End With
        i = i + 1
        CopyMemoryTCP IfRowTable, uByte(4 + (i - 1) * Len(IfRowTable)), Len(IfRowTable)
    Loop Until IfRowTable.dwType = 0
        
    If nOldTCPDown = 0 And nOldTCPUp = 0 Then
        'First run, save data and exit
        nOldTCPDown = nTCPDown
        nOldTCPUp = nTCPUp
        lblTCPMonitorDown.Caption = Format(0, "0.00") & " K/s"
        lblTCPMonitorUp.Caption = Format(0, "0.00") & " K/s"
        Exit Sub
    End If
        
    nTCPSpeedDown = nTCPDown - nOldTCPDown
    nTCPSpeedUp = nTCPUp - nOldTCPUp
        
    If nTCPSpeedDown < 0 Then nTCPSpeedDown = 0
    If nTCPSpeedUp < 0 Then nTCPSpeedUp = 0
        
    Select Case nTCPSpeedDown
        Case 0 To 10 ^ 4 - 1
            'x.xx K/s
            lblTCPMonitorDown.Caption = Format(nTCPSpeedDown / 1024, "0.00") & " K/s"
        Case 10 ^ 4 To 10 ^ 5 - 1
            'xx.x K/s
            lblTCPMonitorDown.Caption = Format(nTCPSpeedDown / 1024, "#0.0") & " K/s"
        Case 10 ^ 5 To 10 ^ 6 - 1
            'xxx  K/s
            lblTCPMonitorDown.Caption = Format(nTCPSpeedDown / 1024, "#00") & " K/s"
        Case 10 ^ 6 To 10 ^ 7 - 1
            'xxxx K/s
            lblTCPMonitorDown.Caption = Format(nTCPSpeedDown / 1024, "#000") & " K/s"
        Case Is > 10 ^ 7
            'over 10 MB/s, not very likely
            lblTCPMonitorDown.Caption = "wowie!"
    End Select
    Select Case nTCPSpeedUp
        Case 0 To 9999
            'x.xx K/s
            lblTCPMonitorUp.Caption = Format(nTCPSpeedUp / 1024, "0.00") & " K/s"
        Case 10000 To 99999
            'xx.x K/s
            lblTCPMonitorUp.Caption = Format(nTCPSpeedUp / 1024, "00.0") & " K/s"
        Case 100000 To 999999
            'xxx  K/s
            lblTCPMonitorUp.Caption = Format(nTCPSpeedUp / 1024, "000") & " K/s"
        Case 1000000 To 9999999
            'xxxx K/s
            lblTCPMonitorUp.Caption = Format(nTCPSpeedUp / 1024, "0000") & " K/s"
        Case Is > 10000000
            'over 10 MB/s, not very likely
            lblTCPMonitorUp.Caption = "wowie!"
    End Select
    If picGraphTCP.Visible Then Graph2PointsDyn CLng(nTCPSpeedDown), CLng(nTCPSpeedUp), MODULE_TCPMONITOR, bSolidGraphs, picGraphTCP
        
    nOldTCPDown = nTCPDown
    nOldTCPUp = nTCPUp
    Exit Sub
    
Error:
    timTCPMonitor.Enabled = False
    ShowError "TCP Monitor", "timTCPMonitor_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timTime_Timer()
    On Error GoTo Error:
    lblTime.Caption = Format(Time, "Hh:Nn:Ss")
    If Minute(Now) = 0 And Second(Now) < 10 Then
        lblDate.Caption = Format(Date, "dd/mm/yyyy")
        lblDate.ToolTipText = Format(Date, "Long date")
    End If
    If bReminderSet Then
        'Dim sReminder$
        'sReminder = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTime")
        'If sReminder = "" Then
        '    bReminderSet = False
        '    timTime.Interval = 1000
        '    Exit Sub
        'End If
        
        Debug.Print Time & " =?= " & sReminderTime
        'Decrease timer interval to decrease change
        'of missing reminder time because of system
        'load... timer tends to lag a bit sometimes
        If Left(Time, 7) = Left(sReminderTime, 7) Then timTime.Interval = 500
        
        If CStr(Time) = sReminderTime Then
            imgBanner.Tag = "alarm"
            Load frmAlarm
            'frmAlarm.Show 1
            timTime.Interval = 1000
        End If
    End If
    Exit Sub
    
Error:
    timTime.Enabled = False
    ShowError "Time", "timTime_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timDiskFreeSpace_Timer()
    If bIsWin95 Then timDiskFreeSpace95_Timer: Exit Sub
    On Error GoTo Error:
    
    Dim lNewBitMask&, i%, lType&
    lNewBitMask = GetLogicalDrives()
    If lNewBitMask <> lBitMask Then
        lBitMask = lNewBitMask
        For i = 0 To 26
            If (lBitMask And 2 ^ i) Then
                lType = GetDriveType(Chr(65 + i) & ":\")
                If lType = DRIVE_FIXED Or lType = DRIVE_RAMDISK Or lType = DRIVE_REMOTE Then
                    ReDim Preserve sDisks(UBound(sDisks) + 1)
                    sDisks(UBound(sDisks)) = Chr(65 + i) & ":\"
                End If
            End If
        Next i
        'Get right caption for menu's, and hide
        'ones for nonexistant disks
        frmMenu.GetMenuDots
    End If
    'sDisks() contains all valid drive paths
    
    Dim cTotalBytes@, cFreeBytes@, cQuotaBytes@
    Dim nFreeSpace As Single, nTotalSpace As Single
    If sCurrentDisk <> "." Then
        'Monitor only one disk
        For i = 1 To UBound(sDisks)
            If sDisks(i) = sCurrentDisk & ":\" Then Exit For
        Next i
        'i now is index of drive to monitor
        If i <= UBound(sDisks) Then
            GetDiskFreeSpaceEx sDisks(i), cQuotaBytes, cTotalBytes, cFreeBytes
            If frmMenu.mnuDiskFreeSpaceQuota.Checked Then
                nFreeSpace = CSng(cQuotaBytes * 10000)
            Else
                nFreeSpace = CSng(cFreeBytes * 10000)
            End If
            nTotalSpace = CSng(cTotalBytes * 10000)
        End If
    Else
        'Monitor all disks, total
        For i = 1 To UBound(sDisks)
            GetDiskFreeSpaceEx sDisks(i), cQuotaBytes, cTotalBytes, cFreeBytes
            If frmMenu.mnuDiskFreeSpaceQuota.Checked Then
                nFreeSpace = nFreeSpace + CSng(cQuotaBytes * 10000)
            Else
                nFreeSpace = nFreeSpace + CSng(cFreeBytes * 10000)
            End If
            nTotalSpace = nTotalSpace + CSng(cTotalBytes * 10000)
        Next i
    End If
    
    'Display free space/total space
    Select Case nFreeSpace
        Case Is < 1024 'bytes
            lblDiskFreeSpace.Caption = CStr(nFreeSpace) & " B/"
        Case 1024 To 1024 ^ 2 - 1 'Kbytes
            lblDiskFreeSpace.Caption = Left(CStr(nFreeSpace / 1024), 5) & " KB/"
        Case 1024 ^ 2 To 1024 ^ 3 - 1 'Mbytes
            lblDiskFreeSpace.Caption = Left(CStr(nFreeSpace / 1024 ^ 2), 5) & " MB/"
        Case Is >= 1024 ^ 3 'Gbytes
            lblDiskFreeSpace.Caption = Left(CStr(nFreeSpace / 1024 ^ 3), 5) & " GB/"
    End Select
    Select Case nTotalSpace
        Case Is < 1024 'B
            lblDiskFreeSpace.Caption = lblDiskFreeSpace.Caption & CStr(nTotalSpace) & " B"
        Case 1024 To 1024 ^ 2 - 1 'KB
            lblDiskFreeSpace.Caption = lblDiskFreeSpace.Caption & Left(CStr(nTotalSpace / 1024), 5) & " KB"
        Case 1024 ^ 2 To 1024 ^ 3 - 1 'MB
            lblDiskFreeSpace.Caption = lblDiskFreeSpace.Caption & Left(CStr(nTotalSpace / 1024 ^ 2), 5) & " MB"
        Case Is >= 1024 ^ 3 'GB
            lblDiskFreeSpace.Caption = lblDiskFreeSpace.Caption & Left(CStr(nTotalSpace / 1024 ^ 3), 5) & " GB"
    End Select
    If shpDiskFreeBack.Visible Then shpDiskFreeFore.Width = shpDiskFreeBack.Width * (nFreeSpace / nTotalSpace)
    Exit Sub
    
Error:
    timDiskFreeSpace.Enabled = False
    ShowError "Disk free space", "timDiskFreeSpace_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timDiskFreeSpace95_Timer()
    Dim lNewBitMask&, i%, lType&
    On Error GoTo Error:
    lNewBitMask = GetLogicalDrives()
    If lNewBitMask <> lBitMask Then
        lBitMask = lNewBitMask
        For i = 0 To 26
            If (lBitMask And 2 ^ i) Then
                lType = GetDriveType(Chr(65 + i) & ":\")
                If lType = DRIVE_FIXED Or lType = DRIVE_RAMDISK Or lType = DRIVE_REMOTE Then
                    ReDim Preserve sDisks(UBound(sDisks) + 1)
                    sDisks(UBound(sDisks)) = Chr(65 + i) & ":\"
                End If
            End If
        Next i
        frmMenu.HideInvalidDisks
        frmMenu.GetMenuDots
    End If
    'sDisks() contains all valid drive paths
    
    Dim lSecPClus&, lBytPSec&, lFreeClus&, lTotClus&
    Dim nFreeSpace As Single, nTotalSpace As Single
    If sCurrentDisk <> "." Then
        'Only monitor one drive
        For i = 1 To UBound(sDisks)
            If sDisks(i) = sCurrentDisk & ":\" Then Exit For
        Next i
        'i now is index of drive to monitor
        GetDiskFreeSpace sDisks(i), lSecPClus, lBytPSec, lFreeClus, lTotClus
        nFreeSpace = lFreeClus * lSecPClus * lBytPSec
        nTotalSpace = lTotClus * lSecPClus * lBytPSec
    Else
        'Monitor all drivers available
        For i = 1 To UBound(sDisks)
            GetDiskFreeSpace sDisks(i), lSecPClus, lBytPSec, lFreeClus, lTotClus
            nFreeSpace = nFreeSpace + lFreeClus * lSecPClus * lBytPSec
            nTotalSpace = nTotalSpace + lTotClus * lSecPClus * lBytPSec
        Next i
    End If
    
    'Display free space/total space
    Select Case nFreeSpace
        Case Is < 1024 'bytes
            lblDiskFreeSpace.Caption = CStr(nFreeSpace) & " B/"
        Case 1024 To 1024 ^ 2 - 1 'Kbytes
            lblDiskFreeSpace.Caption = Left(CStr(nFreeSpace / 1024), 5) & " KB/"
        Case 1024 ^ 2 To 1024 ^ 3 - 1 'Mbytes
            lblDiskFreeSpace.Caption = Left(CStr(nFreeSpace / 1024 ^ 2), 5) & " MB/"
        Case Is >= 1024 ^ 3 'Gbytes
            lblDiskFreeSpace.Caption = Left(CStr(nFreeSpace / 1024 ^ 3), 5) & " GB/"
    End Select
    Select Case nTotalSpace
        Case Is < 1024 'B
            lblDiskFreeSpace.Caption = lblDiskFreeSpace.Caption & CStr(nTotalSpace) & " B"
        Case 1024 To 1024 ^ 2 - 1 'KB
            lblDiskFreeSpace.Caption = lblDiskFreeSpace.Caption & Left(CStr(nTotalSpace / 1024), 5) & " KB"
        Case 1024 ^ 2 To 1024 ^ 3 - 1 'MB
            lblDiskFreeSpace.Caption = lblDiskFreeSpace.Caption & Left(CStr(nTotalSpace / 1024 ^ 2), 5) & " MB"
        Case Is >= 1024 ^ 3 'GB
            lblDiskFreeSpace.Caption = lblDiskFreeSpace.Caption & Left(CStr(nTotalSpace / 1024 ^ 3), 5) & " GB"
    End Select
    If shpDiskFreeBack.Visible Then shpDiskFreeFore.Width = shpDiskFreeBack.Width * (nFreeSpace / nTotalSpace)
    Exit Sub
    
Error:
    timDiskFreeSpace.Enabled = False
    ShowError "Disk free space", "timDiskFreeSpace95_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timIPs_Timer()
    On Error GoTo Error:
    'ex: three IPs: 127.0.0.1, 192.168.0.1, 131.211.4.5
    'ex: two IPs: 127.0.0.1, 10.0.0.19
    'ex: one IP: 127.0.0.1
    If iCurrentIP > IIf(bIgnoreLocalHostIP, 1, 0) Then
        'if not at first IP, display next. else: reread IPs
        lblIPs.Caption = sIPs(iCurrentIP)
        iCurrentIP = iCurrentIP + 1
        If iCurrentIP > UBound(sIPs) Then iCurrentIP = IIf(bIgnoreLocalHostIP, 1, 0)
        Exit Sub
    End If
    
    Dim uHostEnt As HOSTENT, pHostEnt&, i%
    Dim pIP&, MemIP() As Byte
    sHostname = String(255, 0)
    If gethostname(sHostname, 255) = -1 Then GoTo WSockErr:
    sHostname = Left(sHostname, InStr(sHostname, Chr(0)) - 1)
    
    pHostEnt = gethostbyname(sHostname)
    If pHostEnt = 0 Then GoTo WSockErr:
    
    CopyMemory uHostEnt, ByVal pHostEnt, Len(uHostEnt)
    For i = 0 To UBound(sIPs)
        sIPs(i) = ""
    Next i
    sIPs(0) = "127.0.0.1"
    For i = 1 To 10
        CopyMemory pIP, ByVal uHostEnt.hAddrList + 4 * (i - 1), 4
        If pIP = 0 Then Exit For
        ReDim Preserve sIPs(i)
        ReDim MemIP(1 To 4)
        CopyMemory MemIP(1), ByVal pIP, 4
        sIPs(i) = MemIP(1) & "." & MemIP(2) & "." & MemIP(3) & "." & MemIP(4)
    Next i
    iCurrentIP = IIf(bIgnoreLocalHostIP, 1, 0)
    lblIPs.Caption = sIPs(iCurrentIP)
    iCurrentIP = iCurrentIP + 1
    If iCurrentIP > UBound(sIPs) Then iCurrentIP = IIf(bIgnoreLocalHostIP, 1, 0)
    Exit Sub
    
WSockErr:
    MsgBox "Error in IP address module.", vbExclamation, "oops"
    Exit Sub
    
Error:
    timIPs.Enabled = False
    ShowError "IP addresses", "timIPs_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timMemoryRAM_Timer()
    Dim uMemStatus As MEMORYSTATUS, lTotRAM&, lFreeRAM&
    On Error GoTo Error:
    uMemStatus.dwLength = Len(uMemStatus)
    GlobalMemoryStatus uMemStatus
    lTotRAM = uMemStatus.dwTotalPhys
    lFreeRAM = uMemStatus.dwAvailPhys
    lTotRAM = Int(1 + lTotRAM \ 1024 ^ 2)
    lFreeRAM = Int(1 + lFreeRAM \ 1024 ^ 2)
    
    If frmMenu.mnuMemoryRAMDisplay(0).Checked Then
        lblMemoryRAM.Caption = CStr(lFreeRAM) & "/" & CStr(lTotRAM) & " MB"
    Else
        lblMemoryRAM.Caption = CStr(Int(100 * (lFreeRAM / lTotRAM))) & " %"
    End If
    If shpMemoryRAMBack.Visible Then shpMemoryRAMFore.Width = shpMemoryRAMBack.Width * (lFreeRAM / lTotRAM)
    If picGraphRAM.Visible Then GraphPoint lFreeRAM, lTotRAM, MODULE_FREERAM, bSolidGraphs, picGraphRAM
    Exit Sub
    
Error:
    timMemoryRAM.Enabled = False
    ShowError "Free RAM", "timMemoryRAM_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timToggle_Timer()
    On Error GoTo Error:
    If Not bToggleKeysInit Then
        Do
        Loop Until GetAsyncKeyState(VK_CAPITAL) = 0
        Do
        Loop Until GetAsyncKeyState(VK_NUMLOCK) = 0
        Do
        Loop Until GetAsyncKeyState(VK_SCROLL) = 0
        Do
        Loop Until GetAsyncKeyState(VK_INSERT) = 0
        bToggleKeysInit = True
        bToggleKeysState(1) = GetKeyState(VK_CAPITAL)
        bToggleKeysState(2) = GetKeyState(VK_NUMLOCK)
        bToggleKeysState(3) = GetKeyState(VK_SCROLL)
        bToggleKeysState(4) = GetKeyState(VK_INSERT)
    Else
        If GetAsyncKeyState(VK_CAPITAL) Then
            Do
                DoEvents
            Loop Until GetAsyncKeyState(VK_CAPITAL) = 0
            bToggleKeysState(1) = Not bToggleKeysState(1)
        End If
        If GetAsyncKeyState(VK_NUMLOCK) Then
            Do
                DoEvents
            Loop Until GetAsyncKeyState(VK_NUMLOCK) = 0
            bToggleKeysState(2) = Not bToggleKeysState(2)
        End If
        If GetAsyncKeyState(VK_SCROLL) Then
            Do
                DoEvents
            Loop Until GetAsyncKeyState(VK_SCROLL) = 0
            bToggleKeysState(3) = Not bToggleKeysState(3)
        End If
        If GetAsyncKeyState(VK_INSERT) Then
            Do
                DoEvents
            Loop Until GetAsyncKeyState(VK_INSERT) = 0
            bToggleKeysState(4) = Not bToggleKeysState(4)
        End If
    End If
    
    lblToggle(1).BackColor = IIf(bToggleKeysState(1), lColorFore, lColorBack)
    lblToggle(0).BackColor = IIf(bToggleKeysState(2), lColorFore, lColorBack)
    lblToggle(2).BackColor = IIf(bToggleKeysState(3), lColorFore, lColorBack)
    lblToggle(3).BackColor = IIf(bToggleKeysState(4), lColorFore, lColorBack)
    Exit Sub
    
Error:
    timToggle.Enabled = False
    ShowError "Toggle keys status", "timToggle_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timUptime_Timer()
    lblUptime.Caption = GetDuration(GetTickCount(), False)
    If bUptimerLogHourly And Right(lblUptime.Caption, 5) = "00:00" Then LogUptime
    lblUptime.ToolTipText = GetDuration(GetTickCount(), True)
    If shpUptimeBack.Visible Then shpUptimeFore.Width = shpUptimeBack.Width * (CInt(Left(lblUptime.Caption, 2)) / 50)
End Sub

Private Sub timVolume_Timer()
    Dim lVolume&, lMax&
    On Error GoTo Error:
    If waveOutGetNumDevs() = 0 Then
        timVolume.Enabled = False
        lblVolumePerc.Caption = "0 %"
        MsgBox "No wave device found! The master volume module will be disabled.", vbExclamation, "oops"
        Exit Sub
    End If
    lMax = fraVolumeBar.Width - hscVolume.Width
    waveOutGetVolume 0, lVolume
    'Treat speaker as mono for easy use
    'Volume minimum: 0, maximum: 65535
    lVolume = CLng("&H" & Right(Hex(lVolume), 4))
    hscVolume.Left = fraVolumeBar.Left + (lVolume / 65535) * lMax
    lblVolumePerc.Caption = CStr(Int(lVolume * 100 / 65535)) & "%"
    Exit Sub
    
Error:
    timVolume.Enabled = False
    ShowError "Master volume", "timVolume_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timVolume2_Timer()
    Dim lVolume&, lMax&
    On Error GoTo Error:
    If auxGetNumDevs() = 0 Then
        timVolume2.Enabled = False
        lblVolume2Perc.Caption = "0 %"
        MsgBox "No CD player found! The CD player volume module will be disabled.", vbExclamation, "oops"
        Exit Sub
    End If
    lMax = fraVolume2Bar.Width - hscVolume2.Width
    auxGetVolume 0, lVolume
    lVolume = CLng("&H" & Right(Hex(lVolume), 4))
    hscVolume2.Left = fraVolume2Bar.Left + (lVolume / 65535) * lMax
    lblVolume2Perc.Caption = CStr(Int(lVolume * 100 / 65535)) & "%"
    Exit Sub
    
Error:
    timVolume2.Enabled = False
    ShowError "CD player volume", "timVolume2_Timer", Err.Number, Err.Description, True
End Sub

Private Sub timWinamp_Timer()
    Dim hwndWinamp As Long, sBuff$, lRet&
    hwndWinamp = FindWindow("Winamp v1.x", vbNullString)
    If hwndWinamp = 0 Then
        For lRet = 0 To 6
            imgWinampC(lRet).ToolTipText = ""
        Next lRet
        Exit Sub
    End If
    sBuff = String(260, 0)
    lRet = GetWindowText(hwndWinamp, sBuff, 260)
    If lRet > 0 And InStr(sBuff, " - Winamp") > 0 Then
        sBuff = Left(sBuff, lRet)
        sBuff = Left(sBuff, InStr(sBuff, " - Winamp") - 1)
        For lRet = 0 To 6
            imgWinampC(lRet).ToolTipText = sBuff
        Next lRet
    End If
End Sub
