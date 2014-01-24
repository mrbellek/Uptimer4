VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uptimer4 settings"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdHomepage 
      Caption         =   "Visit my homepage"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame fraOfftopic 
      Caption         =   "Totally unrelated to Uptimer4 options"
      Height          =   4335
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame fraTransClick 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1920
         TabIndex        =   134
         Top             =   3885
         Width           =   3855
         Begin VB.CommandButton hscTrans 
            Height          =   255
            Left            =   1920
            TabIndex        =   135
            Top             =   0
            Width           =   135
         End
         Begin VB.Frame fraTransScroll 
            Height          =   30
            Left            =   1920
            TabIndex        =   136
            Top             =   115
            Width           =   1935
         End
         Begin VB.Label lblInfo 
            Caption         =   "Forms transparency: 0%"
            Height          =   195
            Index           =   26
            Left            =   0
            TabIndex        =   137
            Top             =   30
            Width           =   1770
         End
      End
      Begin VB.CommandButton cmdWinKey 
         Caption         =   "WinKey setup"
         Height          =   375
         Left            =   240
         TabIndex        =   88
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox chkEnhKazaa 
         Caption         =   "RefoSearch"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   80
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkEnhKazaa 
         Caption         =   "Grokster"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   79
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox chkEnhKazaa 
         Caption         =   "Morpheus v1.3.3"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   78
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtWinProcess 
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   2940
         Width           =   1335
      End
      Begin VB.TextBox txtMake 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   69
         Top             =   3195
         Width           =   735
      End
      Begin VB.CommandButton cmdMake 
         Caption         =   "Hide:"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   67
         Top             =   3180
         Width           =   1215
      End
      Begin VB.TextBox txtWinHwnd 
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   2940
         Width           =   735
      End
      Begin VB.TextBox txtWinClass 
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2580
         Width           =   2535
      End
      Begin VB.TextBox txtWinCaption 
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "<click a window to see its info>"
         Top             =   2220
         Width           =   2535
      End
      Begin VB.CheckBox chkIgnoreOwn 
         Caption         =   "Ignore own window"
         Height          =   255
         Left            =   3840
         TabIndex        =   59
         Top             =   3300
         Width           =   1815
      End
      Begin VB.Timer timGetTopWindow 
         Interval        =   500
         Left            =   3240
         Top             =   2040
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Top             =   3660
         Width           =   7215
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Height          =   375
         Left            =   6120
         TabIndex        =   57
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtMake 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   56
         Top             =   2730
         Width           =   735
      End
      Begin VB.TextBox txtMake 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   55
         Top             =   2250
         Width           =   735
      End
      Begin VB.CommandButton cmdMake 
         Caption         =   "Make normal:"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   52
         Top             =   2700
         Width           =   1215
      End
      Begin VB.CommandButton cmdMake 
         Caption         =   "Make on top:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   51
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   1740
         Width           =   7215
      End
      Begin VB.CheckBox chkEnhKazaa 
         Caption         =   "KaZaA"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkSmallDesktopIcons 
         Caption         =   "Small desktop icons"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblSunken 
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   9
         Left            =   210
         TabIndex        =   89
         Top             =   3810
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "MorpheusEnhance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   5880
         MousePointer    =   10  'Up Arrow
         TabIndex        =   83
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "by ÅutoBot"
         Height          =   195
         Index           =   6
         Left            =   5880
         TabIndex        =   82
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Hide ad banner space and media bar of FastTrack clients:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   81
         Top             =   960
         Width           =   4110
      End
      Begin VB.Label lblSunken 
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   7
         Left            =   6090
         TabIndex        =   75
         Top             =   3810
         Width           =   1275
      End
      Begin VB.Label lblSunken 
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   6
         Left            =   210
         TabIndex        =   74
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label lblSunken 
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   5
         Left            =   210
         TabIndex        =   73
         Top             =   2670
         Width           =   1275
      End
      Begin VB.Label lblSunken 
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   4
         Left            =   210
         TabIndex        =   72
         Top             =   2190
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "PID:"
         Height          =   195
         Index           =   15
         Left            =   5520
         TabIndex        =   71
         Top             =   3000
         Width           =   315
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "hwnd"
         Height          =   195
         Index           =   14
         Left            =   1560
         TabIndex        =   68
         Top             =   3240
         Width           =   390
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "hwnd:"
         Height          =   195
         Index           =   13
         Left            =   3840
         TabIndex        =   63
         Top             =   3000
         Width           =   435
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Classname:"
         Height          =   195
         Index           =   12
         Left            =   3840
         TabIndex        =   62
         Top             =   2640
         Width           =   810
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Caption:"
         Height          =   195
         Index           =   11
         Left            =   3840
         TabIndex        =   61
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Current topwindow info:"
         Height          =   195
         Index           =   10
         Left            =   3840
         TabIndex        =   60
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "hwnd"
         Height          =   195
         Index           =   9
         Left            =   1560
         TabIndex        =   54
         Top             =   2775
         Width           =   390
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "hwnd"
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   53
         Top             =   2295
         Width           =   390
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Make window always on top:"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   50
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Ripped from"
         Height          =   195
         Index           =   5
         Left            =   5880
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "WinTidy 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   5400
         MousePointer    =   10  'Up Arrow
         TabIndex        =   47
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Ripped from                    by pcmag.com"
         Height          =   195
         Index           =   1
         Left            =   4440
         TabIndex        =   46
         Top             =   600
         Width           =   2790
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "These settings take effect immediately!"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   300
         Width           =   2745
      End
   End
   Begin VB.Frame fraWinKey 
      Caption         =   "Win shortcuts setup - another cool program ripoff I wanted to add as an exercise, so quit bitching!"
      Height          =   4335
      Left            =   120
      TabIndex        =   90
      Top             =   120
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   2
         Left            =   120
         TabIndex        =   94
         Top             =   3660
         Width           =   7215
      End
      Begin VB.CommandButton cmdBack2 
         Caption         =   "Back"
         Height          =   375
         Left            =   6120
         TabIndex        =   92
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Frame fraDummy 
         BorderStyle     =   0  'None
         Height          =   2535
         Index           =   4
         Left            =   120
         TabIndex        =   104
         Top             =   1080
         Visible         =   0   'False
         Width           =   7215
         Begin VB.Frame fraDummy 
            BorderStyle     =   0  'None
            Height          =   1215
            Index           =   1
            Left            =   2880
            TabIndex        =   118
            Top             =   1200
            Width           =   2175
            Begin VB.OptionButton optWinKeyMod 
               Caption         =   "Alt key"
               Height          =   255
               Index           =   3
               Left            =   720
               TabIndex        =   123
               Top             =   960
               Width           =   975
            End
            Begin VB.OptionButton optWinKeyMod 
               Caption         =   "Shift key"
               Height          =   255
               Index           =   2
               Left            =   720
               TabIndex        =   122
               Top             =   720
               Width           =   975
            End
            Begin VB.OptionButton optWinKeyMod 
               Caption         =   "Control key"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   121
               Top             =   480
               Width           =   1095
            End
            Begin VB.OptionButton optWinKeyMod 
               Caption         =   "None extra"
               Height          =   255
               Index           =   0
               Left            =   720
               TabIndex        =   120
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.CheckBox chkWinKeyMod 
               Caption         =   "Windows key"
               Enabled         =   0   'False
               Height          =   255
               Left            =   720
               TabIndex        =   119
               Top             =   0
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "Modifiers:"
               Height          =   195
               Index           =   23
               Left            =   0
               TabIndex        =   124
               Top             =   0
               Width           =   675
            End
         End
         Begin VB.Frame fraDummy 
            BorderStyle     =   0  'None
            Height          =   975
            Index           =   2
            Left            =   5160
            TabIndex        =   113
            Top             =   1200
            Width           =   1935
            Begin VB.OptionButton optWinKeyState 
               Caption         =   "Maximized"
               Height          =   255
               Index           =   2
               Left            =   600
               TabIndex        =   116
               Top             =   720
               Width           =   1215
            End
            Begin VB.OptionButton optWinKeyState 
               Caption         =   "Minimized"
               Height          =   255
               Index           =   1
               Left            =   600
               TabIndex        =   115
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton optWinKeyState 
               Caption         =   "Normal"
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   114
               Top             =   0
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               Caption         =   "State:"
               Height          =   195
               Index           =   25
               Left            =   0
               TabIndex        =   117
               Top             =   0
               Width           =   420
            End
         End
         Begin VB.CommandButton cmdWinKeyGetProg 
            Caption         =   ".."
            Height          =   255
            Left            =   6720
            TabIndex        =   112
            Top             =   30
            Width           =   255
         End
         Begin VB.CommandButton cmdWinKeyGetPath 
            Caption         =   ".."
            Height          =   255
            Left            =   6720
            TabIndex        =   111
            Top             =   750
            Width           =   255
         End
         Begin VB.CommandButton cmdWinKeyCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   1560
            TabIndex        =   110
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton cmdWinKeyOK 
            Caption         =   "OK"
            Height          =   375
            Left            =   120
            TabIndex        =   109
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtWinKeyCmd 
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   108
            Top             =   720
            Width           =   5535
         End
         Begin VB.TextBox txtWinKeyCmd 
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   107
            Top             =   360
            Width           =   5535
         End
         Begin VB.TextBox txtWinKeyCmd 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   106
            Top             =   0
            Width           =   5535
         End
         Begin VB.ComboBox cboWinKey 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   105
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "Key:"
            Height          =   195
            Index           =   24
            Left            =   1080
            TabIndex        =   132
            Top             =   1200
            Width           =   315
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "Start path:"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   131
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "Parameters:"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   130
            Top             =   360
            Width           =   840
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "Program:"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   129
            Top             =   0
            Width           =   630
         End
         Begin VB.Label lblSunken 
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   14
            Left            =   90
            TabIndex        =   128
            Top             =   2010
            Width           =   1275
         End
         Begin VB.Label lblSunken 
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   15
            Left            =   1530
            TabIndex        =   127
            Top             =   2010
            Width           =   1275
         End
         Begin VB.Label lblSunken 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   16
            Left            =   6690
            TabIndex        =   126
            Top             =   0
            Width           =   315
         End
         Begin VB.Label lblSunken 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   17
            Left            =   6690
            TabIndex        =   125
            Top             =   720
            Width           =   315
         End
      End
      Begin VB.Frame fraDummy 
         BorderStyle     =   0  'None
         Height          =   2535
         Index           =   3
         Left            =   240
         TabIndex        =   95
         Top             =   960
         Width           =   6855
         Begin VB.ListBox lstWinKey 
            Height          =   2265
            IntegralHeight  =   0   'False
            Left            =   0
            TabIndex        =   99
            Top             =   240
            Width           =   5775
         End
         Begin VB.CommandButton cmdWinKeyAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   5880
            TabIndex        =   98
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdWinKeyDel 
            Caption         =   "Delete"
            Height          =   375
            Left            =   5880
            TabIndex        =   97
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdWinKeyEdit 
            Caption         =   "Edit"
            Height          =   375
            Left            =   5880
            TabIndex        =   96
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblSunken 
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   13
            Left            =   5850
            TabIndex        =   103
            Top             =   1170
            Width           =   915
         End
         Begin VB.Label lblSunken 
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   12
            Left            =   5850
            TabIndex        =   102
            Top             =   690
            Width           =   915
         End
         Begin VB.Label lblSunken 
            BorderStyle     =   1  'Fixed Single
            Height          =   435
            Index           =   11
            Left            =   5850
            TabIndex        =   101
            Top             =   210
            Width           =   915
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "Installed hotkeys:"
            Height          =   195
            Index           =   19
            Left            =   0
            TabIndex        =   100
            Top             =   0
            Width           =   1230
         End
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmSettings.frx":014A
         Height          =   495
         Index           =   18
         Left            =   120
         TabIndex        =   93
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label lblSunken 
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   10
         Left            =   6090
         TabIndex        =   91
         Top             =   3810
         Width           =   1275
      End
   End
   Begin VB.Frame fraWebserver 
      Caption         =   "Webserver"
      Height          =   2055
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      WhatsThisHelpID =   2020
      Width           =   4215
      Begin VB.ComboBox cboWebserver2 
         Height          =   315
         ItemData        =   "frmSettings.frx":021B
         Left            =   1920
         List            =   "frmSettings.frx":0222
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1590
         Width           =   2055
      End
      Begin VB.CheckBox chkWebserver2 
         Caption         =   "But listen only on adapter of:"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtWebserver 
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "80"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox chkWebserver 
         Caption         =   "Enable webserver on port:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblWebserver 
         Caption         =   $"frmSettings.frx":0231
         Height          =   855
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame fraBarPos 
      Caption         =   "Bar position"
      Height          =   2175
      Left            =   120
      TabIndex        =   22
      Top             =   120
      WhatsThisHelpID =   2010
      Width           =   2055
      Begin VB.OptionButton optBarPos 
         Caption         =   "Floating"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   320
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optBarPos 
         Caption         =   "Bottom of screen"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkAlwaysOnTop 
         Caption         =   "Always on top"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton optBarPos 
         Caption         =   "Top of screen"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optBarPos 
         Caption         =   "Hidden"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.Frame fraColors 
      Caption         =   "Colors"
      Height          =   2175
      Left            =   2280
      TabIndex        =   23
      Top             =   120
      Width           =   2055
      Begin VB.Label lblColorsGraph 
         Caption         =   "2nd line color:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   34
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblColorsGraph 
         Caption         =   "1st line color:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   33
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblColorsGraph 
         Caption         =   "Grid color:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   32
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblColorsGraph 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   31
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label lblColorsGraph 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   30
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lblColorsGraph 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   29
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblColorsInfo 
         Caption         =   "Click a part of the preview module or preview box to change the color:"
         Height          =   795
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Dummy"
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
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   1060
         Width           =   615
      End
      Begin VB.Label lblColorFore 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1060
         Width           =   855
      End
      Begin VB.Label lblColorBack 
         BackColor       =   &H00800080&
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   1060
         Width           =   615
      End
   End
   Begin VB.Frame fraMisc 
      Caption         =   "Misc options"
      Height          =   4335
      Left            =   4440
      TabIndex        =   25
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdAddTimeServer 
         Caption         =   "..."
         Height          =   255
         Left            =   2760
         TabIndex        =   87
         Top             =   3390
         Width           =   255
      End
      Begin VB.ComboBox cboTimeServer 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton cmdOffTopic 
         Caption         =   "Misc. offtopic options..."
         Height          =   375
         Left            =   240
         TabIndex        =   76
         Top             =   3840
         Width           =   2655
      End
      Begin VB.CheckBox chkSunken 
         Caption         =   "Cool sunken buttons in dialogs"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CheckBox chkTooltipsStayWhenLocked 
         Caption         =   "List running processes/Netstat tooltips stay when screen is locked"
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   1660
         Width           =   2775
      End
      Begin VB.Frame fraDummy 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   2520
         Width           =   2535
         Begin VB.OptionButton optSolidGraphs 
            Caption         =   "Line graphs"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optSolidGraphs 
            Caption         =   "Solid graphs"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "Graph style:"
            Height          =   195
            Index           =   17
            Left            =   0
            TabIndex        =   86
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.CheckBox chkIgnoreFullscreen 
         Caption         =   "Ignore fullscreen apps"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   2700
      End
      Begin VB.CheckBox chkMultiRows 
         Caption         =   "Enable multiple rows"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1140
         Width           =   1815
      End
      Begin VB.CheckBox chkTrayIcon 
         Caption         =   "Use tray icon"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   540
         Width           =   1455
      End
      Begin VB.CheckBox chkAutorun 
         Caption         =   "Autorun on boot from Registry"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chkCopyMirc 
         Caption         =   "Copy text in mIRC colors"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblSunken 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   18
         Left            =   2730
         TabIndex        =   133
         Top             =   3360
         Width           =   315
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "NTP time server:"
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   85
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Label lblSunken 
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   8
         Left            =   210
         TabIndex        =   77
         Top             =   3810
         Width           =   2715
      End
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   3
      Left            =   6330
      TabIndex        =   41
      Top             =   4530
      Width           =   1275
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   2
      Left            =   5010
      TabIndex        =   40
      Top             =   4530
      Width           =   1275
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   1
      Left            =   3690
      TabIndex        =   39
      Top             =   4530
      Width           =   1275
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   0
      Left            =   90
      TabIndex        =   38
      Top             =   4530
      Width           =   1995
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sBarOptions$
Private lColors&(1 To 6)
Private lEnableWebserver& '<- is 0 when disabled, port when enabled
Private iEnableWebserverAdapter%
Private iPowerReplace%
Private bPowerReplaceBack As Boolean
Private bAutorun As Boolean
Private bMircColors As Boolean
Private bMultiRows As Boolean
Private bTooltipsStayWhenLocked As Boolean

Private Sub cboTimeServer_Click()
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "TimeServerIndex", cboTimeServer.ListIndex
End Sub

Private Sub cboTimeServer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cboTimeServer.Text = "time.ien.it" Or _
           cboTimeServer.Text = "ntps1-0.cs.tu-berlin.de" Or _
           cboTimeServer.Text = "ntp-cup.external.hp.com" Then
            MsgBox "Only user-defined servers can be deleted.", vbExclamation, "blah"
            Exit Sub
        End If
        If MsgBox("Are you sure you want to delete this time server from the list?", vbYesNo + vbQuestion, "delete server") = vbNo Then Exit Sub
        Dim sDeadServer$, i%, sExtraServer$
        sDeadServer = cboTimeServer.List(cboTimeServer.ListIndex)
        For i = 0 To 99
            sExtraServer = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "ExtraTimeServer" & String(2 - Len(CStr(i)), "0") & CStr(i))
            If sExtraServer = sDeadServer Then
                RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "ExtraTimeServer" & String(2 - Len(CStr(i)), "0") & CStr(i)
                Exit For
            End If
        Next i
        cboTimeServer.RemoveItem cboTimeServer.ListIndex
        cboTimeServer.ListIndex = 0
    End If
End Sub

Private Sub chkAlwaysOnTop_Click()
    chkIgnoreFullscreen.Enabled = CBool(chkAlwaysOnTop.Value)
End Sub

Private Sub chkEnhKazaa_Click(Index As Integer)
    If chkEnhKazaa(Index).Value = 1 Then
        KazaaEnhance True, Index
    Else
        KazaaEnhance False, Index
    End If
End Sub

Private Sub chkSmallDesktopIcons_Click()
    Dim hwndDesktop&, hwndDummy&, uStyleStr As StyleStruct
    hwndDummy = GetDesktopWindow()
    hwndDummy = FindWindowEx(hwndDummy, 0, "Progman", "Program Manager")
    hwndDummy = FindWindowEx(hwndDummy, 0, "SHELLDLL_DefView", vbNullString)
    hwndDesktop = FindWindowEx(hwndDummy, 0, "SysListView32", vbNullString)
    If hwndDesktop = 0 Then Exit Sub
    
    With uStyleStr
        .dwOld = GetWindowLong(hwndDesktop, GWL_STYLE)
        .dwNew = .dwOld And Not LVS_TYPEMASK
        .dwNew = .dwNew Or IIf(chkSmallDesktopIcons.Value = 1, LVS_SMALLICON, LVS_ICON)
    End With
    SendMessage hwndDesktop, WM_STYLECHANGED, GWL_STYLE, uStyleStr
    ShowWindow hwndDesktop, 1
End Sub

Private Sub chkWebserver_Click()
    If chkWebserver.Value = 0 Then
        txtWebserver.Enabled = False
        txtWebserver.BackColor = &H8000000F
        chkWebserver2.Enabled = False
        chkWebserver2.Value = 0
        cboWebserver2.Enabled = False
        cboWebserver2.BackColor = &H8000000F
    Else
        txtWebserver.Enabled = True
        txtWebserver.BackColor = &H80000005
        chkWebserver2.Enabled = True
        chkWebserver2.Value = 1
        cboWebserver2.Enabled = True
        cboWebserver2.BackColor = &H80000005
    End If
End Sub

Private Sub chkWebserver2_Click()
    If chkWebserver2.Value = 0 Then
        cboWebserver2.Enabled = False
        cboWebserver2.BackColor = &H8000000F
    Else
        cboWebserver2.Enabled = True
        cboWebserver2.BackColor = &H80000005
    End If
End Sub

Private Sub cmdAddTimeServer_Click()
    Dim sNewTimeServer$, sMsg$, i%
    sMsg = "Enter the IP address or hostname of the new "
    sMsg = sMsg & "NTP time server. " & vbCrLf & "Note: "
    sMsg = sMsg & "Uptimer4 queries the time server at port "
    sMsg = sMsg & "37, not all time servers may support this."
    sMsg = sMsg & vbCrLf & vbCrLf & "To delete a server from "
    sMsg = sMsg & "the list, select it and hit Delete."
    sNewTimeServer = InputBox(sMsg, "New NTP time server")
    If sNewTimeServer = "" Then Exit Sub
    For i = 0 To 99
        sMsg = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "ExtraTimeServer" & String(2 - Len(CStr(i)), "0") & CStr(i))
        If sMsg = "" Then
            RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "ExtraTimeServer" & String(2 - Len(CStr(i + 1)), "0") & CStr(i), sNewTimeServer
            Exit For
        End If
        If i = 99 Then
            MsgBox "No slots are free for extra time servers. Delete some to free up slots.", vbExclamation, "oops"
            Exit Sub
        End If
    Next i
    cboTimeServer.AddItem sNewTimeServer, 0
    cboTimeServer.ListIndex = 0
End Sub

Private Sub cmdBack_Click()
    fraOfftopic.Visible = False
    
    fraBarPos.Visible = True
    fraColors.Visible = True
    fraMisc.Visible = True
    fraWebserver.Visible = True
End Sub

Private Sub cmdBack2_Click()
    fraWinKey.Visible = False
    fraOfftopic.Visible = True
End Sub

Private Sub cmdMake_Click(Index As Integer)
    Dim hwndWin&
    If IsNumeric(txtMake(Index).Text) Then hwndWin = CLng(txtMake(Index).Text)
    
    Select Case Index
        Case 0 'make window always on top
            ShowWindow hwndWin, 1
            SetWindowPos hwndWin, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
        Case 1 'make window not on top
            ShowWindow hwndWin, 1
            SetWindowPos hwndWin, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
        Case 2 'hide window
            ShowWindow hwndWin, 0
    End Select
End Sub

Private Sub cmdOffTopic_Click()
    fraBarPos.Visible = False
    fraColors.Visible = False
    fraMisc.Visible = False
    fraWebserver.Visible = False
    
    fraOfftopic.Visible = True
End Sub

Private Sub cmdWinKey_Click()
    fraOfftopic.Visible = False
    fraWinKey.Visible = True
    
    lstWinKey.Clear
    lstWinKey.AddItem String(100, "-")
    lstWinKey.AddItem "Win + Break = <System Properties>"
    lstWinKey.AddItem "Win + Tab = <Cycle through windows>"
    lstWinKey.AddItem "Win + F1 = <Windows Help>"
    lstWinKey.AddItem "Win + D = <Show Desktop>"
    lstWinKey.AddItem "Win + E = <Start Explorer>"
    lstWinKey.AddItem "Win + F = <Find Files Dialog>"
    lstWinKey.AddItem "Win + Ctrl + F = <Find Computer Dialog>"
    lstWinKey.AddItem "Win + M = <Minimize all windows>"
    lstWinKey.AddItem "Win + Shift + M = <Unminimize all windows>"
    lstWinKey.AddItem "Win + R = <Run Dialog>"
    If bIsWinXP Then lstWinKey.AddItem "Win + L = <Logoff>"
    
    Dim i%, sDummy$, vDummy As Variant
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(i)), "0") & CStr(i))
    If sDummy <> "" Then
        Do
            vDummy = Split(sDummy, "|")
            sDummy = "Win + "
            Select Case vDummy(4)
                Case 1: sDummy = sDummy & "Ctrl + "
                Case 2: sDummy = sDummy & "Shift + "
                Case 3: sDummy = sDummy & "Alt + "
            End Select
            sDummy = sDummy & vDummy(3) & " = "
            sDummy = sDummy & vDummy(0)
            lstWinKey.AddItem sDummy, i
            
            i = i + 1
            sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(i)), "0") & CStr(i))
        Loop Until sDummy = ""
    End If
End Sub

Private Sub cmdWinKeyAdd_Click()
    fraDummy(3).Visible = False
    fraDummy(4).Visible = True
End Sub

Private Sub cmdWinKeyCancel_Click()
    fraDummy(4).Visible = False
    fraDummy(3).Visible = True
End Sub

Private Sub cmdWinKeyDel_Click()
    'check if selected is system hotkey
    If lstWinKey.List(lstWinKey.ListIndex) = String(100, "-") Then Exit Sub
    If InStr(lstWinKey.List(lstWinKey.ListIndex), "<") > 0 Then
        MsgBox "This is a system shortcut which can't be deleted.", vbExclamation, "oops"
        Exit Sub
    End If
    
    'delete selected hotkey
    Dim i%
    i = lstWinKey.ListIndex
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(i)), "0") & CStr(i)
    lstWinKey.RemoveItem i
    
    'shift other hotkeys back by one
    Dim sDummy$, j%
    j = i + 1
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(j)), "0") & CStr(j))
    Do
        RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(j - 1)), "0") & CStr(j - 1), sDummy
        j = j + 1
        sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(j)), "0") & CStr(j))
    Loop Until sDummy = ""
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(j - 1)), "0") & CStr(j - 1)
    
    'refill listbox
    cmdWinKey_Click
End Sub

Private Sub cmdWinKeyEdit_Click()
    If lstWinKey.List(lstWinKey.ListIndex) = String(50, "-") Then Exit Sub
    If InStr(lstWinKey.List(lstWinKey.ListIndex), "<") > 0 Then
        MsgBox "This is a system shortcut which can't be edited.", vbExclamation, "oops"
        Exit Sub
    End If
    If lstWinKey.ListIndex = -1 Then Exit Sub
    
    Dim sDummy$, vDummy As Variant
    sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(lstWinKey.ListIndex)), "0") & CStr(lstWinKey.ListIndex))
    vDummy = Split(sDummy, "|")
    
    txtWinKeyCmd(0).Text = vDummy(0)
    txtWinKeyCmd(1).Text = vDummy(1)
    txtWinKeyCmd(2).Text = vDummy(2)
    If Len(vDummy(3)) = 1 Then
        cboWinKey.ListIndex = Asc(vDummy(3)) - 65
    Else
        cboWinKey.ListIndex = CInt(Mid(vDummy(3), 2)) + 25
    End If
    optWinKeyMod(vDummy(4)).Value = True
    optWinKeyState(vDummy(5)).Value = True
    
    cmdWinKeyEdit.Tag = CStr(lstWinKey.ListIndex)
    fraDummy(3).Visible = False
    fraDummy(4).Visible = True
End Sub

Private Sub cmdWinKeyGetPath_Click()
    Dim sPath$
    sPath = BrowseForFolder(Me.hwnd, "Select start path:")
    If sPath <> "" Then txtWinKeyCmd(2).Text = sPath
End Sub

Private Sub cmdWinKeyGetProg_Click()
    Dim sProg$
    sProg = GetFileName(True, "Programs (*.exe)|*.exe|All files (*.*)|*.*", "exe", "Select program...")
    If sProg <> "" Then txtWinKeyCmd(0).Text = sProg
End Sub

Private Sub cmdWinKeyOK_Click()
    Dim i%, sWinKey$
    'Validate
    If Len(txtWinKeyCmd(0).Text) < 5 Or Dir(txtWinKeyCmd(0).Text, vbArchive + vbHidden + vbReadOnly + vbSystem) = "" Then
        MsgBox "The program you entered was not found.", vbCritical, "blah"
        txtWinKeyCmd(0).SetFocus
        txtWinKeyCmd(0).SelStart = 0
        txtWinKeyCmd(0).SelLength = Len(txtWinKeyCmd(0).Text)
        Exit Sub
    End If
    If InStr(txtWinKeyCmd(1).Text, "|") > 0 Or _
       InStr(txtWinKeyCmd(1).Text, "<") > 0 Or _
       InStr(txtWinKeyCmd(1).Text, ">") > 0 Then
        MsgBox "Unable to save hotkey: the parameters " & _
               "field contains a pipe ('|') which is " & _
               "used internally in Uptimer4 to save " & _
               "the shortcut key combination.", vbCritical, "oops"
        txtWinKeyCmd(1).SetFocus
        txtWinKeyCmd(1).SelStart = 0
        txtWinKeyCmd(1).SelLength = Len(txtWinKeyCmd(1).Text)
        Exit Sub
    End If
    If Dir(txtWinKeyCmd(2).Text, vbDirectory + vbHidden + vbReadOnly + vbSystem) = "" Then
        MsgBox "The start path you entered was not found.", vbCritical, "blah"
        txtWinKeyCmd(2).SetFocus
        txtWinKeyCmd(2).SelStart = 0
        txtWinKeyCmd(2).SelLength = Len(txtWinKeyCmd(2).Text)
        Exit Sub
    End If
    
    'Save program path, parameters and start path
    sWinKey = txtWinKeyCmd(0).Text & "|" & txtWinKeyCmd(1).Text & "|" & txtWinKeyCmd(2).Text & "|"
    'Save key
    sWinKey = sWinKey & cboWinKey.Text & "|"
    'Save modifier
    sWinKey = sWinKey & IIf(optWinKeyMod(0).Value, "0", IIf(optWinKeyMod(1).Value, "1", IIf(optWinKeyMod(2).Value, "2", "3"))) & "|"
    'Save state
    sWinKey = sWinKey & IIf(optWinKeyState(0).Value, "0", IIf(optWinKeyState(1).Value, "1", "2"))
    
    i = -1
    'Get free slot, if editing existing pick old one
    If cmdWinKeyEdit.Tag <> "" Then
        i = cmdWinKeyEdit.Tag
        cmdWinKeyEdit.Tag = ""
    Else
        Do
            i = i + 1
        Loop Until RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(i)), "0") & CStr(i)) = ""
    End If
    
    'Save to Registry
    RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "WinKey" & String(2 - Len(CStr(i)), "0") & CStr(i), sWinKey
    
    'Update hotkey listbox
    cmdWinKey_Click
    
    fraDummy(4).Visible = False
    fraDummy(3).Visible = True
End Sub

Private Sub Form_Load()
    Dim i%
    On Error GoTo Error:
    cboWebserver2.ListIndex = 0
    'cboPowerReplace.ListIndex = 0
    For i = 0 To 18
        lblSunken(i).Visible = bCoolSunkenButtons
    Next i
    SetFormTransparency Me.hwnd
    If bIsWin2000 = False And bIsWinXP = False Then
    'If 0 Then
        lblInfo(26).Enabled = False
        hscTrans.Visible = False
    Else
        i = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "Transparency", 0)
        If i < 0 Or i > 99 Then i = 0
        hscTrans.Left = fraTransScroll.Left + (i / 99) * (fraTransScroll.Width - hscTrans.Width - 15)
        lblInfo(26).Caption = "Forms transparency: " & CStr(i) & "%"
    End If
    
    Dim sExtraTimeServer$
    For i = 0 To 99
        sExtraTimeServer = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "ExtraTimeServer" & String(2 - Len(CStr(i)), "0") & CStr(i))
        If sExtraTimeServer <> "" Then
            cboTimeServer.AddItem sExtraTimeServer
        Else
            Exit For
        End If
    Next i
    cboTimeServer.AddItem "time.ien.it"
    cboTimeServer.AddItem "ntps1-0.cs.tu-berlin.de"
    cboTimeServer.AddItem "ntp-cup.external.hp.com"
    cboTimeServer.ListIndex = 0
    
    For i = 1 To 26
        cboWinKey.AddItem Chr(i + 64)
    Next i
    For i = 1 To 12
        cboWinKey.AddItem "F" & CStr(i)
    Next i
    cboWinKey.ListIndex = 0
    
    'BarOptions:
    ' 1st #: 0 = Floating
    '        1 = Top of screen
    '        2 = Bottom of screen
    '        3 = Hidden (disabled due to trouble)
    ' 2nd #: 0 = Normal
    '        1 = Always on top
    ' 3rd #: 0 = No tray icon
    '        1 = Tray icon
    'ex: sBarOptions = "210"
    'means bar is docked on bottom of screen (2),
    'always on top (1), no tray icon (0).
    If bBarDocking Then
        sBarOptions = IIf(bBarPos, "1", "2")
    Else
        sBarOptions = IIf(bBarHidden, "3", "0")
    End If
    sBarOptions = sBarOptions & CStr(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlwaysOnTop", 1))
    sBarOptions = sBarOptions & CStr(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TrayIcon", 0))
    
    lColors(1) = lColorText
    lColors(2) = lColorFore
    lColors(3) = lColorBack
    lColors(4) = lColorGraphGrid
    lColors(5) = lColorGraph1st
    lColors(6) = lColorGraph2nd
    
    lEnableWebserver = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "Webserver", 0)
    iEnableWebserverAdapter = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "WebserverAdapter", 0)
    
    'iPowerReplace = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "PowerReplace", 0)
    'bPowerReplaceBack = CBool(IIf(iPowerReplace = 0, 0, RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "PowerReplaceBack", 1)))
    
    If LCase(RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Uptimer4")) = _
       LCase(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "Uptimer4.exe") Then
        bAutorun = True
    Else
        bAutorun = False
    End If
    bMircColors = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "MircColors", 0))
    bMultiRows = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "Multirows", 0))
    bIgnoreFullscreenApps = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "IgnoreFullscreenApps", Abs(CLng(bIgnoreFullscreenApps))))
    bSolidGraphs = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "SolidGraphs", 0))
    bTooltipsStayWhenLocked = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "TooltipsStayWhenLocked", 0))
    
    
    'Apply settings to window
    optBarPos(CInt(Left(sBarOptions, 1))).Value = True
    chkAlwaysOnTop.Value = Mid(sBarOptions, 2, 1)
    chkTrayIcon.Value = Mid(sBarOptions, 3, 1)
    lblColor.ForeColor = lColors(1)
    lblColorFore.BackColor = lColors(2)
    lblColorBack.BackColor = lColors(3)
    lblColorsGraph(0).BackColor = lColors(4)
    lblColorsGraph(1).BackColor = lColors(5)
    lblColorsGraph(2).BackColor = lColors(6)
    
    chkWebserver_Click
    If UBound(sIPs) > 0 Then
        cboWebserver2.Clear
        For i = 0 To UBound(sIPs)
            cboWebserver2.AddItem sIPs(i)
        Next i
    End If
    cboWebserver2.ListIndex = IIf(iEnableWebserverAdapter <> 255, iEnableWebserverAdapter, 0)
    If iEnableWebserverAdapter = 255 Then chkWebserver2.Value = 0
    txtWebserver.Text = IIf(lEnableWebserver = 0, "80", CStr(lEnableWebserver))
    chkWebserver.Value = IIf(lEnableWebserver = 0, 0, 1)
    
    'cboPowerReplace.ListIndex = iPowerReplace
    'If iPowerReplace = 0 Then chkPowerReplace.Enabled = False
    'chkPowerReplace.Value = bPowerReplaceBack
    
    chkAutorun.Value = Abs(CInt(bAutorun))
    chkCopyMirc.Value = Abs(CInt(bMircColors))
    chkMultiRows.Value = Abs(CInt(bMultiRows))
    chkIgnoreFullscreen.Value = Abs(CInt(bIgnoreFullscreenApps))
    If chkAlwaysOnTop.Value = 0 Then chkIgnoreFullscreen.Enabled = False
    optSolidGraphs(Abs(CInt(bSolidGraphs))).Value = True
    chkTooltipsStayWhenLocked.Value = Abs(CInt(bTooltipsStayWhenLocked))
    chkSunken.Value = Abs(CInt(bCoolSunkenButtons))
    
    chkIgnoreOwn.Value = CInt(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "OffTopicIgnoreOwn", 1))
    
    'Check for active KaZaA clone
    If FindWindow("KaZaA", vbNullString) <> 0 Then
        Dim hwndKazaa&, sTitle$, j&
        hwndKazaa = FindWindow("KaZaA", vbNullString)
        
        'Window caption
        j = GetWindowTextLength(hwndKazaa)
        sTitle = String(j + 1, 0)
        GetWindowText hwndKazaa, sTitle, j + 1
        sTitle = Left(sTitle, InStr(sTitle, Chr(0)) - 1)
        
        'If title contains 'KaZaA' -> it's Kazaa, etc
        If InStr(sTitle, "KaZaA") > 0 Then
            chkEnhKazaa(0).Enabled = True
        ElseIf InStr(sTitle, "Morpheus") > 0 Then
            chkEnhKazaa(1).Enabled = True
        ElseIf InStr(sTitle, "Grokster") > 0 Then
            chkEnhKazaa(2).Enabled = True
        ElseIf InStr(sTitle, "RefoSearch") > 0 Then
            chkEnhKazaa(3).Enabled = True
        End If
    End If
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    Exit Sub
    
Error:
    ShowError "Main", "frmSettings_Load", Err.Number, Err.Description, False
End Sub

Private Sub cmdApply_Click()
    Dim sNewBarOptions$, lNewColors&(1 To 6)
    Dim iNewEnableWebserverAdapter%, iNewPowerReplace%
    Dim lNewEnableWebserver&
    Dim bNewPowerReplaceBack As Boolean
    Dim bNewAutorun As Boolean, bNewMircColors As Boolean
    Dim bNewMultiRows As Boolean, bNewIgnoreFullscreenApps As Boolean
    Dim bNewSolidGraphs As Boolean
    Dim bNewCoolSunkenButtons As Boolean
    Dim bNewTooltipsStayWhenLocked As Boolean
    
    On Error GoTo Error:
    
    'Match up new settings
    If optBarPos(0).Value Then sNewBarOptions = "0"
    If optBarPos(1).Value Then sNewBarOptions = "1"
    If optBarPos(2).Value Then sNewBarOptions = "2"
    If optBarPos(3).Value Then sNewBarOptions = "3"
    sNewBarOptions = sNewBarOptions & CStr(chkAlwaysOnTop.Value)
    'Prompt to enable trayicon if hidden mode is selected
    If optBarPos(3).Value And chkTrayIcon.Value = 0 Then
        If MsgBox("If you run Uptimer4 hidden you can " & _
                   "only access it through the trayicon. " & _
                   "Enable trayicon?", vbQuestion + vbYesNo, "hidden mode") = vbYes Then
            chkTrayIcon.Value = 1
        End If
    End If
    sNewBarOptions = sNewBarOptions & CStr(chkTrayIcon.Value)
    
    lNewColors(1) = lblColor.ForeColor
    lNewColors(2) = lblColorFore.BackColor
    lNewColors(3) = lblColorBack.BackColor
    lNewColors(4) = lblColorsGraph(0).BackColor
    lNewColors(5) = lblColorsGraph(1).BackColor
    lNewColors(6) = lblColorsGraph(2).BackColor
    
    lNewEnableWebserver = IIf(chkWebserver.Value = 0, 0, CLng(txtWebserver.Text))
    iNewEnableWebserverAdapter = IIf(chkWebserver2.Value = 1, cboWebserver2.ListIndex, 255)
    
    'iNewPowerReplace = cboPowerReplace.ListIndex
    'bNewPowerReplaceBack = CBool(chkPowerReplace.Value)
    
    bNewAutorun = CBool(chkAutorun.Value)
    bNewMircColors = CBool(chkCopyMirc.Value)
    bNewMultiRows = CBool(chkMultiRows.Value)
    bNewIgnoreFullscreenApps = CBool(chkIgnoreFullscreen.Value)
    bNewSolidGraphs = optSolidGraphs(1).Value
    bNewTooltipsStayWhenLocked = CBool(chkTooltipsStayWhenLocked.Value)
    bNewCoolSunkenButtons = CBool(chkSunken.Value)
    
    
    'Check for changes (above could be shortened)
    If sBarOptions <> sNewBarOptions Then
        'ex: sBarOptions = "210"
        'means bar is docked on bottom of screen (2),
        'always on top (1), no tray icon (0).
        If Left(sNewBarOptions, 1) <> Left(sBarOptions, 1) Then
            If bBarDocking Then MakeAppBar False, True
            If Left(sNewBarOptions, 1) <> "3" Then
                ShowWindow frmMain.hwnd, 1
                bBarHidden = False
            End If
            Select Case Left(sNewBarOptions, 1)
                Case "0" 'Floating
                    'nothing
                Case "1" 'Top of screen
                    MakeAppBar True, True
                Case "2" 'Bottom of screen
                    MakeAppBar True, False
                Case "3" 'Hidden
                    ShowWindow frmMain.hwnd, 0
                    bBarHidden = True
            End Select
            RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "BarHidden", Abs(CLng(bBarHidden))
        End If
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "BarDocking", Abs(CInt(bBarDocking))
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "BarPos", Abs(CInt(bBarPos))
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlwaysOnTop", CLng(Mid(sNewBarOptions, 2, 1))
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "TrayIcon", CLng(Right(sNewBarOptions, 1))
        bAlwaysOnTop = CBool(Mid(sNewBarOptions, 2, 1))
        If CBool(Right(sNewBarOptions, 1)) = True Then
            MakeTrayIcon True, frmMenu.picTrayIcon.hwnd
        Else
            MakeTrayIcon False, 0
        End If
    End If
    
    If lNewColors(1) <> lColors(1) Or _
       lNewColors(2) <> lColors(2) Or _
       lNewColors(3) <> lColors(3) Or _
       lNewColors(4) <> lColors(4) Or _
       lNewColors(5) <> lColors(5) Or _
       lNewColors(6) <> lColors(6) Or _
       bSolidGraphs <> bNewSolidGraphs Then
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ColorText", lNewColors(1)
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ColorFore", lNewColors(2)
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ColorBack", lNewColors(3)
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ColorGraphGrid", lNewColors(4)
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ColorGraph1st", lNewColors(5)
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ColorGraph2nd", lNewColors(6)
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "SolidGraphs", Abs(CInt(bNewSolidGraphs))
        frmMain.GetColors
        bSolidGraphs = bNewSolidGraphs
        frmMain.TriggerTimers MODULE_FREERAM
        frmMain.TriggerTimers MODULE_FREEPAGEFILE
        frmMain.TriggerTimers MODULE_CPUUSAGE
        frmMain.TriggerTimers MODULE_TCPMONITOR
    End If
    
    If lEnableWebserver <> lNewEnableWebserver Or _
       iEnableWebserverAdapter <> iNewEnableWebserverAdapter Then
        'NB: lEnableWebserver is port, 0 is disabled
        '    iEnableWebserverAdapter is 255 when disabled
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "Webserver", lNewEnableWebserver
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "WebserverAdapter", CLng(iNewEnableWebserverAdapter)
        
        Webserver False, "", 0, 0
        DoEvents
        If lNewEnableWebserver Then
            If iNewEnableWebserverAdapter <> 255 Then
                Webserver True, sIPs(cboWebserver2.ListIndex), lNewEnableWebserver, frmMenu.picWebserverAccept.hwnd
            Else
                Webserver True, "0.0.0.0", lNewEnableWebserver, frmMenu.picWebserverAccept.hwnd
            End If
        End If
    End If
    
    If bAutorun <> bNewAutorun Then
        If bNewAutorun Then
            RegSetString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Uptimer4", App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "uptimer4.exe"
        Else
            RegDelValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Uptimer4"
        End If
    End If
    If bMircColors <> bNewMircColors Then
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "MircColors", Abs(CInt(bNewMircColors))
    End If
    If bMultiRows <> bNewMultiRows Then
        bEnableMultiRows = bNewMultiRows
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "Multirows", Abs(CInt(bNewMultiRows))
        Dim lOldBarHeight&
        lOldBarHeight = frmMain.Height
        AlignModules bBarDocking
        If bBarDocking And lOldBarHeight <> frmMain.Height Then
            Dim APD As APPBARDATA
            APD.cbSize = Len(APD)
            APD.hwnd = frmMain.hwnd
            SHAppBarMessage ABM_REMOVE, APD
            DoEvents
            MakeAppBar True, bBarPos
        End If
    End If
    If bIgnoreFullscreenApps <> bNewIgnoreFullscreenApps Then
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "IgnoreFullscreenApps", Abs(CInt(bNewIgnoreFullscreenApps))
        bIgnoreFullscreenApps = bNewIgnoreFullscreenApps
    End If
    If bTooltipsStayWhenLocked <> bNewTooltipsStayWhenLocked Then
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "TooltipsStayWhenLocked", Abs(CInt(bNewTooltipsStayWhenLocked))
    End If
    If bCoolSunkenButtons <> bNewCoolSunkenButtons Then
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "CoolSunkenButtons", Abs(CInt(bNewCoolSunkenButtons))
        bCoolSunkenButtons = bNewCoolSunkenButtons
    End If
    
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "OffTopicIgnoreOwn", CLng(chkIgnoreOwn.Value)
    LoadWinKeyHotKeys False
    LoadWinKeyHotKeys True
    If Not bBarDocking Then SetFormTransparency frmMain.hwnd
    
    Form_Load
    Exit Sub
    
Error:
    ShowError "Main", "frmSettings.cmdApply_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHomepage_Click()
    ShellExecute Me.hwnd, "open", "http://www.geocities.com/merijn_bellekom/new/", "", "", SW_SHOWNORMAL
End Sub

Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub fraTransClick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > fraTransScroll.Left And X < fraTransScroll.Left + fraTransScroll.Width Then
        hscTrans.Left = X - 68
        hscTrans_MouseMove 1, 0, 100, 0
        hscTrans_MouseUp 1, 0, 100, 0
    End If
End Sub

Private Sub fraTransScroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    hscTrans.Left = fraTransScroll.Left + X - 68
    hscTrans_MouseMove 1, 0, 100, 0
End Sub

Private Sub hscTrans_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Dim lNewPos&
    lNewPos = hscTrans.Left + X - 68
    If lNewPos > fraTransScroll.Left And _
       lNewPos < fraTransScroll.Left + fraTransScroll.Width - hscTrans.Width Then
        hscTrans.Left = lNewPos
    End If
    iTransparency = 99 * (hscTrans.Left - fraTransScroll.Left) / (fraTransScroll.Width - hscTrans.Width - 15)
    If iTransparency > 99 Then iTransparency = 99
    If iTransparency < 0 Then iTransparency = 0
    lblInfo(26).Caption = "Forms transparency: " & CStr(iTransparency) & "%"
End Sub

Private Sub hscTrans_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    SetFormTransparency Me.hwnd
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "Transparency", CLng(iTransparency)
End Sub

Private Sub lblColor_Click()
    lblColor.ForeColor = GetColor(lblColor.ForeColor)
End Sub

Private Sub lblColorBack_Click()
    lblColorBack.BackColor = GetColor(lblColorBack.BackColor)
End Sub

Private Sub lblColorFore_Click()
    lblColorFore.BackColor = GetColor(lblColorFore.BackColor)
End Sub

Private Sub lblColorsGraph_Click(Index As Integer)
    lblColorsGraph(Index).BackColor = GetColor(lblColorsGraph(Index).BackColor)
End Sub

Private Sub lblInfo_Click(Index As Integer)
    If Index = 2 Then ShellExecute frmMain.hwnd, "open", "http://www.pcmag.com/article/0,2997,a=16802,00.asp", "", "", 1
    If Index = 4 Then ShellExecute frmMain.hwnd, "open", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=32013", "", "", 1
End Sub

Private Sub lstWinKey_DblClick()
    cmdWinKeyEdit_Click
End Sub

Private Sub timGetTopWindow_Timer()
    Dim hwndForeground&, sBuffer$, i&, lRet&
    
    'get foreground (i.e. active) window hwnd
    hwndForeground = GetForegroundWindow()
    If hwndForeground = 0 Then Exit Sub
    If hwndForeground = Me.hwnd And chkIgnoreOwn.Value = 1 Then Exit Sub
    txtWinHwnd.Text = CStr(hwndForeground)
    
    'get caption of window
    i = GetWindowTextLength(hwndForeground)
    sBuffer = String(i + 1, 0)
    lRet = GetWindowText(hwndForeground, sBuffer, i + 1)
    If lRet = 0 Or lRet <> i Then
        txtWinCaption.Text = "<empty>"
    Else
        txtWinCaption.Text = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    End If
    
    'get classname of window
    sBuffer = String(255, 0)
    lRet = GetClassName(hwndForeground, sBuffer, 255)
    If lRet = 0 Then
        txtWinClass.Text = "<empty>"
    Else
        txtWinClass.Text = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    End If
    
    'get processID of window
    lRet = GetWindowThreadProcessId(hwndForeground, i)
    If lRet = 0 Then
        txtWinProcess.Text = "?"
    Else
        txtWinProcess.Text = CStr(CSng(i) + 2 ^ 31)
    End If
End Sub
