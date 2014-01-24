VERSION 5.00
Begin VB.Form frmAlarm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm clock"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   Icon            =   "frmAlarm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timPlaySound 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   3840
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame fraSetup 
      Caption         =   "Setup"
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   4335
      Begin VB.CommandButton cmdTestSound 
         Caption         =   "Test sound"
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox txtAlarmTime 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Text            =   "hh:mm:ss"
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtPlayDelay 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2700
         TabIndex        =   12
         Text            =   "5"
         Top             =   1305
         Width           =   495
      End
      Begin VB.CommandButton cmdText 
         Caption         =   "Set text..."
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtSnooze 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2700
         TabIndex        =   8
         Text            =   "10"
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optPlaySound 
         Caption         =   "Play wav file [...]"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton optPlaySound 
         Caption         =   "Play default system beep"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.CheckBox chkPlayContinuously 
         Caption         =   "Play sound continuously each              seconds"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.Label lblSunken 
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   5
         Left            =   3090
         TabIndex        =   21
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label lblSunken 
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Index           =   4
         Left            =   3090
         TabIndex        =   20
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label lblAlarmTime 
         AutoSize        =   -1  'True
         Caption         =   "Alarm goes off at:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label lblSnooze 
         Caption         =   "Snooze makes alarm play again in              minutes"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1710
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdSnooze 
      Caption         =   "Snooze"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdTurnOff 
      Caption         =   "Turn off alarm"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   3
      Left            =   3090
      TabIndex        =   19
      Top             =   1770
      Width           =   1275
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   2
      Left            =   210
      TabIndex        =   18
      Top             =   1770
      Width           =   1275
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   1
      Left            =   3090
      TabIndex        =   17
      Top             =   3810
      Width           =   1275
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   0
      Left            =   90
      TabIndex        =   16
      Top             =   3810
      Width           =   1275
   End
   Begin VB.Label lblAlarm 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click here to set font, fontsize, style && color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sSoundFile$

Private Sub chkPlayContinuously_Click()
    If chkPlayContinuously.Value = 1 Then
        txtPlayDelay.Enabled = True
        txtPlayDelay.BackColor = &H80000005
    Else
        txtPlayDelay.Enabled = False
        txtPlayDelay.BackColor = &H8000000F
    End If
End Sub

Private Sub cmdCancel_Click()
    bReminderSet = False
    sReminderTime = ""
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmIsSet"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTime"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmText"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFont"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontSize"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontColor"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontBold"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontItalic"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontUnderLine"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontStrikeOut"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundBeep"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundFile"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundPlayCont"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSnoozeTime"
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'save shit & stuff
    bReminderSet = True
    sReminderTime = txtAlarmTime.Text
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmIsSet", 1
    RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTime", txtAlarmTime.Text
    
    RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "AlarmText", Replace(lblAlarm.Caption, vbCrLf, "\n")
    RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFont", lblAlarm.FontName
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontSize", lblAlarm.FontSize
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontColor", lblAlarm.ForeColor
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontBold", Abs(CLng(lblAlarm.FontBold))
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontItalic", Abs(CLng(lblAlarm.FontItalic))
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontUnderline", Abs(CLng(lblAlarm.FontUnderline))
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontStrikeOut", Abs(CLng(lblAlarm.FontStrikethru))
    
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundBeep", Abs(CLng(optPlaySound(0).Value))
    If optPlaySound(1).Value Then
        RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundFile", optPlaySound(1).Tag
    Else
        RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundFile"
    End If
    If chkPlayContinuously.Value = 1 Then
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundPlayCont", CLng(Val(txtPlayDelay.Text))
    Else
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundPlayCont", 0
    End If
    
    RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSnoozeTime", CLng(Val(txtSnooze.Text))
    Unload Me
End Sub

Private Sub cmdSnooze_Click()
    Dim iSnoozeTime%, sOldTime$, sNewTime$
    timPlaySound.Enabled = False
    PlaySound "", ByVal 0, 0
    iSnoozeTime = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSnoozeTime", 10)
    sOldTime = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTime", Format(Time, "Long Time"))
    sNewTime = DateAdd("n", iSnoozeTime, sOldTime)
    RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTime", sNewTime
    sReminderTime = sNewTime
    Unload Me
End Sub

Private Sub cmdTestSound_Click()
    If cmdTestSound.Caption = "Test sound" Then
        If optPlaySound(0).Value Then
            sSoundFile = "SystemExclamation"
        Else
            sSoundFile = optPlaySound(1).Tag
        End If
        If chkPlayContinuously.Value = 1 Then
            timPlaySound.Interval = 1000 * Val(txtPlayDelay.Text)
            timPlaySound.Enabled = True
        End If
        timPlaySound_Timer
        cmdTestSound.Caption = "Stop"
    Else
        timPlaySound.Enabled = False
        PlaySound vbNull, ByVal 0, 0
        cmdTestSound.Caption = "Test sound"
    End If
End Sub

Private Sub cmdText_Click()
    Dim sAlarmText$
    sAlarmText = InputBox("Type the text you want to display with the alarm. Use \n to indicate a return.", "Set alarm text")
    If sAlarmText = "" Then Exit Sub
    lblAlarm.Caption = Replace(sAlarmText, "\n", vbCrLf)
End Sub

Private Sub cmdTurnOff_Click()
    timPlaySound.Enabled = False
    PlaySound "", ByVal 0, 0
    bReminderSet = False
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmIsSet"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTime"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmText"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFont"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontSize"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontColor"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontBold"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontItalic"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontUnderLine"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontStrikeOut"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundBeep"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundFile"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundPlayCont"
    RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSnoozeTime"
    Unload Me
End Sub

Private Sub Form_Load()
    lblSunken(0).Visible = bCoolSunkenButtons
    lblSunken(1).Visible = bCoolSunkenButtons
    lblSunken(2).Visible = bCoolSunkenButtons
    lblSunken(3).Visible = bCoolSunkenButtons
    lblSunken(4).Visible = bCoolSunkenButtons
    lblSunken(5).Visible = bCoolSunkenButtons
    SetFormTransparency Me.hWnd
    
    'load stuff etc
    If frmMain.imgBanner.Tag = "setup" Then
        If sReminderTime = "" Then
            txtAlarmTime.Text = Format(Time, "Long Time")
        Else
            txtAlarmTime.Text = sReminderTime
        End If
        cmdSnooze.Visible = False
        cmdTurnOff.Visible = False
        cmdOK.Default = True
        fraSetup.Left = 120
        'Me.Height = 4710
        'chkPlayContinuously_Click
    End If
    If frmMain.imgBanner.Tag = "alarm" Then
        fraSetup.Visible = False
        cmdOK.Visible = False
        cmdCancel.Visible = False
        cmdSnooze.Default = True
        Me.Height = 3210
        
        'set alarm text & effects
        With lblAlarm
            .Caption = Replace(RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmText"), "\n", vbCrLf)
            .ForeColor = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontColor", &H80000012)
            .FontName = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFont", "MS Sans Serif")
            .FontSize = RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontSize", 18)
            .FontBold = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontBold", 1))
            .FontItalic = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontItalic", 0))
            .FontUnderline = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontUnderline", 0))
            .FontStrikethru = CBool(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmTextFontStrikeOut", 0))
        End With
        
        'set off alarm! whoohoo!
        Me.Show
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 1 Or SWP_NOMOVE Or SWP_NOSIZE
        If RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundBeep", 1) = 1 Then
            sSoundFile = "SystemExclamation"
        Else
            sSoundFile = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundFile")
        End If
        timPlaySound.Interval = 1000 * RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "AlarmSoundPlayCont", 5)
        If timPlaySound.Interval > 0 Then timPlaySound.Enabled = True
        timPlaySound_Timer
    End If
End Sub

Private Sub lblAlarm_Click()
    If fraSetup.Visible Then GetFont lblAlarm, lblAlarm.FontName, lblAlarm.FontSize, lblAlarm.FontBold, lblAlarm.FontItalic, lblAlarm.FontUnderline, lblAlarm.FontStrikethru, lblAlarm.ForeColor
End Sub

Private Sub optPlaySound_Click(Index As Integer)
    If Index = 1 Then
        Dim sWAVFile$
        sWAVFile = GetFileName(True, "Wav files (*.wav)|*.wav|All files (*.*)|*.*", "", "Select sound file...")
        If sWAVFile = "" Then
            optPlaySound(0).Value = True
            Exit Sub
        End If
        optPlaySound(1).Tag = sWAVFile
        optPlaySound(1).Caption = "Play wav file [" & LCase(Dir(sWAVFile)) & "]"
    End If
End Sub

Private Sub timPlaySound_Timer()
    If InStr(sSoundFile, "\") > 0 Then
        PlaySound sSoundFile, ByVal 0, SND_ASYNC Or SND_FILENAME
    Else
        PlaySound sSoundFile, ByVal 0, SND_ASYNC Or SND_ALIAS
    End If
End Sub
