VERSION 5.00
Begin VB.Form frmLocked 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Screen is locked"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timMove 
      Interval        =   5000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Screen is locked"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   555
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
   End
End
Attribute VB_Name = "frmLocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    SetWindowPos Me.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    'DoEvents
    'PaintDesktop Me.hdc
    'DoEvents
    lblInfo.Caption = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "LockMsg", "Screen is locked by Uptimer4")
    lblInfo.Left = (Me.Width - lblInfo.Width) / 2
    lblInfo.Top = (Me.Height - lblInfo.Height) / 2
    Me.Show
    Randomize
End Sub

Private Sub timMove_Timer()
    lblInfo.Visible = False
    'DoEvents
    'PaintDesktop Me.hdc
    'DoEvents
    lblInfo.Left = Rnd * (Screen.Width - lblInfo.Width)
    lblInfo.Top = frmMain.Height + Rnd * (Screen.Height - lblInfo.Height - frmMain.Height)
    lblInfo.Visible = True
End Sub
