VERSION 5.00
Begin VB.Form frmInterval 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set interval"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2130
      TabIndex        =   1
      Text            =   "60"
      Top             =   1170
      Width           =   495
   End
   Begin VB.HScrollBar hscInterval 
      Height          =   255
      LargeChange     =   10
      Left            =   480
      Max             =   60
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   2
      Left            =   2730
      TabIndex        =   11
      Top             =   1650
      Width           =   1155
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   1
      Left            =   1410
      TabIndex        =   10
      Top             =   1650
      Width           =   1155
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   1650
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "60"
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   8
      Top             =   870
      Width           =   180
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   870
      Width           =   90
   End
   Begin VB.Label lblInfo 
      Caption         =   "Update module every                 seconds"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblInfo 
      Caption         =   "Use the slider to set the interval at which the [..] module should be updated, or enter an exact number in the textbox."
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmMenu.picTrayIcon.Tag = "-1"
    Unload Me
End Sub

Private Sub cmdDefault_Click()
    frmMenu.picTrayIcon.Tag = cmdDefault.Tag
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtInterval.Text = "" Or _
      Not IsNumeric(txtInterval.Text) Or _
      Val(txtInterval.Text) < CInt(lblInfo(1).Caption) Or _
      Val(txtInterval.Text) > CInt(lblInfo(2).Caption) Then
        MsgBox "Invalid value entered, try again.", vbExclamation, "oops"
        Exit Sub
    End If
    frmMenu.picTrayIcon.Tag = txtInterval.Text
    Unload Me
End Sub

Private Sub Form_Load()
    lblSunken(0).Visible = bCoolSunkenButtons
    lblSunken(1).Visible = bCoolSunkenButtons
    lblSunken(2).Visible = bCoolSunkenButtons
    SetFormTransparency Me.hWnd
End Sub

Private Sub hscInterval_Change()
    txtInterval.Text = CStr(hscInterval.Value)
End Sub

Private Sub hscInterval_Scroll()
    txtInterval.Text = CStr(hscInterval.Value)
End Sub

Private Sub txtInterval_Change()
    On Error Resume Next
    hscInterval.Value = txtInterval.Text
End Sub
