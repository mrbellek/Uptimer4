VERSION 5.00
Begin VB.Form frmPassword 
   Caption         =   "Enter password"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtPassNew 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtPassOld 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   435
      Index           =   1
      Left            =   1890
      TabIndex        =   8
      Top             =   1650
      Width           =   1275
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   435
      Index           =   0
      Left            =   450
      TabIndex        =   7
      Top             =   1650
      Width           =   1275
   End
   Begin VB.Label lblInfo 
      Caption         =   "New password:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   "Old password:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label lblInfo 
      Caption         =   "Change password for Lock Screen module"
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmMain.imgBanner.Tag = "0"
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Select Case lblInfo(0).Tag
        Case 0 'get old+new password & verify
            If txtPassOld.Text = txtPassOld.Tag Then
                MsgBox "Password changed successfully.", vbInformation, "done"
                frmMain.imgBanner.Tag = txtPassNew.Text
            Else
                MsgBox "The old password is incorrect.", vbExclamation, "change password"
                txtPassOld.SelStart = 0
                txtPassOld.SelLength = Len(txtPassOld.Text)
                txtPassOld.SetFocus
                Exit Sub
            End If
        Case 1 'get password & verify
            If txtPassOld.Text = txtPassOld.Tag Then
                frmMain.imgBanner.Tag = "1"
            Else
                MsgBox "The password is incorrect", vbExclamation, "enter password"
                txtPassOld.SelStart = 0
                txtPassOld.SelLength = Len(txtPassOld.Text)
                txtPassOld.SetFocus
                Exit Sub
            End If
        Case 2 'get password
            frmMain.imgBanner.Tag = txtPassOld.Text
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
    lblSunken(0).Visible = bCoolSunkenButtons
    lblSunken(1).Visible = bCoolSunkenButtons
    SetFormTransparency Me.hWnd
End Sub

Private Sub txtPassNew_GotFocus()
    txtPassNew.SelStart = 0
    txtPassNew.SelLength = Len(txtPassNew.Text)
End Sub
