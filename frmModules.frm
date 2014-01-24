VERSION 5.00
Begin VB.Form frmModules 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modules"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   ControlBox      =   0   'False
   Icon            =   "frmModules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSavedDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdSasvedSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdSavedLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox cboSaved 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2900
      Width           =   1575
   End
   Begin VB.TextBox txtSpacer 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "15"
      Top             =   3510
      Width           =   375
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdDown 
      Height          =   435
      Left            =   3000
      Picture         =   "frmModules.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1260
      Width           =   495
   End
   Begin VB.CommandButton cmdUp 
      Height          =   435
      Left            =   3000
      Picture         =   "frmModules.frx":008D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox lstUsed 
      DragIcon        =   "frmModules.frx":010D
      Height          =   2205
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   2775
   End
   Begin VB.ListBox lstUnused 
      DragIcon        =   "frmModules.frx":0417
      Height          =   2205
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   9
      Left            =   2970
      TabIndex        =   27
      Top             =   1890
      Width           =   555
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   8
      Left            =   2970
      TabIndex        =   26
      Top             =   1230
      Width           =   555
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   7
      Left            =   2970
      TabIndex        =   25
      Top             =   690
      Width           =   555
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   6
      Left            =   2970
      TabIndex        =   24
      Top             =   90
      Width           =   555
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   5
      Left            =   5130
      TabIndex        =   23
      Top             =   3450
      Width           =   1275
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   4
      Left            =   3810
      TabIndex        =   22
      Top             =   3450
      Width           =   1275
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   3
      Left            =   2490
      TabIndex        =   21
      Top             =   3450
      Width           =   1275
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   2
      Left            =   5370
      TabIndex        =   20
      Top             =   2850
      Width           =   1035
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   1
      Left            =   4290
      TabIndex        =   19
      Top             =   2850
      Width           =   1035
   End
   Begin VB.Label lblSunken 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Index           =   0
      Left            =   3210
      TabIndex        =   18
      Top             =   2850
      Width           =   1035
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Saved module sets:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2955
      Width           =   1395
   End
   Begin VB.Label lblSpacer 
      Caption         =   "Horizontal spacer:           pixels"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3540
      Width           =   2295
   End
   Begin VB.Label lblSpaceMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Used: 0000 / 0000 pixels"
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
      Left            =   240
      TabIndex        =   15
      Top             =   2460
      Width           =   6015
   End
   Begin VB.Shape shpSpace 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   330
      Left            =   150
      Top             =   2430
      Width           =   3000
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   6255
   End
End
Attribute VB_Name = "frmModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DoSpaceBar()
    'We need to get the total width the selected modules
    'occupy, and display the text message and set the bar
    
    Dim lSpace&, lSpc&, i%, iRows%
    On Error GoTo Error:
    If bBarDocking Then
        'Max is frmmain.width
        lSpc = CLng(Val(txtSpacer.Text)) * Screen.TwipsPerPixelX
        With frmMain
            lSpace = lSpace + .fraBanner.Width + lSpc
            For i = 0 To lstUsed.ListCount - 1
                Select Case lstUsed.List(i)
                    Case "CD player volume":   lSpace = lSpace + .fraVolume2.Width + lSpc
                    Case "CPU usage":          lSpace = lSpace + .fraCPU.Width + lSpc
                    Case "Date":               lSpace = lSpace + .fraDate.Width + lSpc
                    Case "Disk free space":    lSpace = lSpace + .fraDiskFree.Width + lSpc
                    Case "Exit Windows":       lSpace = lSpace + .fraExitWin.Width + lSpc
                    Case "Free pagefile":      lSpace = lSpace + .fraMemoryPage.Width + lSpc
                    Case "Free RAM":           lSpace = lSpace + .fraMemoryRAM.Width + lSpc
                    Case "IP addresses":       lSpace = lSpace + .fraIPs.Width + lSpc
                    Case "Lock screen":        lSpace = lSpace + .fraLock.Width + lSpc
                    Case "Master volume":      lSpace = lSpace + .fraVolume.Width + lSpc
                    Case "Power status":       lSpace = lSpace + .fraPower.Width + lSpc
                    Case "Screen resolution":  lSpace = lSpace + .fraResolution.Width + lSpc
                    Case "Time":               lSpace = lSpace + .fraTime.Width + lSpc
                    Case "Toggle keys status": lSpace = lSpace + .fraToggle.Width + lSpc
                    Case "Uptime":             lSpace = lSpace + .fraUptime.Width + lSpc
                    Case "WinAmp controls":    lSpace = lSpace + .fraWinamp.Width + lSpc
                    Case "Windows version":    lSpace = lSpace + .fraOS.Width + lSpc
                    'Add more modules below
                    Case "List running processes": lSpace = lSpace + .fraProcesses.Width + lSpc
                    Case "Mouse idle time":    lSpace = lSpace + .fraMouseIdle.Width + lSpc
                    Case "TCP Monitor":        lSpace = lSpace + .fraTCPMonitor.Width + lSpc
                    Case "MSIE version":       lSpace = lSpace + .fraMSIE.Width + lSpc
                    Case "DirectX version":    lSpace = lSpace + .fraDX.Width + lSpc
                    Case "RAS connection":     lSpace = lSpace + .fraRAS.Width
                    Case "Netstat":            lSpace = lSpace + .fraNetstat.Width + lSpc
                End Select
            Next i
            lSpace = lSpace - lSpc
        End With
        If lSpace > frmMain.Width Then
            shpSpace.Width = lblSpace.Width - 45
            If frmMain.imgBanner.Tag <> "recover" And Not bEnableMultiRows Then cmdApply.Enabled = False
            If Not bEnableMultiRows Then cmdOK.Enabled = False
        Else
            shpSpace.Width = lblSpace.Width * (lSpace / frmMain.Width)
            If frmMain.imgBanner.Tag <> "recover" Then cmdApply.Enabled = True
            cmdOK.Enabled = True
        End If
        lblSpaceMsg.Caption = "Used: " & CStr(Int(lSpace / Screen.TwipsPerPixelX)) & " / " & CStr(frmMain.Width / Screen.TwipsPerPixelX) & " pixels"
        If lSpace > frmMain.Width And bEnableMultiRows Then
            iRows = Int(lSpace / frmMain.Width) + 1
            lblSpaceMsg.Caption = lblSpaceMsg.Caption & " (±" & CStr(iRows) & " rows)"
        End If
    Else
        'Max is screen.height - 375 (the titlebar height)
        lSpc = 255 + CLng(Val(txtSpacer.Text)) * Screen.TwipsPerPixelY
        lSpace = (lstUsed.ListCount - 1) * (105 + lSpc)
        If lSpace + 375 > Screen.Height Then
            shpSpace.Width = lblSpace.Width - 45
            If frmMain.imgBanner.Tag <> "recover" Then cmdApply.Enabled = False
            cmdOK.Enabled = False
        Else
            shpSpace.Width = lblSpace.Width * (lSpace / Screen.Height)
            If frmMain.imgBanner.Tag <> "recover" Then cmdApply.Enabled = True
            cmdOK.Enabled = True
        End If
        lblSpaceMsg.Caption = "Used: " & CStr(Int(lSpace / Screen.TwipsPerPixelY)) & " / " & CStr(Screen.Height / Screen.TwipsPerPixelY) & " pixels"
    End If
    Exit Sub
    
Error:
    ShowError "Main", "frmModules.DoSpaceBar", Err.Number, Err.Description, False
End Sub

Private Sub cmdAdd_Click()
    If lstUnused.ListIndex <> -1 Then
        Dim i%
        On Error GoTo Error:
        lstUsed.AddItem lstUnused.List(lstUnused.ListIndex)
        lstUsed.ListIndex = lstUsed.ListCount - 1
        i = lstUnused.ListIndex
        lstUnused.RemoveItem lstUnused.ListIndex
        If i > lstUnused.ListCount - 1 Then
            i = lstUnused.ListCount - 1
        ElseIf i < 0 Then
            i = 0
        End If
        lstUnused.ListIndex = i
        DoSpaceBar
    End If
    Exit Sub
    
Error:
    ShowError "Main", "frmModules.cmdAdd_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmdApply_Click()
    If bBarDocking Then
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSpacingHorz", CLng(Val(txtSpacer.Text)) * Screen.TwipsPerPixelX
    Else
        RegSetDword HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSpacingVert", CLng(Val(txtSpacer.Text)) * Screen.TwipsPerPixelX
    End If
    frmMain.GetModules False 'Save modules
    frmMain.GetModules True  'Re-read modules
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
    
    frmMain.TriggerTimers
End Sub

Private Sub cmdCancel_Click()
    If frmMain.imgBanner.Tag = "recover" Then
        Unload frmMain
        End
    Else
        Unload Me
    End If
End Sub

Private Sub cmdDown_Click()
    If lstUsed.ListIndex <> -1 And lstUsed.ListIndex <> lstUsed.ListCount - 1 Then
        Dim sItem$, i%
        On Error GoTo Error:
        i = lstUsed.ListIndex
        sItem = lstUsed.List(lstUsed.ListIndex)
        lstUsed.RemoveItem i
        If i > lstUsed.ListCount - 1 Then i = lstUsed.ListCount - 1
        lstUsed.AddItem sItem, i + 1
        lstUsed.ListIndex = i + 1
    End If
    Exit Sub
    
Error:
    ShowError "Main", "frmModules.cmdDown_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    If lstUsed.ListIndex <> -1 Then
        Dim i%
        On Error GoTo Error:
        i = lstUnused.ListIndex
        lstUnused.AddItem lstUsed.List(lstUsed.ListIndex)
        lstUnused.ListIndex = i
        i = lstUsed.ListIndex
        If i > lstUsed.ListCount - 2 Then
            i = lstUsed.ListCount - 2
        ElseIf i < 0 Then
            i = 0
        End If
        lstUsed.RemoveItem lstUsed.ListIndex
        lstUsed.ListIndex = i
        DoSpaceBar
    End If
    Exit Sub
    
Error:
    ShowError "Main", "frmModules.cmdRemove_Click", Err.Number, Err.Description, False
End Sub

Private Sub cmdSasvedSave_Click()
    Dim sModSetName$, sDummy$, i%
    sModSetName = InputBox("Enter a name to save the current module set with, e.g. 'Home config'.", "input module set name")
    If sModSetName = "" Then Exit Sub
        
    'Check if name already exists
    For i = 0 To 99
        sDummy = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSet" & IIf(i < 10, "0" & CStr(i), CStr(i)), "blaat")
        If sDummy = "blaat" Then Exit For
        If sModSetName = Left(sDummy, InStr(sDummy, "|") - 1) Then
                
            'Name exists, overwrite prompt
            If MsgBox("A module set with that name already exists. Overwrite it?", vbQuestion + vbYesNo, "overwrite?") = vbNo Then
                Exit Sub
            Else
                Exit For
            End If
        End If
    Next i
    
    'Save module set with set name
    SavedModSets 2, sModSetName, i
    If sDummy <> "blaat" Then
        If sModSetName <> Left(sDummy, InStr(sDummy, "|") - 1) Then cboSaved.ListIndex = cboSaved.ListCount - 1
    End If
End Sub

Private Sub cmdSavedDel_Click()
    If cboSaved.List(cboSaved.ListIndex) = "<empty>" Then Exit Sub
    If MsgBox("Are you sure you want to delete the module set '" & cboSaved.List(cboSaved.ListIndex) & "'?", vbQuestion + vbYesNo, "delete module set") = vbYes Then
        SavedModSets 3, cboSaved.List(cboSaved.ListIndex)
    End If
End Sub

Private Sub cmdSavedLoad_Click()
    If cboSaved.List(cboSaved.ListIndex) = "<empty>" Then Exit Sub
    SavedModSets 1, cboSaved.List(cboSaved.ListIndex)
End Sub

Private Sub cmdUp_Click()
    If lstUsed.ListIndex > 0 Then
        Dim sItem$, i%
        On Error GoTo Error:
        sItem = lstUsed.List(lstUsed.ListIndex)
        lstUsed.AddItem sItem, lstUsed.ListIndex - 1
        i = lstUsed.ListIndex - 2
        lstUsed.RemoveItem lstUsed.ListIndex
        lstUsed.ListIndex = i
    End If
    Exit Sub
    
Error:
    ShowError "Main", "frmModules.cmdUp_Click", Err.Number, Err.Description, False
End Sub

Private Sub Form_Load()
    Dim sDummy$
    sDummy = "," & sModules & ","
    lstUnused.Clear
    lstUsed.Clear
    lblSunken(0).Visible = bCoolSunkenButtons
    lblSunken(1).Visible = bCoolSunkenButtons
    lblSunken(2).Visible = bCoolSunkenButtons
    lblSunken(3).Visible = bCoolSunkenButtons
    lblSunken(4).Visible = bCoolSunkenButtons
    lblSunken(5).Visible = bCoolSunkenButtons
    lblSunken(6).Visible = bCoolSunkenButtons
    lblSunken(7).Visible = bCoolSunkenButtons
    lblSunken(8).Visible = bCoolSunkenButtons
    lblSunken(9).Visible = bCoolSunkenButtons
    SetFormTransparency Me.hwnd
    
    'If InStr(sDummy, ",1,") = 0 Then lstUnused.AddItem "CD player volume"
    If InStr(sDummy, ",2,") = 0 Then lstUnused.AddItem "CPU usage"
    If InStr(sDummy, ",3,") = 0 Then lstUnused.AddItem "Date"
    If InStr(sDummy, ",4,") = 0 Then lstUnused.AddItem "Disk free space"
    If InStr(sDummy, ",5,") = 0 Then lstUnused.AddItem "Exit Windows"
    If InStr(sDummy, ",6,") = 0 Then lstUnused.AddItem "Free pagefile"
    If InStr(sDummy, ",7,") = 0 Then lstUnused.AddItem "Free RAM"
    If InStr(sDummy, ",8,") = 0 Then lstUnused.AddItem "IP addresses"
    If InStr(sDummy, ",9,") = 0 Then lstUnused.AddItem "Lock screen"
    If InStr(sDummy, ",10,") = 0 Then lstUnused.AddItem "Master volume"
    If InStr(sDummy, ",11,") = 0 Then lstUnused.AddItem "Power status"
    If InStr(sDummy, ",12,") = 0 Then lstUnused.AddItem "Screen resolution"
    If InStr(sDummy, ",13,") = 0 Then lstUnused.AddItem "Time"
    If InStr(sDummy, ",14,") = 0 Then lstUnused.AddItem "Toggle keys status"
    If InStr(sDummy, ",15,") = 0 Then lstUnused.AddItem "Uptime"
    If InStr(sDummy, ",16,") = 0 Then lstUnused.AddItem "WinAmp controls"
    If InStr(sDummy, ",17,") = 0 Then lstUnused.AddItem "Windows version"
    'Add more modules below...
    If InStr(sDummy, ",18,") = 0 Then lstUnused.AddItem "List running processes"
    If InStr(sDummy, ",19,") = 0 Then lstUnused.AddItem "Mouse idle time"
    If InStr(sDummy, ",20,") = 0 Then lstUnused.AddItem "TCP Monitor"
    If InStr(sDummy, ",21,") = 0 Then lstUnused.AddItem "MSIE version"
    If InStr(sDummy, ",22,") = 0 Then lstUnused.AddItem "DirectX version"
    If InStr(sDummy, ",23,") = 0 Then lstUnused.AddItem "RAS connection"
    If InStr(sDummy, ",24,") = 0 Then lstUnused.AddItem "Netstat"
    
    Dim vDisplay As Variant, i%
    vDisplay = Split(sModules, ",")
    For i = 0 To UBound(vDisplay)
        Select Case vDisplay(i)
            Case 1:  lstUsed.AddItem "CD player volume"
            Case 2:  lstUsed.AddItem "CPU usage"
            Case 3:  lstUsed.AddItem "Date"
            Case 4:  lstUsed.AddItem "Disk free space"
            Case 5:  lstUsed.AddItem "Exit Windows"
            Case 6:  lstUsed.AddItem "Free pagefile"
            Case 7:  lstUsed.AddItem "Free RAM"
            Case 8:  lstUsed.AddItem "IP addresses"
            Case 9:  lstUsed.AddItem "Lock screen"
            Case 10: lstUsed.AddItem "Master volume"
            Case 11: lstUsed.AddItem "Power status"
            Case 12: lstUsed.AddItem "Screen resolution"
            Case 13: lstUsed.AddItem "Time"
            Case 14: lstUsed.AddItem "Toggle keys status"
            Case 15: lstUsed.AddItem "Uptime"
            Case 16: lstUsed.AddItem "WinAmp controls"
            Case 17: lstUsed.AddItem "Windows version"
            'Add more modules below
            Case 18: lstUsed.AddItem "List running processes"
            Case 19: lstUsed.AddItem "Mouse idle time"
            Case 20: lstUsed.AddItem "TCP Monitor"
            Case 21: lstUsed.AddItem "MSIE version"
            Case 22: lstUsed.AddItem "DirectX version"
            Case 23: lstUsed.AddItem "RAS connection"
            Case 24: lstUsed.AddItem "Netstat"
        End Select
    Next i
    If bBarDocking Then
        txtSpacer.Text = CStr(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSpacingHorz", 105) \ Screen.TwipsPerPixelY)
        lblSpacer.Caption = "Horizontal spacer:           pixels"
    Else
        txtSpacer.Text = CStr(RegGetDword(HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSpacingVert", 225) \ Screen.TwipsPerPixelX)
        lblSpacer.Caption = "    Vertical spacer:           pixels"
    End If
    DoSpaceBar
    If frmMain.imgBanner.Tag = "recover" Then
        cmdApply.Enabled = False
        cmdCancel.Enabled = False
    End If
    
    SavedModSets 0
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
End Sub

Private Sub txtSpacer_Change()
    DoSpaceBar
End Sub

Private Sub SavedModSets(iAction%, Optional sName$, Optional iPos%)
    'iAction:
    ' 0 = Load all to combobox
    ' 1 = Load
    ' 2 = Save
    ' 3 = Delete
    
    Dim i%, sSetName$(0 To 99)
    If iAction = 0 Then 'Load module sets to combobox
        cboSaved.Clear
        For i = 0 To 99
            
            'Get set in format "MySet|10,1,5,7,2"
            sSetName(i) = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSet" & IIf(i < 10, "0" & CStr(i), CStr(i)), "blaat")
            If sSetName(i) = "blaat" Then Exit For
            
            'Add set name to combobox
            cboSaved.AddItem Left(sSetName(i), InStr(sSetName(i), "|") - 1)
        Next i
        If cboSaved.ListCount = 0 Then cboSaved.AddItem "<empty>"
        cboSaved.ListIndex = 0
        Exit Sub
    End If
    
    
    If iAction = 1 Then 'Load selected module set
        
        'Get module set from name
        For i = 0 To 99
            sSetName(0) = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSet" & IIf(i < 10, "0", "") & CStr(i), "blaat")
            If sSetName(0) = "blaat" Then Exit For
            If sName = Left(sSetName(0), InStr(sSetName(0), "|") - 1) Then
                Exit For
            End If
        Next i
        
        If sSetName(0) = "blaat" Then
            'O dear, unable to get module set from name
            MsgBox "Unable to load module set '" & sName & "'!", vbExclamation, "oops"
            Exit Sub
        End If
        
        'Get module set from Registry string and apply
        i = InStr(sSetName(0), "[") - InStr(sSetName(0), "|") - 1
        sSetName(1) = sModules
        sModules = Mid(sSetName(0), InStr(sSetName(0), "|") + 1, i)
        Form_Load
        txtSpacer.Text = CStr(Val(Mid(sSetName(0), InStr(sSetName(0), "[") + 1)))
        sModules = sSetName(1)
        Exit Sub
    End If
    
    
    If iAction = 2 Then 'Save current modules to module set
        
        'Get module set from listbox
        Dim sModSet$
        For i = 0 To lstUsed.ListCount - 1
            Select Case lstUsed.List(i)
                Case "CD player volume":       sModSet = sModSet & "1,"
                Case "CPU usage":              sModSet = sModSet & "2,"
                Case "Date":                   sModSet = sModSet & "3,"
                Case "Disk free space":        sModSet = sModSet & "4,"
                Case "Exit Windows":           sModSet = sModSet & "5,"
                Case "Free pagefile":          sModSet = sModSet & "6,"
                Case "Free RAM":               sModSet = sModSet & "7,"
                Case "IP addresses":           sModSet = sModSet & "8,"
                Case "Lock screen":            sModSet = sModSet & "9,"
                Case "Master volume":          sModSet = sModSet & "10,"
                Case "Power status":           sModSet = sModSet & "11,"
                Case "Screen resolution":      sModSet = sModSet & "12,"
                Case "Time":                   sModSet = sModSet & "13,"
                Case "Toggle keys status":     sModSet = sModSet & "14,"
                Case "Uptime":                 sModSet = sModSet & "15,"
                Case "WinAmp controls":        sModSet = sModSet & "16,"
                Case "Windows version":        sModSet = sModSet & "17,"
                'Add more modules below
                Case "List running processes": sModSet = sModSet & "18,"
                Case "Mouse idle time":        sModSet = sModSet & "19,"
                Case "TCP Monitor":            sModSet = sModSet & "20,"
                Case "MSIE version":           sModSet = sModSet & "21,"
                Case "DirectX version":        sModSet = sModSet & "22,"
                Case "RAS connection":         sModSet = sModSet & "23,"
                Case "Netstat":                sModSet = sModSet & "24,"
            End Select
        Next i
        sModSet = Left(sModSet, Len(sModSet) - 1)
        
        'Save module set
        RegSetString HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSet" & IIf(iPos < 10, "0" & CStr(iPos), CStr(iPos)), sName & "|" & sModSet & "[" & CStr(txtSpacer.Text) & "]"
        
        'Update combo box
        SavedModSets 0
        Exit Sub
    End If
    
    If iAction = 3 Then 'Delete currently selected module set
        
        'Get Registry position of currently selected module set
        For i = 0 To 99
            sSetName(0) = RegGetString(HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSet" & IIf(i < 10, "0" & CStr(i), CStr(i)), "blaat")
            If Left(sSetName(0), InStr(sSetName(0), "|") - 1) = cboSaved.List(cboSaved.ListIndex) Then Exit For
        Next i
        
        'If i = 99 then not found
        If i = 99 Then
            MsgBox "Unable to find Registry key for module set '" & cboSaved.List(cboSaved.ListIndex) & "'!", vbExclamation, "oops"
            Exit Sub
        End If
        
        'Delete module set from Registry
        RegDelValue HKEY_LOCAL_MACHINE, sKeySettings, "ModuleSet" & IIf(i < 10, "0", "") & CStr(i)
        SavedModSets 0
    End If
End Sub
