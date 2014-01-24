Attribute VB_Name = "modGraph"
Option Explicit
Private iGraphRAM(1 To 60)
Private iGraphPage(1 To 60)
Private iGraphCPU(1 To 60)
Private lGraphTCPDown(1 To 60) 'Down and up must have
Private lGraphTCPUp(1 To 60)   'same limits!!

Public Sub ClearTCPMonitorData()
    Dim i%
    For i = 1 To 60
        lGraphTCPDown(i) = 0
        lGraphTCPUp(i) = 0
    Next i
End Sub

Public Sub GraphPoint(lValue&, lMax&, iModule%, bSolid As Boolean, objGraph As PictureBox)
    If iModule <> 2 And iModule <> 6 And iModule <> 7 Then Exit Sub
    Dim i%, iVal%, lSpc&
    
    lSpc = Screen.TwipsPerPixelX
    iVal = CInt(lValue / lMax * objGraph.Height)
    If iVal < 20 Then iVal = 20
    
    'Clear picturebox and draw grid
    DrawGrid objGraph
    
    'Shift previous data back one place
    Select Case iModule
        Case MODULE_CPUUSAGE
            For i = UBound(iGraphCPU) - 1 To 2 Step -1
                iGraphCPU(i) = iGraphCPU(i - 1)
            Next i
            iGraphCPU(1) = iVal
        Case MODULE_FREEPAGEFILE
            For i = UBound(iGraphPage) - 1 To 2 Step -1
                iGraphPage(i) = iGraphPage(i - 1)
            Next i
            iGraphPage(1) = iVal
        Case MODULE_FREERAM
            For i = UBound(iGraphRAM) - 1 To 2 Step -1
                iGraphRAM(i) = iGraphRAM(i - 1)
            Next i
            iGraphRAM(1) = iVal
    End Select
    
    'Draw all new values on grid
    'iGraphRAM(1) contains new value, iGraphRAM(100)
    'contains the oldest. Oldest value should be drawn
    'on left side of graph picturebox
    With objGraph
        Select Case iModule
            Case MODULE_CPUUSAGE
                For i = 1 To UBound(iGraphCPU) - 1
                    If iGraphCPU(i) <> 0 Then
                        If Not bSolid Then
                            objGraph.Line (.Width - (i - 1) * lSpc, .Height - iGraphCPU(i))-(.Width - i * lSpc, .Height - iGraphCPU(i + 1)), lColorGraph1st
                        Else
                            objGraph.Line (.Width - i * lSpc, .Height - iGraphCPU(i + 1))-(.Width - i * lSpc, .Height), lColorGraph1st
                        End If
                    End If
                Next i
            Case MODULE_FREEPAGEFILE
                For i = 1 To UBound(iGraphPage) - 1
                    If iGraphPage(i) <> 0 Then
                        If Not bSolid Then
                            objGraph.Line (.Width - (i - 1) * lSpc, .Height - iGraphPage(i))-(.Width - i * lSpc, .Height - iGraphPage(i + 1)), lColorGraph1st
                        Else
                            objGraph.Line (.Width - i * lSpc, .Height - iGraphPage(i + 1))-(.Width - i * lSpc, .Height), lColorGraph1st
                        End If
                    End If
                Next i
            Case MODULE_FREERAM
                For i = 1 To UBound(iGraphRAM) - 1
                    If iGraphRAM(i) <> 0 Then
                        If Not bSolid Then
                            objGraph.Line (.Width - (i - 1) * lSpc, .Height - iGraphRAM(i))-(.Width - i * lSpc, .Height - iGraphRAM(i + 1)), lColorGraph1st
                        Else
                            objGraph.Line (.Width - i * lSpc, .Height - iGraphRAM(i + 1))-(.Width - i * lSpc, .Height), lColorGraph1st
                        End If
                    End If
                Next i
        End Select
    End With
End Sub

Public Sub Graph2PointsDyn(lValDown&, lValUp&, iModule%, bSolid As Boolean, objGraph As PictureBox)
    If iModule <> 20 Then Exit Sub
    Dim i%, iVal1%, iVal2%, lMax&, lSpc&, bUpFirst As Boolean
    lSpc = Screen.TwipsPerPixelX
    'Sub draws two points simultaneously on picturebox
    'and adjusts maximum on the fly - just like DUMeter
    
    'Select Case iModule
    '    Case 20: below code
    '    Case Else: etc
    'End Select
    
    'Clear picturebox & draw grid
    DrawGrid objGraph
    
    'Shift values back one place
    For i = UBound(lGraphTCPDown) - 1 To 2 Step -1
        lGraphTCPDown(i) = lGraphTCPDown(i - 1)
    Next i
    lGraphTCPDown(1) = lValDown
    For i = UBound(lGraphTCPUp) - 1 To 2 Step -1
        lGraphTCPUp(i) = lGraphTCPUp(i - 1)
    Next i
    lGraphTCPUp(1) = lValUp
    
    'Get to max value
    lMax = GetMaxValue(MODULE_TCPMONITOR)
    
    'Draw points on graph, converting to picbox height
    With objGraph
        For i = 1 To UBound(lGraphTCPDown) - 1
            If lGraphTCPDown(i) > lGraphTCPUp(i) Then
                'Down traffic bigger than up traffic,
                'so draw down traffic first to prevent it
                'from overlapping the up traffic
                '(Both in line and solid mode!)
                bUpFirst = False
            Else
                'Up traffic bigger than down traffic,
                'so draw up traffic first
                bUpFirst = True
                'Blegh - GoTo makes messy code
            End If
        
            'Below code draws Downstream traffic
            If bUpFirst Then GoTo DrawUp:
DrawDown:
            If lGraphTCPDown(i) <> 0 Then
                iVal1 = .Height * lGraphTCPDown(i) / lMax
                If iVal1 < 20 Then iVal1 = 20
                iVal2 = .Height * lGraphTCPDown(i + 1) / lMax
                If iVal2 < 20 Then iVal2 = 20
                
                If Not bSolid Then
                    objGraph.Line (.Width - (i - 1) * lSpc, .Height - iVal1)-(.Width - i * lSpc, .Height - iVal2), lColorGraph1st
                    If i > 1 Then
                        If lGraphTCPDown(i - 1) = 0 Then
                            'To prevent half-peaks, draw line from
                            'point i to point i-1, height 0
                            iVal2 = 20
                            objGraph.Line (.Width - (i - 1) * lSpc, .Height - iVal1)-(.Width - (i - 2) * lSpc, .Height - iVal2), lColorGraph1st
                        End If
                    End If
                Else
                    objGraph.Line (.Width - i * lSpc, .Height - iVal1)-(.Width - i * lSpc, .Height), lColorGraph1st
                End If
            End If
            If bUpFirst Then GoTo DrawNext:
        
            'Below code draws Upstream TCP traffic
DrawUp:
            If lGraphTCPUp(i) <> 0 Then
                'iVal1 is new value, iVal2 1st old one
                iVal1 = .Height * lGraphTCPUp(i) / lMax
                If iVal1 < 20 Then iVal1 = 20
                iVal2 = .Height * lGraphTCPUp(i + 1) / lMax
                If iVal2 < 20 Then iVal2 = 20
                
                If Not bSolid Then
                    'Draw line from new to old value
                    objGraph.Line (.Width - (i - 1) * lSpc, .Height - iVal1)-(.Width - i * lSpc, .Height - iVal2), lColorGraph2nd
                    If i > 1 Then
                        If lGraphTCPUp(i - 1) = 0 Then
                            'To prevent half-peaks, draw line from
                            'point i to point i-1, height 0
                            iVal2 = 20
                            objGraph.Line (.Width - (i - 1) * lSpc, .Height - iVal1)-(.Width - (i - 2) * lSpc, .Height - iVal2), lColorGraph2nd
                        End If
                    End If
                Else
                    objGraph.Line (.Width - i * lSpc, .Height - iVal1)-(.Width - i * lSpc, .Height), lColorGraph2nd
                End If
            End If
            If bUpFirst Then GoTo DrawDown:
DrawNext:
        Next i
    End With
End Sub

Public Sub DrawGrid(objGraph As PictureBox)
    objGraph.Cls
    
    With objGraph
        objGraph.Line (0, 0)-(.Width, 0), lColorGraphGrid
        objGraph.Line (0, 120)-(.Width, 120), lColorGraphGrid
        objGraph.Line (0, 240)-(.Width, 240), lColorGraphGrid
        If objGraph.Height > 255 Then
            objGraph.Line (0, 360)-(.Width, 360), lColorGraphGrid
            objGraph.Line (0, 480)-(.Width, 480), lColorGraphGrid
        End If
        
        objGraph.Line (0, 0)-(0, .Height), lColorGraphGrid
        objGraph.Line (120, 0)-(120, .Height), lColorGraphGrid
        objGraph.Line (240, 0)-(240, .Height), lColorGraphGrid
        objGraph.Line (360, 0)-(360, .Height), lColorGraphGrid
        objGraph.Line (480, 0)-(480, .Height), lColorGraphGrid
        objGraph.Line (600, 0)-(600, .Height), lColorGraphGrid
        objGraph.Line (720, 0)-(720, .Height), lColorGraphGrid
        objGraph.Line (840, 0)-(840, .Height), lColorGraphGrid
    End With
End Sub

Private Function GetMaxValue&(iModule%)
    If iModule <> 20 Then Exit Function
    Dim lHighest&, i%
    'Select Case iModule
    '    Case 20: below code
    '    Case Else: etc
    'End Select
    
    For i = 1 To UBound(lGraphTCPDown)
        If lGraphTCPDown(i) > lHighest Then lHighest = lGraphTCPDown(i)
    Next i
    For i = 1 To UBound(lGraphTCPUp)
        If lGraphTCPUp(i) > lHighest Then lHighest = lGraphTCPUp(i)
    Next i
    GetMaxValue = lHighest
End Function
