Attribute VB_Name = "Module1"
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Global lpPrevWndProc As Long

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

Type Position
 X As Long
 Y As Long
End Type
Type Destination
 RD As Long
 UD As Long
 LD As Long
 DD As Long
End Type


Global IFoundit As Boolean

Function OpenDialog(xFilter As String, xOwnerForm As Form) As String
    Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = xOwnerForm.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = xFilter
    'Sample Filter --> "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Rich Text Files (*.rtf)" + Chr$(0) + "*.rtf" + Chr$(0)
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = CurDir
        ofn.lpstrTitle = "Open File"
        ofn.flags = 0
        Dim a
        a = GetOpenFileName(ofn)

        If (a) Then
                OpenDialog = Trim$(ofn.lpstrFile)
        Else
                OpenDialog = ""
        End If
End Function



Sub Sensor(XPosition As Position, DestBuf As Destination, XXWay As Integer)
Dim BFound As Boolean
Dim BlFound As Boolean
Form1.Map.ForeColor = vbRed
 
If XXWay <> 3 Then
 Do While Not BFound
  t = t + 15
  Select Case Form1.Map.Point(XPosition.X + t, XPosition.Y)
   Case vbBlack
    BFound = True: DestBuf.RD = t
   Case vbRed, 16777214
    If XXWay = 1 Then BFound = True: DestBuf.RD = t
   Case vbBlue
    For tt = 1 To t Step 15
     Form1.Shape1.Left = Form1.Shape1.Left + 15
     Form1.Map.PSet (Form1.Shape1.Left + 45, Form1.Shape1.Top + 45), vbRed
     For hh = 1 To 1000
      DoEvents
     Next hh
    Next tt
     IFoundit = True: BFound = True
  End Select
 Loop
End If

If XXWay <> 4 Then
t = 0
BFound = False
 Do While Not BFound
  t = t + 15
  Select Case Form1.Map.Point(XPosition.X, XPosition.Y - t)
   Case vbBlack
    BFound = True: DestBuf.UD = t
   Case vbRed, 16777214
    If XXWay = 2 Then BFound = True: DestBuf.UD = t
   Case vbBlue
    For tt = 1 To t Step 15
     Form1.Shape1.Top = Form1.Shape1.Top - 15
     Form1.Map.PSet (Form1.Shape1.Left + 45, Form1.Shape1.Top + 45), vbRed
     For hh = 1 To 1000
      DoEvents
     Next hh
    Next tt
     IFoundit = True: BFound = True
  End Select
 Loop
End If

If XXWay <> 1 Then
t = 0
BFound = False
 Do While Not BFound
  t = t + 15
  Select Case Form1.Map.Point(XPosition.X - t, XPosition.Y)
   Case vbBlack
    BFound = True: DestBuf.LD = t
   Case vbRed, 16777214
    If XXWay = 3 Then BFound = True: DestBuf.LD = t
   Case vbBlue
    For tt = 1 To t Step 15
     Form1.Shape1.Left = Form1.Shape1.Left - 15
     Form1.Map.PSet (Form1.Shape1.Left + 45, Form1.Shape1.Top + 45), vbRed
     For hh = 1 To 1000
      DoEvents
     Next hh
    Next tt
     IFoundit = True: BFound = True
  End Select
 Loop
End If

If XXWay <> 2 Then
t = 0
BFound = False
 Do While Not BFound
  t = t + 15
  Select Case Form1.Map.Point(XPosition.X, XPosition.Y + t)
   Case vbBlack
    BFound = True: DestBuf.DD = t
   Case vbRed, 16777214
    If XXWay = 4 Then BFound = True: DestBuf.DD = t
   Case vbBlue
    For tt = 1 To t Step 15
     Form1.Shape1.Top = Form1.Shape1.Top + 15
     Form1.Map.PSet (Form1.Shape1.Left + 45, Form1.Shape1.Top + 45), vbRed
     For hh = 1 To 1000
      DoEvents
     Next hh
    Next tt
     IFoundit = True: BFound = True
  End Select
 Loop
End If
End Sub
Sub Xmain(XWay As Integer, StartPos As Position)
Dim BFound As Boolean
Dim Changed As Boolean
Dim DST As Destination
Dim DSTBACK As Destination
Dim XPOS As Position
Sensor StartPos, DSTBACK, 0

 Select Case XWay
  Case 1
   Do While (Not BFound) And (Not IFoundit)
    t = t + 15
    XPOS.X = StartPos.X + t
    XPOS.Y = StartPos.Y
    Sensor XPOS, DST, 1
    
    If DSTBACK.UD >= DST.UD Then
      DSTBACK.UD = DST.UD
     Else
      Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
      Xmain 2, XPOS
      DSTBACK.UD = DST.UD
    End If
    If IFoundit Then Exit Sub
    If DSTBACK.DD >= DST.DD Then
      DSTBACK.DD = DST.DD
     Else
      Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
      Xmain 4, XPOS
      DSTBACK.DD = DST.DD
    End If
    If IFoundit Then Exit Sub
    
    
    Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
    Form1.Shape1.Top = XPOS.Y - 45: Form1.Shape1.Left = XPOS.X - 45
    For h = 1 To 1000
     DoEvents
    Next h
    If DST.RD = 15 Then
     For hh = 1 To XPOS.X / 15 - StartPos.X / 15
      Form1.Map.PSet (Form1.Shape1.Left + 45, Form1.Shape1.Top + 45), 16777214
      Form1.Shape1.Left = Form1.Shape1.Left - 15
      For h = 1 To 1000
       DoEvents
      Next h
     Next hh
     BFound = True
    End If
   Loop
  Case 2
   Do While (Not BFound) And (Not IFoundit)
    t = t + 15
    XPOS.X = StartPos.X
    XPOS.Y = StartPos.Y - t
    Sensor XPOS, DST, 2
    
    If DSTBACK.RD >= DST.RD Then
      DSTBACK.RD = DST.RD
     Else
      Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
      Xmain 1, XPOS
      DSTBACK.RD = DST.RD
    End If
    If IFoundit Then Exit Sub
    If DSTBACK.LD >= DST.LD Then
      DSTBACK.LD = DST.LD
     Else
      Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
      Xmain 3, XPOS
      DSTBACK.LD = DST.LD
    End If
    If IFoundit Then Exit Sub

    Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
    Form1.Shape1.Top = XPOS.Y - 45: Form1.Shape1.Left = XPOS.X - 45
    For h = 1 To 1000
     DoEvents
    Next h
    If DST.UD = 15 Then
     For hh = 1 To StartPos.Y / 15 - XPOS.Y / 15
      Form1.Map.PSet (Form1.Shape1.Left + 45, Form1.Shape1.Top + 45), 16777214
      Form1.Shape1.Top = Form1.Shape1.Top + 15
      For h = 1 To 1000
       DoEvents
      Next h
     Next hh
     BFound = True
    End If
   Loop
  Case 3
   Do While (Not BFound) And (Not IFoundit)
    t = t + 15
    XPOS.X = StartPos.X - t
    XPOS.Y = StartPos.Y
    Sensor XPOS, DST, 3
    
    If DSTBACK.UD >= DST.UD Then
      DSTBACK.UD = DST.UD
     Else
      Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
      Xmain 2, XPOS
      DSTBACK.UD = DST.UD
    End If
    If IFoundit Then Exit Sub
    If DSTBACK.DD >= DST.DD Then
      DSTBACK.DD = DST.DD
     Else
      Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
      Xmain 4, XPOS
      DSTBACK.DD = DST.DD
    End If
    If IFoundit Then Exit Sub

    Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
    Form1.Shape1.Top = XPOS.Y - 45: Form1.Shape1.Left = XPOS.X - 45
    For h = 1 To 1000
     DoEvents
    Next h
    DoEvents
    If DST.LD = 15 Then
     For hh = 1 To StartPos.X / 15 - XPOS.X / 15
      Form1.Map.PSet (Form1.Shape1.Left + 45, Form1.Shape1.Top + 45), 16777214
      Form1.Shape1.Left = Form1.Shape1.Left + 15
      For h = 1 To 1000
       DoEvents
      Next h
     Next hh
     BFound = True
    End If
   Loop
  Case 4
   Do While (Not BFound) And (Not IFoundit)
    t = t + 15
    XPOS.X = StartPos.X
    XPOS.Y = StartPos.Y + t
    Sensor XPOS, DST, 4
    
    If DSTBACK.RD >= DST.RD Then
      DSTBACK.RD = DST.RD
     Else
      Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
      Xmain 1, XPOS
      DSTBACK.RD = DST.RD
    End If
    If IFoundit Then Exit Sub
    If DSTBACK.LD >= DST.LD Then
      DSTBACK.LD = DST.LD
     Else
      Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
      Xmain 3, XPOS
      DSTBACK.LD = DST.LD
    End If
    If IFoundit Then Exit Sub

    Form1.Map.PSet (XPOS.X, XPOS.Y), vbRed
    Form1.Shape1.Top = XPOS.Y - 45: Form1.Shape1.Left = XPOS.X - 45
    For h = 1 To 1000
     DoEvents
    Next h
    If DST.DD = 15 Then
     For hh = 1 To XPOS.Y / 15 - StartPos.Y / 15
      Form1.Map.PSet (Form1.Shape1.Left + 45, Form1.Shape1.Top + 45), 16777214
      Form1.Shape1.Top = Form1.Shape1.Top - 15
      For h = 1 To 1000
       DoEvents
      Next h
     Next hh
     BFound = True
    End If
   Loop
 End Select
End Sub

