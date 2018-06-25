VERSION 5.00
Begin VB.Form Graph_Call 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "示波器"
   ClientHeight    =   3090
   ClientLeft      =   210
   ClientTop       =   4920
   ClientWidth     =   6885
   Icon            =   "Graph_Call.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3315
      Left            =   75
      TabIndex        =   6
      Top             =   -570
      Width           =   810
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "RxLev 60"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   11
         Top             =   825
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "RxLev 40"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   10
         Top             =   1365
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "RxLev 20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   9
         Top             =   1905
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "RxLev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   2445
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "RxQual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   225
         Index           =   4
         Left            =   165
         TabIndex        =   7
         Top             =   3015
         Width           =   615
      End
   End
   Begin VB.PictureBox My_Picture 
      BackColor       =   &H00FFFFFF&
      Height          =   2610
      Left            =   975
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   393
      TabIndex        =   5
      Top             =   0
      Width           =   5955
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   -5
         X2              =   403
         Y1              =   131
         Y2              =   131
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   0
         X2              =   408
         Y1              =   58
         Y2              =   58
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   0
         X2              =   408
         Y1              =   22
         Y2              =   22
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   0
         X2              =   408
         Y1              =   94
         Y2              =   94
      End
   End
   Begin VB.CommandButton Play 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1230
      TabIndex        =   4
      Top             =   2715
      Width           =   1080
   End
   Begin VB.CommandButton Pause 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2310
      TabIndex        =   3
      Top             =   2715
      Width           =   1080
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3390
      TabIndex        =   2
      Top             =   2715
      Width           =   1080
   End
   Begin VB.CommandButton Step 
      Caption         =   "Step"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4470
      TabIndex        =   1
      Top             =   2715
      Width           =   1080
   End
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5550
      TabIndex        =   0
      Top             =   2715
      Width           =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6510
      Top             =   90
   End
End
Attribute VB_Name = "Graph_Call"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag0 As Integer, go_step As Integer
Dim SelTbl, strx, stry, old_ci, CellNo As String
Dim X, Y, rtimes As Single
Dim setup, start_rec, static_Value(9) As Integer
'************************************************************************
Dim OldPosX As Integer, OldRxLevn1 As Integer, OldRxLevn2 As Integer
Dim MessageIndex As Integer
Dim MyMessageIndex As Integer
Dim StartRow As Long, ReplayRows As Long, TotalRow As Long
Dim MyLayer3(50) As MessageType
Dim MyCallEvent(50) As MessageType
Dim MyLayer3Count As Integer, MyCallEventCount As Integer
Dim IsOldData As Boolean
Dim MyOldBcch As Integer, MyNewBcch As Integer

Private Sub Close_Click()
    On Error Resume Next

    Unload Me
    Screen.MousePointer = 0
'   MDIMain.SUB_532.Enabled = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    On Error Resume Next
    mapinfo.do "set map redraw off"
    mapinfo.do "Set Map Layer 0 Editable On  "
    mapinfo.do "set map redraw on"
    
    mapinfo.do " reload Custom Symbols From " + Chr(34) + Gsm_Path + "\mysymb" + Chr(34)
    Back_Sel = 1
    rtimes = 0
    flag0 = 0
    go_step = 0
    setup = 0
    StartRow = 0
    ReplayRows = 0
    MyLayer3Count = 0
    MyCallEventCount = 0
    If mapinfo.eval("tableinfo(" & tblname & ",4)") <> 150 Then
        IsOldData = True
    End If
    SelTbl = tblname
    If Right(SelTbl, 1) = "f" Then
       MyMessageIndex = 3
    Else
       MyMessageIndex = 1
    End If
    MessageIndex = 2
    strx = SelTbl + ".lon"
    stry = SelTbl + ".lat"
    For i = 0 To 8
        static_Value(i) = 0
    Next i

    OldPosX = -2
    OldRxLevn1 = 0
    OldRxLevn2 = 0

    old_ci = ""   'mapinfo.eval("Selection.CI_SERV")
    
    StartRow = 1
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
           mapid = mapinfo.eval("windowid(" & i & ")")
           If mapid = mapinfo.eval("frontwindow()") Then
              Exit For
           End If
        End If
    Next
    mapinfo.do "Fetch Rec " & StartRow & " FROM " & SelTbl

abc:
    ReplayRows = StartRow
    start_rec = StartRow
    DrawWidth = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mapinfo.do "set map redraw off"
    mapinfo.do "delete  from cosmetic1 "
    mapinfo.do "Set Map Layer 0 Editable Off  "
    mapinfo.do "set map redraw on"
    Unload FrmLayer3
    Unload frmCallEvent
    flag0 = 0
    Screen.MousePointer = 0
End Sub

Private Sub My_Picture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
       PopupMenu MDIMain.MnuGraphyControl
    End If
End Sub

Private Sub Pause_Click()
    On Error Resume Next
    flag0 = 0
    Screen.MousePointer = 0
End Sub

Private Sub Play_Click()
    On Error Resume Next
    Screen.MousePointer = 11
    flag0 = 1
    ReplayRows = StartRow
    Timer1.Enabled = True
End Sub

Private Sub step_Click()
    On Error Resume Next
        Play_graph
        go_step = 1
        Screen.MousePointer = 0
End Sub

Private Sub Stop_Click()
    On Error Resume Next

     Screen.MousePointer = 0
    
     Dim h As Integer
     For h = 1 To 8
            static_Value(h) = 0
     Next h

    flag0 = 0
    OldPosX = -2
    OldRxLevn1 = 0
    OldRxLevn2 = 0
    
    ReplayRows = start_rec
    On Error Resume Next
    mapinfo.do "Fetch Rec " & ReplayRows & " FROM " & SelTbl
    My_Picture.Cls
    MyCallEventCount = 0
    MyLayer3Count = 0
End Sub

Private Sub Timer1_Timer()
    
    On Error Resume Next
    If flag0 = 1 Then
       Play_graph
     '  Play_graph
    End If
End Sub

Private Sub Play_graph()
    Dim Cls_Mark As Boolean
    Dim Is_Found As Boolean
    Dim MyValue As Integer
    Dim LineColor As Long
    Dim NcellRxlev1 As Integer, NcellRxlev2 As Integer
    Dim ColorFlag As Integer, MaxFlag As Integer
    Dim NcellColor1 As Long, NcellColor2 As Long
    Dim i As Integer
    Dim MyMessage As String
    Dim SendResult As Long
    Dim MyMark As String
    Dim IsAdd As Boolean
    Dim MystrTemp As String
    
    On Error Resume Next
NextRow:
    If mapinfo.eval("EOT(" & tblname & ")") = "T" Then
       Screen.MousePointer = 0
       Timer1.Enabled = False
       flag0 = 0
       Exit Sub
    End If
    If MyNewBcch <> mapinfo.eval(SelTbl + ".bcch_serv") Then
        MyOldBcch = MyNewBcch
        MyNewBcch = mapinfo.eval(SelTbl + ".bcch_serv")
    End If
    If flag0 = 1 Or go_step = 1 Then
       If MyMessageIndex <> 3 Then
          MyMessage = mapinfo.eval(SelTbl + ".MESSAGE")
          If Not IsOldData Then
             MyMark = mapinfo.eval(SelTbl + ".mark1")
          End If
          If UCase(MyMessage) = "DEDICATED REPORT" Then
             MessageIndex = 1
          ElseIf UCase(MyMessage) = "IDLE MODE REPORT" Then
             MessageIndex = 2
          Else
             If UCase(MyMessage) <> "HEADER" And UCase(MyMessage) <> "UNKNOWN" And UCase(MyMessage) <> "END-OF-FILE MARKER" Then
                 If MyLayer3Count < 50 Then
                    MyLayer3Count = MyLayer3Count + 1
                 Else
                    For i = 1 To 49
                        MyLayer3(i) = MyLayer3(i + 1)
                    Next
                 End If
                MyLayer3(MyLayer3Count).RecordMessage = mapinfo.eval(SelTbl + ".MESSAGE")
                If FrmLayer3.List1.ListCount = 50 Then
                   LockWindowUpdate (FrmLayer3.hWnd)
                   FrmLayer3.List1.RemoveItem 0
                   FrmLayer3.List1.AddItem "  " & MyLayer3(MyLayer3Count).RecordMessage, 49
                   FrmLayer3.List1.ListIndex = 49
                   SendResult = SendMessage(FrmLayer3.List1.hWnd, LB_SETANCHORINDEX, 49, ByVal 0&)
                   LockWindowUpdate (0)
                Else
                   FrmLayer3.List1.AddItem "  " & MyLayer3(MyLayer3Count).RecordMessage
                   SendResult = SendMessage(FrmLayer3.List1.hWnd, LB_SETANCHORINDEX, FrmLayer3.List1.ListCount - 1, ByVal 0&)
                End If
            End If
            IsAdd = False
            If MyMark <> "" Then
              If Left(MyMark, 2) = "CF" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "呼叫建立失败 [建立过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""conn_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,16711680,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf Left(MyMark, 2) = "CA" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                   MystrTemp = Trim(Right(MyMark, Len(MyMark) - 2))
                   MystrTemp = Trim(Left(MystrTemp, InStr(MystrTemp, ",") - 1))
                  MyCallEvent(MyCallEventCount).RecordMessage = "第" & MystrTemp & "个呼叫：建立尝试 [建立过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""conn_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,32768,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf Left(MyMark, 2) = "CS" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "建立通话     [建立过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""good.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,65280,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf Left(MyMark, 3) = "HOS" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "切换成功     [通话过程]"
                  mapinfo.do "Set Style Symbol MakeCustomSymbol(""hand_c.bmp"",19711765,24,0)"
                  'mapinfo.do "Set Style Symbol MakeFontSymbol (121,16711680,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
                    My_Picture.DrawWidth = 1
                    My_Picture.Line (OldPosX + 1, 0)-(OldPosX + 1, 290), 0
          
          'mapinfo.do "fetch Prev from " & SelTbl
          My_Picture.CurrentX = OldPosX - 12
          My_Picture.CurrentY = 10
          My_Picture.ForeColor = &HFF0000
          My_Picture.Print Format(MyOldBcch)
          'mapinfo.do "fetch next from " & SelTbl
          My_Picture.CurrentX = OldPosX + 2
          My_Picture.CurrentY = 23
          My_Picture.ForeColor = &HC000&
          My_Picture.Print Format(MyNewBcch)
                    
              ElseIf Left(MyMark, 3) = "HOF" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "切换失败     [通话过程]"
                  mapinfo.do "Set Style Symbol MakeCustomSymbol(""hand_f.bmp"",19711765,24,0)"
                  IsAdd = True
                    My_Picture.DrawWidth = 1
                    My_Picture.Line (OldPosX + 1, 0)-(OldPosX + 1, 290), &HFF&
              ElseIf Left(MyMark, 5) = "CD 掉话" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "掉话         [释放过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""rele_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,14680288,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf Left(MyMark, 5) = "CD 正常" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "正常释放     [释放过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""rele_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,208,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf mapinfo.eval(SelTbl + ".mark") = "Noisy Call" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "噪音通话     [通话过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""rele_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,16744448,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf mapinfo.eval(SelTbl + ".mark") = "Blocked Call" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "建立拥塞     [建立过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""rele_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,7368959,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              End If
              If IsAdd Then
                  mapinfo.do "Create Point(" & tblname & ".lon," & tblname & ".lat)"
                  If frmCallEvent.List1.ListCount = 50 Then
                     LockWindowUpdate (frmCallEvent.hWnd)
                     frmCallEvent.List1.RemoveItem 0
                     If Left(MyCallEvent(MyCallEventCount).RecordMessage, 1) = "第" Then
                        frmCallEvent.List1.AddItem MyCallEvent(MyCallEventCount).RecordMessage, 49
                     Else
                         frmCallEvent.List1.AddItem "     " & MyCallEvent(MyCallEventCount).RecordMessage, 49
                     End If
                     frmCallEvent.List1.ListIndex = 49
                     SendResult = SendMessage(frmCallEvent.List1.hWnd, LB_SETANCHORINDEX, 49, ByVal 0&)
                     LockWindowUpdate (0)
                  Else
                     If Left(MyCallEvent(MyCallEventCount).RecordMessage, 1) = "第" Then
                        frmCallEvent.List1.AddItem MyCallEvent(MyCallEventCount).RecordMessage
                    Else
                        frmCallEvent.List1.AddItem "     " & MyCallEvent(MyCallEventCount).RecordMessage
                    End If
                     SendResult = SendMessage(frmCallEvent.List1.hWnd, LB_SETANCHORINDEX, frmCallEvent.List1.ListCount - 1, ByVal 0&)
                  End If
              End If
            End If
          
             ReplayRows = ReplayRows + 1
             mapinfo.do "Fetch next  from " & SelTbl
             GoTo NextRow
          End If
     Else
           MyMessage = mapinfo.eval(SelTbl + ".MESSAGE")
          If UCase(MyMessage) = "DEDICATED REPORT" Then
             MessageIndex = 1
          ElseIf UCase(MyMessage) = "IDLE MODE REPORT" Then
             MessageIndex = 2
          End If
           
            If UCase(MyMessage) <> "HEADER" And UCase(MyMessage) <> "UNKNOWN" And UCase(MyMessage) <> "END-OF-FILE MARKER" Then
                 If MyLayer3Count < 50 Then
                    MyLayer3Count = MyLayer3Count + 1
                 Else
                    For i = 1 To 49
                        MyLayer3(i) = MyLayer3(i + 1)
                    Next
                 End If
            
            MyLayer3(MyLayer3Count).RecordMessage = mapinfo.eval(SelTbl + ".MESSAGE")
            If FrmLayer3.List1.ListCount = 50 Then
               LockWindowUpdate (FrmLayer3.hWnd)
               FrmLayer3.List1.RemoveItem 0
               FrmLayer3.List1.AddItem "  " & MyLayer3(MyLayer3Count).RecordMessage, 49
               FrmLayer3.List1.ListIndex = 49
               SendResult = SendMessage(FrmLayer3.List1.hWnd, LB_SETANCHORINDEX, 49, ByVal 0&)
               LockWindowUpdate (0)
            Else
               FrmLayer3.List1.AddItem "  " & MyLayer3(MyLayer3Count).RecordMessage
               SendResult = SendMessage(FrmLayer3.List1.hWnd, LB_SETANCHORINDEX, FrmLayer3.List1.ListCount - 1, ByVal 0&)
            End If
       End If
       
          If Not IsOldData Then
            MyMark = mapinfo.eval(SelTbl + ".mark1")
          End If
          IsAdd = False
            If MyMark <> "" Then
              If Left(MyMark, 2) = "CF" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "呼叫建立失败 [建立过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""conn_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,16711680,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf Left(MyMark, 2) = "CA" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                   MystrTemp = Trim(Right(MyMark, Len(MyMark) - 2))
                   MystrTemp = Trim(Left(MystrTemp, InStr(MystrTemp, ",") - 1))
                  MyCallEvent(MyCallEventCount).RecordMessage = "第" & MystrTemp & "个呼叫：建立尝试 [建立过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""conn_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,32768,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf Left(MyMark, 2) = "CS" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "建立通话     [建立过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""good.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,65280,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf Left(MyMark, 3) = "HOS" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "切换成功     [通话过程]"
                  mapinfo.do "Set Style Symbol MakeCustomSymbol(""hand_c.bmp"",19711765,24,0)"
                  'mapinfo.do "Set Style Symbol MakeFontSymbol (121,16711680,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
                    My_Picture.DrawWidth = 1
                    My_Picture.Line (OldPosX + 1, 0)-(OldPosX + 1, 290), 0
          
          'mapinfo.do "fetch Prev from " & SelTbl
          My_Picture.CurrentX = OldPosX - 12
          My_Picture.CurrentY = 10
          My_Picture.ForeColor = &HFF0000
          My_Picture.Print Format(MyOldBcch)
          'mapinfo.do "fetch next from " & SelTbl
          My_Picture.CurrentX = OldPosX + 2
          My_Picture.CurrentY = 23
          My_Picture.ForeColor = &HC000&
          My_Picture.Print Format(MyNewBcch)
                    
              ElseIf Left(MyMark, 3) = "HOF" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "切换失败     [通话过程]"
                  mapinfo.do "Set Style Symbol MakeCustomSymbol(""hand_f.bmp"",19711765,24,0)"
                  IsAdd = True
                    My_Picture.DrawWidth = 1
                    My_Picture.Line (OldPosX + 1, 0)-(OldPosX + 1, 290), &HFF&
              ElseIf Left(MyMark, 5) = "CD 掉话" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "掉话         [释放过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""rele_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,14680288,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf Left(MyMark, 5) = "CD 正常" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "正常释放     [释放过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""rele_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,208,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf mapinfo.eval(SelTbl + ".mark") = "Noisy Call" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "噪音通话     [通话过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""rele_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,16744448,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              ElseIf mapinfo.eval(SelTbl + ".mark") = "Blocked Call" Then
                   If MyCallEventCount < 50 Then
                      MyCallEventCount = MyCallEventCount + 1
                   Else
                      For i = 1 To 49
                          MyCallEvent(i) = MyCallEvent(i + 1)
                      Next
                   End If
                  MyCallEvent(MyCallEventCount).RecordMessage = "建立拥塞     [建立过程]"
                  'mapinfo.do "Set Style Symbol MakeCustomSymbol(""rele_f.bmp"",19711765,24,0)"
                  mapinfo.do "Set Style Symbol MakeFontSymbol (121,7368959,15,""Monotype Sorts"",0,0)"
                  IsAdd = True
              End If
              If IsAdd Then
                  mapinfo.do "Create Point(" & tblname & ".lon," & tblname & ".lat)"
                  If frmCallEvent.List1.ListCount = 50 Then
                     LockWindowUpdate (frmCallEvent.hWnd)
                     frmCallEvent.List1.RemoveItem 0
                     If Left(MyCallEvent(MyCallEventCount).RecordMessage, 1) = "第" Then
                        frmCallEvent.List1.AddItem MyCallEvent(MyCallEventCount).RecordMessage, 49
                     Else
                         frmCallEvent.List1.AddItem "     " & MyCallEvent(MyCallEventCount).RecordMessage, 49
                     End If
                     frmCallEvent.List1.ListIndex = 49
                     SendResult = SendMessage(frmCallEvent.List1.hWnd, LB_SETANCHORINDEX, 49, ByVal 0&)
                     LockWindowUpdate (0)
                  Else
                     If Left(MyCallEvent(MyCallEventCount).RecordMessage, 1) = "第" Then
                        frmCallEvent.List1.AddItem MyCallEvent(MyCallEventCount).RecordMessage
                    Else
                        frmCallEvent.List1.AddItem "     " & MyCallEvent(MyCallEventCount).RecordMessage
                    End If
                     SendResult = SendMessage(frmCallEvent.List1.hWnd, LB_SETANCHORINDEX, frmCallEvent.List1.ListCount - 1, ByVal 0&)
                  End If
              End If
          End If
     End If
       go_step = 0
       Cls_Mark = False
       If OldPosX >= 397 Then
          Cls_Mark = True
          My_Picture.Cls
          OldPosX = -2
          OldRxLevn1 = 0
          OldRxLevn2 = 0
       Else
          If OldPosX = -2 Then
             Cls_Mark = True
          End If
       End If
       My_Picture.DrawWidth = 2
    
       If MessageIndex = 2 Then
          LineColor = &HC0C0C0
       ElseIf MessageIndex = 1 Then
          If Back_Sel = 0 Then
             LineColor = &H808080
          Else
             LineColor = &H808000 '&H800080
          End If
       Else
          LineColor = &HFFC0C0
       End If
       
       If Back_Sel = 0 Then
          MyValue = Val(mapinfo.eval(SelTbl + ".RXQUAL_F"))
       Else
          MyValue = Val(mapinfo.eval(SelTbl + ".RXQUAL_S"))
       End If
       If MyValue > 0 Then
          My_Picture.Line (OldPosX + 2, 174)-(OldPosX + 2, 174 - MyValue * 5), RGB(255, 128, 64)
       End If
       If Back_Sel = 0 Then
          MyValue = mapinfo.eval(SelTbl + ".RXLEV_F")
          If MyValue > 80 Then
             My_Picture.Line (OldPosX + 2, 131)-(OldPosX + 2, 131 - 80 * 1.8), LineColor
          Else
             My_Picture.Line (OldPosX + 2, 131)-(OldPosX + 2, 131 - MyValue * 1.8), LineColor
          End If
       Else
          MyValue = mapinfo.eval(SelTbl + ".RXLEV_S")
          If MyValue > 80 Then
             My_Picture.Line (OldPosX + 2, 131)-(OldPosX + 2, 131 - 80 * 1.8), LineColor
          Else
             If MyValue = 0 Then
                MyValue = mapinfo.eval(SelTbl + ".RXLEV_F")
                My_Picture.Line (OldPosX + 2, 131)-(OldPosX + 2, 131 - MyValue * 1.8), LineColor
             Else
                My_Picture.Line (OldPosX + 2, 131)-(OldPosX + 2, 131 - MyValue * 1.8), LineColor
             End If
          End If
       End If

        NcellRxlev1 = mapinfo.eval(SelTbl + ".rxlev_n1")
        ColorFlag = 1
        For i = 2 To 6
            If NcellRxlev1 < mapinfo.eval(SelTbl + ".rxlev_n" & Format(i)) Then
               NcellRxlev1 = mapinfo.eval(SelTbl + ".rxlev_n" & Format(i))
               ColorFlag = i
            End If
        Next
        If NcellRxlev1 = 0 Then
           GoTo NonNcell
        End If
        MaxFlag = ColorFlag
        NcellColor1 = GetColor(ColorFlag)
        If MaxFlag = 1 Then
           NcellRxlev2 = mapinfo.eval(SelTbl + ".rxlev_n2")
           ColorFlag = 2
        Else
           NcellRxlev2 = mapinfo.eval(SelTbl + ".rxlev_n1")
           ColorFlag = 1
        End If
        For i = 1 To 6
            If i <> MaxFlag Then
               If NcellRxlev2 < mapinfo.eval(SelTbl + ".rxlev_n" & Format(i)) Then
                  NcellRxlev2 = mapinfo.eval(SelTbl + ".rxlev_n" & Format(i))
                  ColorFlag = i
               End If
            End If
        Next
        NcellColor2 = GetColor(ColorFlag)
        If NcellRxlev1 > 80 Then
           NcellRxlev1 = 80
        End If
        If NcellRxlev2 > 80 Then
           NcellRxlev2 = 80
        End If
        If OldPosX = -2 Or OldRxLevn1 = 0 Then
           My_Picture.Line (OldPosX, 131 - NcellRxlev1 * 1.8)-(OldPosX + 2, 131 - NcellRxlev1 * 1.8), NcellColor1
           If NcellRxlev2 > 0 Then
              My_Picture.Line (OldPosX, 131 - NcellRxlev2 * 1.8)-(OldPosX + 2, 131 - NcellRxlev2 * 1.8), NcellColor2
           End If
        Else
           My_Picture.Line (OldPosX, 131 - OldRxLevn1 * 1.8)-(OldPosX + 2, 131 - NcellRxlev1 * 1.8), NcellColor1
           If NcellRxlev2 > 0 Then
              My_Picture.Line (OldPosX, 131 - OldRxLevn2 * 1.8)-(OldPosX + 2, 131 - NcellRxlev2 * 1.8), NcellColor2
           End If
        End If
NonNcell:

       If setup = 0 Then
          'My_Picture.Line (xpos, ypos)-(old_xx, old_yy), RGB(0, 162, 215)
       Else
          'My_Picture.Line (X, ypos)-(old_xx, old_yy), RGB(150, 150, 200)
          'My_Picture.Line (OldPosX + 1, 0)-(OldPosX + 1, 290), 0
       End If
       
       strx = SelTbl + ".lon"
       stry = SelTbl + ".lat"
       mapinfo.do "Set Style Symbol MakeSymbol(33,255,4)"
       Msg = "Create Point(" & strx & "," & stry & ")"
       On Error Resume Next
       mapinfo.do Msg
       GoTo MyIgnore
       Dim ci_str
       ci_str = mapinfo.eval("" & SelTbl & ".CI_SERV")
           Msg = mapinfo.eval(SelTbl + ".MESSAGE")
           If Msg = "SETUP" Then
              static_Value(0) = static_Value(0) + 1
           End If
           If Msg = "SETUP Failed" Then
              static_Value(1) = static_Value(1) + 1
           End If

           If Msg = "HANDOVER COMPLETE" Then
              static_Value(2) = static_Value(2) + 1
           End If
           If Msg = "HANDOVER Failed" Then
              static_Value(3) = static_Value(3) + 1
           End If

           If Msg = "RELEASE COMPLETE" Then
              static_Value(4) = static_Value(4) + 1
           End If
           If Msg = "RELEASE Failed" Then
              static_Value(5) = static_Value(5) + 1
           End If

           If Msg = "Location Update access" Then
              static_Value(6) = static_Value(6) + 1
           End If

           If Msg = "Location Update Failed" Then
              static_Value(7) = static_Value(7) + 1
           End If
MyIgnore:
           
        ReplayRows = ReplayRows + 1
        mapinfo.do "Fetch next  from " & SelTbl
        'mapinfo.Do "fetch next from " & SelTbl
        If Replay_flag = 2 Then
           Msg = LCase(mapinfo.eval(SelTbl + ".message"))
           If LCase(Msg) = LCase(rmsg2) Then
              rtimes = rtimes + 1
           End If

           If rtimes = Replay_Time Then
              Unload Me
              Screen.MousePointer = 0
'******              SUB_532.Enabled = 0
           End If
         End If
         OldRxLevn1 = NcellRxlev1
         OldRxLevn2 = NcellRxlev2
         OldPosX = OldPosX + 2
    End If
    
End Sub

Function GetColor(ColorIndex As Integer) As Long
    On Error Resume Next
    Select Case ColorIndex
       Case 1
           GetColor = &HFF&      '红色
       Case 2
           GetColor = &HFF00&    '绿色
       Case 3
           GetColor = &HFF0000   '蓝色
       Case 4
           GetColor = &HFFFF&    '黄色
       Case 5
           GetColor = &HFF00FF   '粉红色
       Case 6
           GetColor = &HFFFF00   '天蓝色
       Case 7
           GetColor = &HE2EE00   '天蓝色  Scan
       Case 8
           GetColor = &HFA00&    '绿色    Scan
    End Select
End Function

