VERSION 5.00
Begin VB.Form Graph_1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "示波器"
   ClientHeight    =   3090
   ClientLeft      =   210
   ClientTop       =   4920
   ClientWidth     =   6885
   Icon            =   "Graph_1.frx":0000
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
Attribute VB_Name = "Graph_1"
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
Dim StartRow As Long, ReplayRows As Long, TotalRow As Long

Private Sub Close_Click()
    On Error Resume Next
    Unload msgdis
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
    rtimes = 0
    flag0 = 0
    go_step = 0
    'setup = 1
    setup = 0
    StartRow = 0
    ReplayRows = 0
     
    SelTbl = mapinfo.eval("selectionInfo(1)")
    If Right(SelTbl, 1) = "f" Then
       MessageIndex = 3
    End If
    strx = SelTbl + ".lon"
    stry = SelTbl + ".lat"
    
    For i = 0 To 8
        static_Value(i) = 0
    Next i

    OldPosX = -2
    OldRxLevn1 = 0
    OldRxLevn2 = 0

    old_ci = ""   'mapinfo.eval("Selection.CI_SERV")
    
    StartRow = Val(mapinfo.eval("selectionInfo(3)"))  ' SEL_INFO_NROWS
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
           mapid = mapinfo.eval("windowid(" & i & ")")
           If mapid = mapinfo.eval("frontwindow()") Then
              Exit For
           End If
        End If
    Next
    If StartRow <> 0 Then
       TotalRow = Val(mapinfo.eval("tableinfo(" & SelTbl & ",8)"))
       StartRow = Val(mapinfo.eval("searchpoint(" & mapid & ",selection.lon,selection.lat)"))
       StartRow = Val(mapinfo.eval("SearchInfo(1, 2)"))
       mapinfo.do "Fetch Rec " & StartRow & " FROM " & SelTbl
    End If

    If Replay_flag = 1 Or Replay_flag = 2 Then
       msg = LCase(mapinfo.eval(SelTbl + ".message"))
       While msg <> rmsg1
                    mapinfo.do "Fetch next from " & SelTbl
                    msg = LCase(mapinfo.eval(SelTbl + ".message"))
                    StartRow = StartRow + 1
                    If StartRow = TotalRow Then
                       MsgBox "No This message!"
                       Unload msgdis
                       Unload Me
                       Screen.MousePointer = 0
'******                       SUB_532.Enabled = 0
                    End If

                    If msg = LCase(rmsg1) Then
                       GoTo abc
                    End If

       Wend
     End If
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
End Sub

Private Sub step_Click()
    On Error Resume Next
        Play_graph
        go_step = 1
        Screen.MousePointer = 0
End Sub

Private Sub Stop_Click()
    On Error Resume Next
'    mapinfo.do "set map redraw off"
'    mapinfo.do "delete  from cosmetic1 "
'    mapinfo.do "set map redraw on"

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
End Sub

Private Sub Timer1_Timer()
    
    On Error Resume Next
    If flag0 = 1 Then
       Play_graph
       Play_graph
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
    
    On Error Resume Next
NextRow:
    If ReplayRows >= TotalRow Then
       Screen.MousePointer = 0
       Timer1.Enabled = False
       flag0 = 0
       Exit Sub
    End If
    If flag0 = 1 Or go_step = 1 Then
       If MessageIndex <> 3 Then
          MyMessage = mapinfo.eval(SelTbl + ".MESSAGE")
          If UCase(MyMessage) = "DEDICATED REPORT" Then
             MessageIndex = 1
          ElseIf UCase(MyMessage) = "IDLE MODE REPORT" Then
             MessageIndex = 2
          Else
             ReplayRows = ReplayRows + 1
             mapinfo.do "Fetch next  from " & SelTbl
             GoTo NextRow
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
       msg = mapinfo.eval(SelTbl + ".MESSAGE")
       If msg = "HANDOVER COMPLETE" Then
          My_Picture.DrawWidth = 1
          My_Picture.Line (OldPosX + 1, 0)-(OldPosX + 1, 290), 0
          mapinfo.do "Set Style Symbol MakeSymbol(37,255,24)"
       Else
          If msg = "SETUP" Then
             mapinfo.do "Set Style Symbol MakeSymbol(47,65535,24)"
             setup = 1
          Else
             If msg = "RELEASE COMPLETE" Then
                mapinfo.do "Set Style Symbol MakeSymbol(48,65280,24)"
                setup = 0
             Else
                mapinfo.do "Set Style Symbol MakeSymbol(33,255,4)"
             End If
          End If
       End If
       msg = "Create Point(" & strx & "," & stry & ")"
       On Error Resume Next
       mapinfo.do msg
       Dim ci_str
       ci_str = mapinfo.eval("" & SelTbl & ".CI_SERV")
       If old_ci <> ci_str Then
          Dim k, all As Integer
          Dim ci(3) As String
          old_ci = ci_str
          k = 1
          Is_Found = False
          all = Val(mapinfo.eval("tableinfo(cell,8)"))
          mapinfo.do "Fetch first from cell"
          For k = 1 To all
              If mapinfo.eval("cell.ci") = ci_str Then
                 Is_Found = True
                 msg = mapinfo.eval("cell.cell_name")
                 If InStr(msg, Chr(0)) > 0 Then
                    msg = Trim(Left(msg, InStr(msg, Chr(0)) - 1))
                 End If
                 Exit For
              End If
              mapinfo.do "fetch next from cell"
          Next
          If Is_Found = False Then
             msg = "Unkown"
          End If
       Else
          msg = mapinfo.eval("cell.cell_name")
          If InStr(msg, Chr(0)) > 0 Then
             msg = Trim(Left(msg, InStr(msg, Chr(0)) - 1))
          End If
       End If
       msgdis.BN.Text = msg
       Dim Lev As Integer
       If DisFlag = 0 Then
          msgdis.Text1(0).Text = mapinfo.eval(SelTbl + ".BCCH_SERV")
          msgdis.Text1(1).Text = mapinfo.eval(SelTbl + ".BSIC_SERV")
          msgdis.Text1(2).Text = mapinfo.eval(SelTbl + ".NUM_DCH")
          msgdis.Text1(3).Text = mapinfo.eval(SelTbl + ".TA")
          msgdis.Text1(4).Text = mapinfo.eval(SelTbl + ".Tn_dch")
          Lev = mapinfo.eval(SelTbl + ".RXLEV_F")
          msgdis.Text1(5).Text = Lev
          msgdis.Text1(6).Text = mapinfo.eval(SelTbl + ".RXQUAL_F")
          Lev = mapinfo.eval(SelTbl + ".RXLEV_S")
          msgdis.Text1(7).Text = Lev
          Lev = mapinfo.eval(SelTbl + ".Rxlev_n1")
          msgdis.Text1(8).Text = Lev
          msgdis.Text1(26).Text = mapinfo.eval(SelTbl + ".RXQUAL_S")
          Lev = mapinfo.eval(SelTbl + ".Rxlev_n2")
          msgdis.Text1(9).Text = Lev
          Lev = mapinfo.eval(SelTbl + ".Rxlev_n3")
          msgdis.Text1(10).Text = Lev
          Lev = mapinfo.eval(SelTbl + ".Rxlev_n4")
          msgdis.Text1(11).Text = Lev
          Lev = mapinfo.eval(SelTbl + ".Rxlev_n5")
          msgdis.Text1(12).Text = Lev
          Lev = mapinfo.eval(SelTbl + ".Rxlev_n6")
          msgdis.Text1(13).Text = Lev
          msgdis.Text1(14).Text = mapinfo.eval(SelTbl + ".Bcch_N1")
          msgdis.Text1(15).Text = mapinfo.eval(SelTbl + ".Bcch_N2")
          msgdis.Text1(16).Text = mapinfo.eval(SelTbl + ".Bcch_N3")
          msgdis.Text1(17).Text = mapinfo.eval(SelTbl + ".Bcch_N4")
          msgdis.Text1(18).Text = mapinfo.eval(SelTbl + ".Bcch_N5")
          msgdis.Text1(19).Text = mapinfo.eval(SelTbl + ".Bcch_N6")
          msgdis.Text1(20).Text = mapinfo.eval(SelTbl + ".bsic_n1")
          msgdis.Text1(21).Text = mapinfo.eval(SelTbl + ".bsic_n2")
          msgdis.Text1(22).Text = mapinfo.eval(SelTbl + ".bsic_n3")
          msgdis.Text1(23).Text = mapinfo.eval(SelTbl + ".bsic_n4")
          msgdis.Text1(24).Text = mapinfo.eval(SelTbl + ".bsic_n5")
          msgdis.Text1(25).Text = mapinfo.eval(SelTbl + ".bsic_n6")
          msgdis.Text1(32).Text = mapinfo.eval(SelTbl + ".MESSAGE")
       End If
        If DisFlag = 1 Then
           msgdis.Text2(0).Text = mapinfo.eval(SelTbl + ".MCC_SERV") + "-" + mapinfo.eval(SelTbl + ".MNC_SERV") + "-" + mapinfo.eval(SelTbl + ".LAC_SERV") + "-" + mapinfo.eval(SelTbl + ".CI_SERV")
           msgdis.Text2(1).Text = msg
           msgdis.Text2(2).Text = mapinfo.eval(SelTbl + ".BSIC_SERV")
           msgdis.Text2(3).Text = mapinfo.eval(SelTbl + ".BCCH_SERV")
           msgdis.Text2(4).Text = mapinfo.eval(SelTbl + ".MCC_SERV")
           msgdis.Text2(5).Text = mapinfo.eval(SelTbl + ".MNC_SERV")
           msgdis.Text2(6).Text = mapinfo.eval(SelTbl + ".LAC_SERV")
           msgdis.Text2(7).Text = mapinfo.eval(SelTbl + ".CI_SERV")
           msgdis.Text2(10).Text = mapinfo.eval(SelTbl + ".RXLEV_F")

        End If

        If DisFlag = 2 Then
           msgdis.Text3(0).Text = mapinfo.eval(SelTbl + ".TX_POWER")
           msgdis.Text3(1).Text = mapinfo.eval(SelTbl + ".TA")
           msgdis.Text3(2).Text = mapinfo.eval(SelTbl + ".Act_Rlink")
           msgdis.Text3(3).Text = mapinfo.eval(SelTbl + ".Max_Rlink")
           msgdis.Text3(4).Text = mapinfo.eval(SelTbl + ".RXLEV_F")
           msgdis.Text3(5).Text = mapinfo.eval(SelTbl + ".RXQUAL_F")
           msgdis.Text3(6).Text = mapinfo.eval(SelTbl + ".RXLEV_S")
           msgdis.Text3(7).Text = mapinfo.eval(SelTbl + ".RXQUAL_S")

        End If

        If DisFlag = 3 Then
           msgdis.Text5(0).Text = mapinfo.eval(SelTbl + ".Bcch_N1")
           msgdis.Text5(1).Text = mapinfo.eval(SelTbl + ".Bcch_N2")
           msgdis.Text5(2).Text = mapinfo.eval(SelTbl + ".Bcch_N3")
           msgdis.Text5(3).Text = mapinfo.eval(SelTbl + ".Bcch_N4")
           msgdis.Text5(4).Text = mapinfo.eval(SelTbl + ".Bcch_N5")
           msgdis.Text5(5).Text = mapinfo.eval(SelTbl + ".Bcch_N6")
           msgdis.Text5(6).Text = mapinfo.eval(SelTbl + ".BCCH_SERV")

        End If

           msg = mapinfo.eval(SelTbl + ".MESSAGE")
           If msg = "SETUP" Then
              static_Value(0) = static_Value(0) + 1
           End If
           If msg = "SETUP Failed" Then
              static_Value(1) = static_Value(1) + 1
           End If

           If msg = "HANDOVER COMPLETE" Then
              static_Value(2) = static_Value(2) + 1
           End If
           If msg = "HANDOVER Failed" Then
              static_Value(3) = static_Value(3) + 1
           End If

           If msg = "RELEASE COMPLETE" Then
              static_Value(4) = static_Value(4) + 1
           End If
           If msg = "RELEASE Failed" Then
              static_Value(5) = static_Value(5) + 1
           End If

           If msg = "Location Update access" Then
              static_Value(6) = static_Value(6) + 1
           End If

           If msg = "Location Update Failed" Then
              static_Value(7) = static_Value(7) + 1
           End If

        If DisFlag = 4 Then
           msgdis.Text4(0).Text = static_Value(0)
           msgdis.Text4(1).Text = static_Value(1)
           msgdis.Text4(2).Text = static_Value(2)
           msgdis.Text4(3).Text = static_Value(3)
           msgdis.Text4(4).Text = static_Value(4)
           msgdis.Text4(5).Text = static_Value(5)
           msgdis.Text4(6).Text = static_Value(6)
           msgdis.Text4(7).Text = static_Value(7)
        End If
           
        ReplayRows = ReplayRows + 1
        mapinfo.do "Fetch next  from " & SelTbl
        'mapinfo.Do "fetch next from " & SelTbl
        If Replay_flag = 2 Then
           msg = LCase(mapinfo.eval(SelTbl + ".message"))
           If LCase(msg) = LCase(rmsg2) Then
              rtimes = rtimes + 1
           End If

           If rtimes = Replay_Time Then
              Unload msgdis
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

