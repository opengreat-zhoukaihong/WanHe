VERSION 5.00
Begin VB.Form Graphjam 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "多径衰落与干扰趋势"
   ClientHeight    =   3540
   ClientLeft      =   75
   ClientTop       =   5130
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   Icon            =   "Graphjam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   11820
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "图例说明"
      ForeColor       =   &H8000000A&
      Height          =   2010
      Left            =   9555
      TabIndex        =   18
      Top             =   990
      Width           =   2055
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Rxlev"
         ForeColor       =   &H8000000A&
         Height          =   180
         Index           =   22
         Left            =   630
         TabIndex        =   24
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "True Distance"
         ForeColor       =   &H8000000A&
         Height          =   180
         Index           =   21
         Left            =   630
         TabIndex        =   23
         Top             =   1425
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Timing Advance"
         ForeColor       =   &H8000000A&
         Height          =   180
         Index           =   20
         Left            =   600
         TabIndex        =   22
         Top             =   1125
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "RxQual"
         ForeColor       =   &H8000000A&
         Height          =   180
         Index           =   19
         Left            =   645
         TabIndex        =   21
         Top             =   870
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tx_Power"
         ForeColor       =   &H8000000A&
         Height          =   180
         Index           =   18
         Left            =   630
         TabIndex        =   20
         Top             =   585
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Max_Tx_BTS"
         ForeColor       =   &H8000000A&
         Height          =   180
         Index           =   17
         Left            =   615
         TabIndex        =   19
         Top             =   315
         Width           =   900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFC0C0&
         Index           =   5
         X1              =   240
         X2              =   480
         Y1              =   1755
         Y2              =   1755
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         Index           =   4
         X1              =   240
         X2              =   480
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         Index           =   3
         X1              =   240
         X2              =   480
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FFFF&
         Index           =   2
         X1              =   240
         X2              =   480
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FF00&
         Index           =   1
         X1              =   240
         X2              =   480
         Y1              =   675
         Y2              =   675
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF00FF&
         Index           =   0
         X1              =   240
         X2              =   480
         Y1              =   435
         Y2              =   435
      End
   End
   Begin VB.PictureBox my_Picture 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   480
      ScaleHeight     =   2850
      ScaleWidth      =   8865
      TabIndex        =   16
      Top             =   105
      Width           =   8925
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   0
         X2              =   8970
         Y1              =   2175
         Y2              =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   0
         X2              =   8970
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   0
         X2              =   8970
         Y1              =   320
         Y2              =   320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   0
         X2              =   8970
         Y1              =   640
         Y2              =   640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   -105
         X2              =   8865
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Sub"
      Height          =   240
      Left            =   10635
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Full"
      Height          =   240
      Left            =   9825
      TabIndex        =   14
      Top             =   600
      Value           =   -1  'True
      Width           =   690
   End
   Begin VB.CommandButton Play 
      Caption         =   "Play"
      Height          =   320
      Left            =   765
      TabIndex        =   13
      Top             =   3165
      Width           =   1080
   End
   Begin VB.CommandButton Pause 
      Caption         =   "&Pause"
      Height          =   320
      Left            =   1860
      TabIndex        =   12
      Top             =   3165
      Width           =   1080
   End
   Begin VB.CommandButton Stop 
      Caption         =   "S&top"
      Height          =   320
      Left            =   2940
      TabIndex        =   11
      Top             =   3165
      Width           =   1080
   End
   Begin VB.CommandButton Step 
      Caption         =   "St&ep"
      Height          =   320
      Left            =   4020
      TabIndex        =   10
      Top             =   3165
      Width           =   1080
   End
   Begin VB.CommandButton Close 
      Caption         =   "C&lose"
      Height          =   320
      Left            =   5100
      TabIndex        =   9
      Top             =   3165
      Width           =   1080
   End
   Begin VB.TextBox CName 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   9825
      TabIndex        =   8
      Text            =   "Unkown"
      Top             =   120
      Width           =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5850
      Top             =   3060
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "50"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   25
      Top             =   1590
      Width           =   180
   End
   Begin VB.Line Line2 
      Index           =   9
      X1              =   390
      X2              =   540
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line Line2 
      Index           =   8
      X1              =   390
      X2              =   540
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   390
      X2              =   540
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   390
      X2              =   540
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   390
      X2              =   540
      Y1              =   2625
      Y2              =   2625
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   390
      X2              =   540
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   390
      X2              =   540
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "15"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   15
      Left            =   150
      TabIndex        =   17
      Top             =   45
      Width           =   180
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   390
      X2              =   540
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   390
      X2              =   540
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   405
      X2              =   555
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "5"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   14
      Left            =   240
      TabIndex        =   7
      Top             =   690
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "10"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   6
      Left            =   150
      TabIndex        =   6
      Top             =   360
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "20"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   13
      Left            =   150
      TabIndex        =   5
      Top             =   2235
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "40"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   12
      Left            =   150
      TabIndex        =   4
      Top             =   1905
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "60"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   11
      Left            =   150
      TabIndex        =   3
      Top             =   1260
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   8
      Left            =   240
      TabIndex        =   2
      Top             =   1005
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "10"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   7
      Left            =   150
      TabIndex        =   1
      Top             =   2550
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   5
      Left            =   225
      TabIndex        =   0
      Top             =   2865
      Width           =   90
   End
End
Attribute VB_Name = "Graphjam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag0, go_step, i, j, row, xpos, ypos As Integer
Dim SelTbl, strx, stry, old_ci, CellNo, ci(3) As String
Dim X, Y As Single
Dim old_x(6), Old_y(6), start_rec, static_Value(9) As Integer

Private Sub Close_Click()
    On Error Resume Next
    Screen.MousePointer = 0
    Unload msgdis
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    mapinfo.Do "set map redraw off"
    mapinfo.Do "Set Map Layer 0 Editable On  "
    mapinfo.Do "set map redraw on"
    
    Timer1.Enabled = True
'    MDIMain.SUB_533.Enabled = 0
    flag0 = 0
    go_step = 0
    SelTbl = mapinfo.eval("selectionInfo(1)")
    strx = SelTbl + ".lon"
    stry = SelTbl + ".lat"

    For i = 0 To 8
      static_Value(i) = 0
    Next i

    i = Val(mapinfo.eval("selectionInfo(3)"))  ' SEL_INFO_NROWS
    
    CurrentX = 0
    CurrentY = 0
    old_x(1) = 0
    Old_y(1) = 0
    old_x(2) = 0
    Old_y(2) = 0
    old_x(3) = 0
    Old_y(3) = 0
    old_x(4) = 0
    Old_y(4) = 0
    old_x(5) = 0
    Old_y(5) = 0
    old_x(6) = 0
    Old_y(6) = 0

    old_ci = "99"   'mapinfo.eval("Selection.CI_SERV")
    If i <> 0 Then
       row = Val(mapinfo.eval("tableinfo(" & SelTbl & ",8)"))
       i = mapinfo.eval("searchpoint(" & mapid & ",selection.lon,selection.lat)")
       i = Val(mapinfo.eval("SearchInfo(1, 2)"))
       xpos = 0
       mapinfo.Do "Fetch Rec " & i & " FROM " & SelTbl
    End If
    j = i
    start_rec = i
    xpos = 8888
End Sub

Private Sub Play_graph()
    Dim Is_Found As Boolean
    Dim Cls_Mark As Boolean
    On Error Resume Next
    If j >= row Then
       Screen.MousePointer = 0
       Timer1.Enabled = False
       flag0 = 0
       Exit Sub
    End If
    Cls_Mark = False
    If flag0 = 1 Or go_step = 1 Then
           If xpos >= 8055 Then
              My_Picture.Cls
              Cls_Mark = True
              xpos = 0
              old_x(1) = 0
              old_x(2) = 0
              old_x(3) = 0
              old_x(4) = 0
              old_x(5) = 0
              old_x(6) = 0
           End If

           go_step = 0
           
           strx = SelTbl + ".lon"
           stry = SelTbl + ".lat"
           msg = mapinfo.eval(SelTbl + ".MESSAGE")
           If msg = "HANDOVER COMPLETE" Then
              mapinfo.Do "Set Style Symbol MakeSymbol(37,255,20)"
           Else
              If msg = "SETUP" Then
                 mapinfo.Do "Set Style Symbol MakeSymbol(47,65535,20)"
              Else
                 If msg = "RELEASE COMPLETE" Then
                    My_Picture.Line (xpos, 0)-(xpos, 2860), RGB(25, 150, 10), BF
                    mapinfo.Do "Set Style Symbol MakeSymbol(48,65280,20)"
                 Else
                     mapinfo.Do "Set Style Symbol MakeSymbol(33,255,4)"
                 End If
              End If
           End If
           msg = "Create Point(" & strx & "," & stry & ")"
           mapinfo.Do msg
           
       Dim ci_str
       ci_str = mapinfo.eval("" & SelTbl & ".CI_SERV")
       If old_ci <> ci_str Then
          Dim k, all As Integer
          Dim ci(3) As String
          old_ci = ci_str
          k = 1
          Is_Found = False
          all = Val(mapinfo.eval("tableinfo(cell,8)"))
          mapinfo.Do "Fetch first from cell"
          For k = 1 To all
              If mapinfo.eval("cell.ci") = ci_str Then
                 Is_Found = True
                 CellNo = mapinfo.eval("cell.cell_name")
                 If InStr(CellNo, Chr(0)) > 0 Then
                    CellNo = Trim(Left(CellNo, InStr(CellNo, Chr(0)) - 1))
                 End If
                 Exit For
              End If
              mapinfo.Do "fetch next from cell"
          Next
          If Is_Found = False Then
             CellNo = "Unkown"
          End If
       Else
          CellNo = mapinfo.eval("cell.cell_name")
          If InStr(CellNo, Chr(0)) > 0 Then
             CellNo = Trim(Left(CellNo, InStr(CellNo, Chr(0)) - 1))
          End If
       End If
          
          CName.Text = CellNo
           
           xpos = xpos + 20
           msg = SelTbl + ".Ta"
           ypos = Val(mapinfo.eval(msg))          'TA
           If ypos <> 0 Then
              ypos = 2850 - ypos * 5 * 32
              If ypos < 1210 Then
                 ypos = 1210
              End If
              If Cls_Mark = True Then
                 Old_y(1) = ypos
              End If
              My_Picture.Line (xpos, ypos)-(old_x(1), Old_y(1)), RGB(255, 0, 0)
           Else
              My_Picture.Line (xpos, 2850)-(old_x(1), 2850), RGB(255, 0, 0)
              ypos = 2850
           End If
           old_x(1) = xpos
           Old_y(1) = ypos

           ccc = SelTbl + ".lon"
           ddd = SelTbl + ".lat"
            msg = "x1=distance(cell.lon,cell.lat," & ccc & ", " & ddd & ",""m"")"
            mapinfo.Do msg
           ypos = Val(mapinfo.eval("x1")) * 5 * 32 / 500              'TRUE DIST
                 
              ypos = 2850 - ypos
              If ypos < 1210 Then
                 ypos = 1210
              End If
              If Cls_Mark = True Then
                 Old_y(2) = ypos
              End If
           If cell_no <> "Unkown" Then
              My_Picture.Line (xpos, ypos)-(old_x(2), Old_y(2)), RGB(0, 0, 255)
              old_x(2) = xpos
              Old_y(2) = ypos
           
              msg = "cell.max_tx_bts"                              'MAX_TX_MTS
              i = Val(mapinfo.eval(msg))
              ypos = 2850 - i * 64 - 1890
              If Cls_Mark = True Then
                 Old_y(3) = ypos
              End If
              My_Picture.Line (xpos, ypos)-(old_x(3), Old_y(3)), &HFF00FF
              old_x(3) = xpos
              Old_y(3) = ypos
           End If

           msg = SelTbl + ".Tx_power"                             ' Tx_power
           i = Val(mapinfo.eval(msg))
           ypos = 2850 - i * 64 - 1890
           If Cls_Mark = True Then
              Old_y(4) = ypos
           End If
           My_Picture.Line (xpos, ypos)-(old_x(4), Old_y(4)), RGB(0, 255, 0)
           old_x(4) = xpos
           Old_y(4) = ypos
            
           If Option1.Value = True Then
              msg = SelTbl + ".Rxqual_f"                             ' Rxqual_f
           Else
              msg = SelTbl + ".Rxqual_s"
           End If
           i = Val(mapinfo.eval(msg))
           ypos = 2850 - i * 64 - 1890
           If Cls_Mark = True Then
              Old_y(5) = ypos
           End If
           My_Picture.Line (xpos, ypos)-(old_x(5), Old_y(5)), RGB(255, 255, 0)
           old_x(5) = xpos
           Old_y(5) = ypos
            
           If Option1.Value = True Then
              msg = SelTbl + ".rxlev_f"                       'Rxlev
           Else
              msg = SelTbl + ".rxlev_s"
           End If
           ypos = Val(mapinfo.eval(msg))
           ypos = 2850 - ypos * 32
           If ypos < 1210 Then
              ypos = 1210
           End If
           
           xpos = xpos + 30
           If Cls_Mark = True Then
              Old_y(6) = ypos
           End If
           My_Picture.Line (xpos, ypos)-(old_x(6), Old_y(6)), &HFFC0C0
           old_x(6) = xpos
           Old_y(6) = ypos

           j = j + 1
           
           On Error Resume Next
           mapinfo.Do "Fetch next  from " & SelTbl
    
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    mapinfo.Do "set map redraw off"
    mapinfo.Do "delete  from cosmetic1 "
    mapinfo.Do "Set Map Layer 0 Editable Off  "
    mapinfo.Do "set map redraw on"
    flag0 = 0
End Sub


Private Sub Pause_Click()
    On Error Resume Next
    Screen.MousePointer = 0
    flag0 = 0
End Sub

Private Sub Play_Click()
    On Error Resume Next
    Screen.MousePointer = 11
    flag0 = 1
    j = i
End Sub

Private Sub step_Click()
On Error Resume Next
    Screen.MousePointer = 0
        Play_graph
        go_step = 1
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
    go_step = 0
    CurrentX = 0
    CurrentY = 0
               old_x(1) = 360
               old_x(2) = 360
               old_x(3) = 360
               old_x(4) = 360
               old_x(5) = 360
               old_x(6) = 360
    
    j = start_rec
    Graphjam.Cls
    mapinfo.Do "Fetch Rec " & j & " FROM " & SelTbl
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If flag0 = 1 Then
        Play_graph
    End If
End Sub

