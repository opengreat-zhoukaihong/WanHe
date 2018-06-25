VERSION 5.00
Begin VB.Form Graph_old 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "·ÅÏñ»ú"
   ClientHeight    =   2925
   ClientLeft      =   225
   ClientTop       =   5895
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2925
   ScaleWidth      =   6990
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3720
      Top             =   0
   End
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Step 
      Caption         =   "Step"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Pause 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Play 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "RxQual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   12
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "RxLev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "6"
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   1680
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   6600
      X2              =   6600
      Y1              =   1920
      Y2              =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   255
   End
   Begin VB.Line Line6 
      BorderStyle     =   3  'Dot
      X1              =   600
      X2              =   6600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      BorderStyle     =   3  'Dot
      X1              =   600
      X2              =   6600
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line4 
      BorderStyle     =   3  'Dot
      X1              =   600
      X2              =   6600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   600
      X2              =   600
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   600
      X2              =   6600
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "Graph_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag0, go_step, i, j, row, xpos, ypos As Integer
Dim SelTbl, strx, stry, old_ci, CellNo As String
Dim x, y, rtimes As Single
Dim old_xx, old_yy, old_x(6), Old_y(6), setup, start_rec, static_Value(9) As Integer

Private Sub Close_Click()
On Error Resume Next
    Unload msgdis
    Unload Me
    Screen.MousePointer = 0
'   MDIMain.SUB_532.Enabled = 0
End Sub
Private Sub Form_Load()
    On Error Resume Next
    MapInfo.Do "set map redraw off"
    MapInfo.Do "Set Map Layer 0 Editable On  "
    MapInfo.Do "set map redraw on"
    
    On Error Resume Next
    MapInfo.Do " reload Custom Symbols From " + Chr(34) + Gsm_Path + "\mysymb" + Chr(34)
    rtimes = 0
    flag0 = 0
    go_step = 0
    setup = 1
    i = 0
    j = 0
     
    SelTbl = MapInfo.eval("selectionInfo(1)")
    strx = SelTbl + ".lon"
    stry = SelTbl + ".lat"
    
    For i = 0 To 8
        static_Value(i) = 0
    Next i

    i = Val(MapInfo.eval("selectionInfo(3)"))  ' SEL_INFO_NROWS
    
    CurrentX = 0
    CurrentY = 0
    old_x(1) = 600
    old_x(2) = 600
    old_x(3) = 600
    old_x(4) = 600
    old_x(5) = 600
    old_x(6) = 600
    Old_y(1) = 1920
    Old_y(2) = 1920
    Old_y(3) = 1920
    Old_y(4) = 1920
    Old_y(5) = 1920
    Old_y(6) = 1920
    old_xx = 600
    old_yy = 1920

    old_ci = "99"   'mapinfo.eval("Selection.CI_SERV")
    If i <> 0 Then
       row = Val(MapInfo.eval("tableinfo(" & SelTbl & ",8)"))
       i = Val(MapInfo.eval("searchpoint(" & mapid & ",selection.lon,selection.lat)"))
       i = Val(MapInfo.eval("SearchInfo(1, 2)"))
       xpos = 600
       MapInfo.Do "Fetch Rec " & i & " FROM " & SelTbl
    End If

    If Replay_flag = 1 Or Replay_flag = 2 Then
       ff = SelTbl + ".message"
       msg = LCase(MapInfo.eval(ff))
       While msg <> rmsg1
                    MapInfo.Do "Fetch next from " & SelTbl
                    msg = LCase(MapInfo.eval(ff))
                    i = i + 1
                    If i = row Then
                       MsgBox "No This message!"
                       Unload msgdis
                       Unload Me
                       Screen.MousePointer = 0
                       SUB_532.Enabled = 0
                    End If

                    If msg = LCase(rmsg1) Then
                       GoTo CM
                    End If

       Wend
     End If
CM:
    j = i
    start_rec = i
    DrawWidth = 2
End Sub

Private Sub Play_graph()
On Error Resume Next
   If j >= row Then
      Screen.MousePointer = 0
      Timer1.Enabled = False
      flag0 = 0
      Exit Sub
   End If

   If flag0 = 1 Or go_step = 1 Then
           go_step = 0
           If Back_Sel = 0 Then
              msg = SelTbl + ".RXLEV_F"
           Else
              msg = SelTbl + ".RXLEV_S"
           End If
           ypos = 1920 - (Val(MapInfo.eval(msg))) * 24
           xpos = xpos + 30
           If xpos >= 6600 Then
              Graph.Cls
              xpos = 600
              old_x(1) = 600
              old_x(2) = 600
              old_x(3) = 600
              old_x(4) = 600
              old_x(5) = 600
              old_x(6) = 600
              Old_y(1) = 1920
              Old_y(2) = 1920
              Old_y(3) = 1920
              Old_y(4) = 1920
              Old_y(5) = 1920
              Old_y(6) = 1920
              old_xx = 600
              old_yy = 1920
           End If
           DrawWidth = 1
           BorderWidth = 1
           If setup = 0 Then
   '           Line (xpos, ypos)-(xpos, 1920), RGB(150, 150, 200)
               Line (xpos, ypos)-(old_xx, old_yy), RGB(150, 150, 200)
           Else
    '          Line (xpos, ypos)-(xpos, 1920), RGB(200, 200, 200)
               Line (xpos, ypos)-(old_xx, old_yy), 0
           End If
           old_xx = xpos
           old_yy = ypos

           msg = SelTbl + ".Rxlev_n1"
           On Error Resume Next
           On Error GoTo 0
           ypos = 1920 - (Val(MapInfo.eval(msg))) * 24
'           PSet (xpos, ypos), RGB(0, 0, 255)
           Line (xpos, ypos)-(old_x(1), Old_y(1)), RGB(255, 0, 0)
           old_x(1) = xpos
           Old_y(1) = ypos
             
             
           msg = SelTbl + ".Rxlev_n2"
           ypos = 1920 - (Val(MapInfo.eval(msg))) * 24
'           PSet (xpos, ypos), RGB(255, 255, 0)
           Line (xpos, ypos)-(old_x(2), Old_y(2)), RGB(0, 255, 0)
           old_x(2) = xpos
           Old_y(2) = ypos
           
           msg = SelTbl + ".Rxlev_n3"
           ypos = 1920 - (Val(MapInfo.eval(msg))) * 24
'           PSet (xpos, ypos), RGB(255, 0, 0)
           Line (xpos, ypos)-(old_x(3), Old_y(3)), RGB(0, 0, 255)
           old_x(3) = xpos
           Old_y(3) = ypos
           
           msg = SelTbl + ".Rxlev_n4"
           ypos = 1920 - (Val(MapInfo.eval(msg))) * 24
'           PSet (xpos, ypos), RGB(255, 0, 0)
           Line (xpos, ypos)-(old_x(4), Old_y(4)), RGB(255, 0, 255)
           old_x(4) = xpos
           Old_y(4) = ypos
           
           msg = SelTbl + ".Rxlev_n5"
           ypos = 1920 - (Val(MapInfo.eval(msg))) * 24
'           PSet (xpos, ypos), RGB(255, 0, 0)
           Line (xpos, ypos)-(old_x(5), Old_y(5)), RGB(255, 255, 0)
           old_x(5) = xpos
           Old_y(5) = ypos
           
           msg = SelTbl + ".Rxlev_n6"
           ypos = 1920 - (Val(MapInfo.eval(msg))) * 24
'           PSet (xpos, ypos), RGB(255, 0, 0)
           Line (xpos, ypos)-(old_x(6), Old_y(6)), RGB(0, 255, 255)
           old_x(6) = xpos
           Old_y(6) = ypos

           If Back_Sel = 0 Then
              msg = SelTbl + ".RXQUAL_F"
           Else
              msg = SelTbl + ".RXQUAL_S"
           End If
           ypos = 1920 - Val(MapInfo.eval(msg)) * 60
'           PSet (xpos, ypos), RGB(255, 0, 255)
'           Line (xpos, ypos)-(xpos, 1920), RGB(200, 0, 255)
           Line (xpos, ypos)-(xpos, 1920), 9464832

           strx = SelTbl + ".lon"
           stry = SelTbl + ".lat"
           msg = MapInfo.eval(SelTbl + ".MESSAGE")
           If msg = "HANDOVER COMPLETE" Then
              Line (xpos, 0)-(xpos, 1920), RGB(25, 100, 10)
'              mapinfo.do " reload Custom Symbols From ""\gsm\mysymb"""
'              msg = "symbol(""hand_c.bmp""" + ",255,18,0)"
'              mapinfo.eval (msg)
              MapInfo.Do "Set Style Symbol MakeSymbol(37,255,24)"

           Else
              If msg = "SETUP" Then
                 MapInfo.Do "Set Style Symbol MakeSymbol(47,65535,24)"
                 setup = 1
              Else
                 If msg = "RELEASE COMPLETE" Then
                    MapInfo.Do "Set Style Symbol MakeSymbol(48,65280,24)"
                    setup = 0
                 Else
                    MapInfo.Do "Set Style Symbol MakeSymbol(33,255,4)"
                 End If
              End If
           End If

           msg = "Create Point(" & strx & "," & stry & ")"
           On Error Resume Next
           MapInfo.Do msg

           Dim ci_str
           ci_str = MapInfo.eval("" & SelTbl & ".CI_SERV")
           If old_ci <> ci_str Then
              Dim k, all As Integer
              Dim ci(3) As String
              
              MapInfo.Do "Fetch first from BASE"
              k = 1
              all = Val(MapInfo.eval("tableinfo(BASE,8)"))
              ci(1) = LCase(MapInfo.eval("BASE.ci_1"))
              ci(2) = LCase(MapInfo.eval("BASE.ci_2"))
              ci(3) = LCase(MapInfo.eval("BASE.ci_3"))
              While ci_str <> ci(1) And ci_str <> ci(2) And ci_str <> ci(3) And k < all
                    MapInfo.Do "Fetch next from BASE"
                    ci(1) = LCase(MapInfo.eval("BASE.ci_1"))
                    ci(2) = LCase(MapInfo.eval("BASE.ci_2"))
                    ci(3) = LCase(MapInfo.eval("BASE.ci_3"))
                    k = k + 1
              Wend
            If ci_str = ci(1) Then
               CellNo = " A"
            Else
               If ci_str = ci(2) Then
                  CellNo = " B"
               Else
               If ci_str = ci(3) Then
                  CellNo = " C"
               End If
               End If
            End If

            If k < all Then
                 old_ci = ci_str
                 msg = MapInfo.eval("Base.bs_name")
                 msg = msg + CellNo
            Else
                 msg = "Unkown"
            End If
         Else
                 msg = MapInfo.eval("Base.bs_name")
                 msg = msg + CellNo
         End If

          
        msgdis.BN.Text = msg
        Dim Lev As Integer
        If DisFlag = 0 Then
           msgdis.Text1(0).Text = MapInfo.eval(SelTbl + ".BCCH_SERV")
           msgdis.Text1(1).Text = MapInfo.eval(SelTbl + ".BSIC_SERV")
           msgdis.Text1(2).Text = MapInfo.eval(SelTbl + ".NUM_DCH")
           msgdis.Text1(3).Text = MapInfo.eval(SelTbl + ".TA")
           msgdis.Text1(4).Text = MapInfo.eval(SelTbl + ".Tn_dch")
           Lev = MapInfo.eval(SelTbl + ".RXLEV_F")
'           If lev = 110 Then
'              lev = 0
'           End If
           msgdis.Text1(5).Text = Lev
           msgdis.Text1(6).Text = MapInfo.eval(SelTbl + ".RXQUAL_F")

           Lev = MapInfo.eval(SelTbl + ".RXLEV_S")
'           If lev = 110 Then
'              lev = 0
'           End If
           msgdis.Text1(7).Text = Lev

           Lev = MapInfo.eval(SelTbl + ".Rxlev_n1")
'           If lev = 110 Then
'              lev = 0
'           End If
           msgdis.Text1(8).Text = Lev
           msgdis.Text1(26).Text = MapInfo.eval(SelTbl + ".RXQUAL_S")

           Lev = MapInfo.eval(SelTbl + ".Rxlev_n2")
'           If lev = 110 Then
'              lev = 0
'           End If
           msgdis.Text1(9).Text = Lev

           Lev = MapInfo.eval(SelTbl + ".Rxlev_n3")
'           If lev = 110 Then
'              lev = 0
'           End If
           msgdis.Text1(10).Text = Lev

           Lev = MapInfo.eval(SelTbl + ".Rxlev_n4")
'           If lev = 110 Then
'              lev = 0
'           End If
           msgdis.Text1(11).Text = Lev

           Lev = MapInfo.eval(SelTbl + ".Rxlev_n5")
'           If lev = 110 Then
'              lev = 0
'           End If
           msgdis.Text1(12).Text = Lev

           Lev = MapInfo.eval(SelTbl + ".Rxlev_n6")
'           If lev = 110 Then
'              lev = 0
'           End If
           msgdis.Text1(13).Text = Lev

           msgdis.Text1(14).Text = MapInfo.eval(SelTbl + ".Bcch_N1")
           msgdis.Text1(15).Text = MapInfo.eval(SelTbl + ".Bcch_N2")
           msgdis.Text1(16).Text = MapInfo.eval(SelTbl + ".Bcch_N3")
           msgdis.Text1(17).Text = MapInfo.eval(SelTbl + ".Bcch_N4")
           msgdis.Text1(18).Text = MapInfo.eval(SelTbl + ".Bcch_N5")
           msgdis.Text1(19).Text = MapInfo.eval(SelTbl + ".Bcch_N6")

           msgdis.Text1(20).Text = MapInfo.eval(SelTbl + ".bsic_n1")
           msgdis.Text1(21).Text = MapInfo.eval(SelTbl + ".bsic_n2")
           msgdis.Text1(22).Text = MapInfo.eval(SelTbl + ".bsic_n3")
           msgdis.Text1(23).Text = MapInfo.eval(SelTbl + ".bsic_n4")
           msgdis.Text1(24).Text = MapInfo.eval(SelTbl + ".bsic_n5")
           msgdis.Text1(25).Text = MapInfo.eval(SelTbl + ".bsic_n6")
           msgdis.Text1(32).Text = MapInfo.eval(SelTbl + ".MESSAGE")
        End If

        If DisFlag = 1 Then
           msgdis.Text2(0).Text = MapInfo.eval(SelTbl + ".MCC_SERV") + "-" + MapInfo.eval(SelTbl + ".MNC_SERV") + "-" + MapInfo.eval(SelTbl + ".LAC_SERV") + "-" + MapInfo.eval(SelTbl + ".CI_SERV")
           msgdis.Text2(1).Text = msg
           msgdis.Text2(2).Text = MapInfo.eval(SelTbl + ".BSIC_SERV")
           msgdis.Text2(3).Text = MapInfo.eval(SelTbl + ".BCCH_SERV")
           msgdis.Text2(4).Text = MapInfo.eval(SelTbl + ".MCC_SERV")
           msgdis.Text2(5).Text = MapInfo.eval(SelTbl + ".MNC_SERV")
           msgdis.Text2(6).Text = MapInfo.eval(SelTbl + ".LAC_SERV")
           msgdis.Text2(7).Text = MapInfo.eval(SelTbl + ".CI_SERV")
           msgdis.Text2(10).Text = MapInfo.eval(SelTbl + ".RXLEV_F")

        End If

        If DisFlag = 2 Then
           msgdis.Text3(0).Text = MapInfo.eval(SelTbl + ".TX_POWER")
           msgdis.Text3(1).Text = MapInfo.eval(SelTbl + ".TA")
           msgdis.Text3(2).Text = MapInfo.eval(SelTbl + ".Act_Rlink")
           msgdis.Text3(3).Text = MapInfo.eval(SelTbl + ".Max_Rlink")
           msgdis.Text3(4).Text = MapInfo.eval(SelTbl + ".RXLEV_F")
           msgdis.Text3(5).Text = MapInfo.eval(SelTbl + ".RXQUAL_F")
           msgdis.Text3(6).Text = MapInfo.eval(SelTbl + ".RXLEV_S")
           msgdis.Text3(7).Text = MapInfo.eval(SelTbl + ".RXQUAL_S")

        End If

        If DisFlag = 3 Then
           msgdis.Text5(0).Text = MapInfo.eval(SelTbl + ".Bcch_N1")
           msgdis.Text5(1).Text = MapInfo.eval(SelTbl + ".Bcch_N2")
           msgdis.Text5(2).Text = MapInfo.eval(SelTbl + ".Bcch_N3")
           msgdis.Text5(3).Text = MapInfo.eval(SelTbl + ".Bcch_N4")
           msgdis.Text5(4).Text = MapInfo.eval(SelTbl + ".Bcch_N5")
           msgdis.Text5(5).Text = MapInfo.eval(SelTbl + ".Bcch_N6")
           msgdis.Text5(6).Text = MapInfo.eval(SelTbl + ".BCCH_SERV")

        End If

           msg = MapInfo.eval(SelTbl + ".MESSAGE")
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

           
        j = j + 1
           
        MapInfo.Do "Fetch next  from " & SelTbl

        If Replay_flag = 2 Then
           ff = SelTbl + ".message"
           msg = LCase(MapInfo.eval(ff))
           If LCase(msg) = LCase(rmsg2) Then
              rtimes = rtime + 1
           End If

           If rtimes = Replay_Time Then
              Unload msgdis
              Unload Me
              Screen.MousePointer = 0
              SUB_532.Enabled = 0
           End If
         End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    MapInfo.Do "set map redraw off"
    MapInfo.Do "delete  from cosmetic1 "
    MapInfo.Do "Set Map Layer 0 Editable Off  "
    MapInfo.Do "set map redraw on"
    flag0 = 0
    Screen.MousePointer = 0
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
    j = i
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
    CurrentX = 0
    CurrentY = 0
    xpos = 600
    old_x(1) = 600
    old_x(2) = 600
    old_x(3) = 600
    old_x(4) = 600
    old_x(5) = 600
    old_x(6) = 600
    Old_y(1) = 1920
    Old_y(2) = 1920
    Old_y(3) = 1920
    Old_y(4) = 1920
    Old_y(5) = 1920
    Old_y(6) = 1920
    
    j = start_rec
    On Error Resume Next
    MapInfo.Do "Fetch Rec " & j & " FROM " & SelTbl
    Graph.Cls
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If flag0 = 1 Then
        Play_graph
    End If
End Sub
