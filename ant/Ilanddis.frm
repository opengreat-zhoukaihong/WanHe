VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Iland_Dis 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "µºÐ§Ó¦·ÖÎö"
   ClientHeight    =   3420
   ClientLeft      =   30
   ClientTop       =   5190
   ClientWidth     =   11880
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3420
   ScaleMode       =   0  'User
   ScaleWidth      =   14720.87
   Begin VB.TextBox CName 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   10380
      TabIndex        =   19
      Text            =   "Unkown"
      Top             =   60
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   9600
      Top             =   2400
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2655
      Left            =   10440
      TabIndex        =   13
      Top             =   480
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   4683
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Close 
         Caption         =   "C&lose"
         Default         =   -1  'True
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
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Step 
         Caption         =   "St&ep"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Stop 
         Caption         =   "S&top"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Pause 
         Caption         =   "&Pause"
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
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   975
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
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   10120
      TabIndex        =   25
      Top             =   300
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   20
      Left            =   10120
      TabIndex        =   24
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   19
      Left            =   10120
      TabIndex        =   23
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   18
      Left            =   10120
      TabIndex        =   22
      Top             =   1380
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   17
      Left            =   10120
      TabIndex        =   21
      Top             =   1740
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   195
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      Index           =   4
      X1              =   360.587
      X2              =   10080.33
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      Index           =   3
      X1              =   360.587
      X2              =   10080.33
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   15
      Left            =   10120
      TabIndex        =   12
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   14
      Left            =   10120
      TabIndex        =   11
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   10120
      TabIndex        =   10
      Top             =   2340
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      Index           =   8
      X1              =   360.587
      X2              =   10080.33
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      Index           =   7
      X1              =   360.587
      X2              =   10080.33
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   165
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   165
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Index           =   5
      X1              =   360.587
      X2              =   10080.33
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      Index           =   2
      X1              =   360.587
      X2              =   10080.33
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   360.587
      X2              =   10080.33
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   10120
      TabIndex        =   4
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   2
      Top             =   3045
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Timing Advance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   3
      Left            =   8100
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "RxQual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   360.587
      X2              =   10080.33
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      Height          =   3255
      Left            =   360
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Iland_Dis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag0, go_step, i, j, row, xpos, ypos, nn As Integer
Dim SelTbl, strx, stry, old_ci, CellNo, ci(3) As String
Dim x, y As Single
Dim old_x(6), Old_y(6), start_rec, static_Value(9) As Integer


Private Sub Close_Click()
    On Error Resume Next
    Screen.MousePointer = 0
    Unload msgdis
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    MapInfo.do "set map redraw off"
    MapInfo.do "Set Map Layer 0 Editable On  "
    MapInfo.do "set map redraw on"
    
    old_ci = rmsg1
    CellNo = rmsg2
    CName.Text = CellNo

    Timer1.Enabled = True
    SUB_533.Enabled = 0
    flag0 = 0
    go_step = 0
    nn = 0
    SelTbl = MapInfo.eval("selectionInfo(1)")
    strx = SelTbl + ".lon"
    stry = SelTbl + ".lat"

    For i = 0 To 8
      static_Value(i) = 0
    Next i

    i = Val(MapInfo.eval("selectionInfo(3)"))  ' SEL_INFO_NROWS
    
    CurrentX = 0
    CurrentY = 0
    xpos = 360
    old_x(1) = 360
    Old_y(1) = 3240
    old_x(2) = 360
    Old_y(2) = 3240
    old_x(3) = 360
    Old_y(3) = 3240
    old_x(4) = 360
    Old_y(4) = 3240
    old_x(5) = 360
    Old_y(5) = 3240
    old_x(6) = 360
    Old_y(6) = 3240

    If i <> 0 Then
       row = Val(MapInfo.eval("tableinfo(" & SelTbl & ",8)"))
       i = MapInfo.eval("searchpoint(" & mapid & ",selection.lon,selection.lat)")
       i = Val(MapInfo.eval("SearchInfo(1, 2)"))
       xpos = 360
       MapInfo.do "Fetch Rec " & i & " FROM " & SelTbl
    End If
    j = i
    start_rec = i
    DrawWidth = 2
End Sub

Private Sub Play_graph()
On Error Resume Next
   If j >= row Or xpos >= 10080 Then
      Screen.MousePointer = 0
      Timer1.Enabled = False
      flag0 = 0
      Exit Sub
   End If

    If flag0 = 1 Or go_step = 1 Then
           If xpos >= 10080 Then
                
'              Graphjam.Cls
'               xpos = 360
'               old_x(1) = 360
'               old_x(2) = 360
'               old_x(3) = 360
'               old_x(4) = 360
'               old_x(5) = 360
'               old_x(6) = 360
           End If

           go_step = 0
           
           strx = SelTbl + ".lon"
           stry = SelTbl + ".lat"
           msg = MapInfo.eval(SelTbl + ".MESSAGE")
           If msg = "HANDOVER COMPLETE" Then
              MapInfo.do "Set Style Symbol MakeSymbol(37,255,20)"
           Else
              If msg = "SETUP" Then
                 MapInfo.do "Set Style Symbol MakeSymbol(47,65535,20)"
              Else
                 If msg = "RELEASE COMPLETE" Then
                    Line (xpos, 0)-(xpos, 3240), RGB(25, 150, 10), BF
                    MapInfo.do "Set Style Symbol MakeSymbol(48,65280,20)"
                 Else
                     MapInfo.do "Set Style Symbol MakeSymbol(33,255,4)"
                 End If
              End If
           End If
           msg = "Create Point(" & strx & "," & stry & ")"
           MapInfo.do msg
           
           ci_str = MapInfo.eval("" & SelTbl & ".CI_SERV")
           CName.Text = CellNo
           
           If (ci_str = rmsg1) Then
              xpos = xpos + 5
              nn = 0

              msg = SelTbl + ".Ta"
              ypos = Val(MapInfo.eval(msg))                        'TA
              If ypos <> 0 Then
                 ypos = 3240 - ypos * 180 + 25
                 PSet (xpos, ypos), RGB(255, 255, 0)
              End If
              old_x(1) = xpos
              Old_y(1) = ypos

              msg = SelTbl + ".Rxqual_f"                             ' Rxqual_f
              i = Val(MapInfo.eval(msg))
              ypos = 3240 - i * 360 + 5
              PSet (xpos, ypos), RGB(255, 0, 0)
              old_x(5) = xpos
              Old_y(5) = ypos
          Else
             nn = nn + 1
             If nn <= 20 Then
                xpos = xpos + 5
                Line (xpos, 3225)-(xpos - 5, 3225 - 8 * 360), RGB(100, 150, 220)
             End If
          End If

'           ccc = SelTbl + ".lon"
'           ddd = SelTbl + ".lat"
'           msg = "x1=distance(cell.lon,cell.lat," & ccc & ", " & ddd & ",""m"")"
'           mapinfo.do msg
'           ypos = Val(mapinfo.eval("x1")) / 500               'TRUE DIST
'           If ypos < 10 Then
'              ypos = 3240 - ypos * 120 + 15
'           Else
'              If ypos > 63 Then
'                 ypos = 3270
'              Else
'                 ypos = 3240 - 10 * 120 - (ypos - 10) * 12 + 15
'              End If
'           End If

           j = j + 1
           MapInfo.do "Fetch next  from " & SelTbl
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    MapInfo.do "set map redraw off"
    MapInfo.do "delete  from cosmetic1 "
    MapInfo.do "Set Map Layer 0 Editable Off  "
    MapInfo.do "set map redraw on"
    flag0 = 0
End Sub

Private Sub Label5_Click(Index As Integer)

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
    MapInfo.do "Fetch Rec " & j & " FROM " & SelTbl
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If flag0 = 1 Then
        Play_graph
    End If
End Sub

