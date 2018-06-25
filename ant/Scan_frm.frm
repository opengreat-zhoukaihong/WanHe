VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form Scan_Frm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "扫频回放"
   ClientHeight    =   7095
   ClientLeft      =   2760
   ClientTop       =   1320
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Scan_frm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   5895
   Begin VB.Frame Frame3 
      Caption         =   "图例说明"
      Height          =   1440
      Left            =   3675
      TabIndex        =   51
      Top             =   5355
      Width           =   1995
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   0
         Left            =   375
         Top             =   720
         Width           =   165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "非服务信道"
         Height          =   180
         Index           =   0
         Left            =   705
         TabIndex        =   54
         Top             =   705
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FF00&
         BorderColor     =   &H0000FF00&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   1
         Left            =   375
         Top             =   390
         Width           =   165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "可服务信道"
         Height          =   180
         Index           =   1
         Left            =   705
         TabIndex        =   53
         Top             =   360
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   165
         Index           =   2
         Left            =   375
         Top             =   1065
         Width           =   165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "异网信道"
         Height          =   180
         Index           =   2
         Left            =   705
         TabIndex        =   52
         Top             =   1065
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "自动暂停条件"
      Height          =   1785
      Left            =   3675
      TabIndex        =   38
      Top             =   3450
      Width           =   1995
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   0
         Left            =   1110
         TabIndex        =   45
         Text            =   "0"
         Top             =   330
         Width           =   450
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   1
         Left            =   1110
         TabIndex        =   44
         Text            =   "99"
         Top             =   675
         Width           =   450
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   1110
         TabIndex        =   43
         Text            =   "0"
         Top             =   1035
         Width           =   465
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   1110
         TabIndex        =   42
         Text            =   "-93"
         Top             =   1380
         Width           =   450
      End
      Begin VB.CheckBox Check2 
         Caption         =   "BCCH ="
         Height          =   240
         Left            =   195
         TabIndex        =   40
         Top             =   1050
         Width           =   840
      End
      Begin VB.CheckBox Check1 
         Caption         =   "BCCH ="
         Height          =   240
         Left            =   195
         TabIndex        =   39
         Top             =   345
         Width           =   840
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   41
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text1(0)"
         BuddyDispid     =   196613
         BuddyIndex      =   0
         OrigLeft        =   1650
         OrigTop         =   330
         OrigRight       =   1890
         OrigBottom      =   585
         Max             =   124
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   46
         Top             =   675
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text1(1)"
         BuddyDispid     =   196613
         BuddyIndex      =   1
         OrigLeft        =   1845
         OrigTop         =   780
         OrigRight       =   2085
         OrigBottom      =   1035
         Max             =   99
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   2
         Left            =   1575
         TabIndex        =   47
         Top             =   1035
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text1(2)"
         BuddyDispid     =   196613
         BuddyIndex      =   2
         OrigLeft        =   1830
         OrigTop         =   1170
         OrigRight       =   2070
         OrigBottom      =   1425
         Max             =   124
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   48
         Top             =   1365
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text1(3)"
         BuddyDispid     =   196613
         BuddyIndex      =   3
         OrigLeft        =   1830
         OrigTop         =   1515
         OrigRight       =   2070
         OrigBottom      =   1770
         Max             =   0
         Min             =   -126
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "BSIC ="
         Height          =   180
         Index           =   0
         Left            =   465
         TabIndex        =   50
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxLev <"
         Height          =   180
         Index           =   1
         Left            =   330
         TabIndex        =   49
         Top             =   1425
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "控制板"
      Height          =   3015
      Left            =   3675
      TabIndex        =   30
      Top             =   315
      Width           =   1995
      Begin ComctlLib.Slider Slider1 
         Height          =   300
         Left            =   105
         TabIndex        =   37
         Top             =   2520
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   529
         _Version        =   327680
         MouseIcon       =   "Scan_frm.frx":030A
         LargeChange     =   10
         Max             =   100
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.CommandButton Command5 
         Caption         =   "单步"
         Height          =   320
         Left            =   465
         TabIndex        =   35
         Top             =   1335
         Width           =   1080
      End
      Begin VB.CommandButton Command4 
         Caption         =   "暂停"
         Height          =   320
         Left            =   465
         TabIndex        =   34
         Top             =   705
         Width           =   1080
      End
      Begin VB.CommandButton Command3 
         Caption         =   "中止"
         Height          =   320
         Left            =   465
         TabIndex        =   33
         Top             =   1020
         Width           =   1080
      End
      Begin VB.CommandButton Command2 
         Caption         =   "关闭"
         Height          =   320
         Left            =   465
         TabIndex        =   32
         Top             =   1650
         Width           =   1080
      End
      Begin VB.CommandButton Command1 
         Caption         =   "开始"
         Height          =   320
         Left            =   465
         TabIndex        =   31
         Top             =   390
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "播放速度:"
         Height          =   180
         Left            =   210
         TabIndex        =   36
         Top             =   2220
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   240
      ScaleHeight     =   165
      ScaleWidth      =   4845
      TabIndex        =   22
      Top             =   15
      Width           =   4845
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxLev"
         Height          =   180
         Index           =   5
         Left            =   3405
         TabIndex        =   28
         Top             =   0
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "(dbm)"
         Height          =   180
         Index           =   4
         Left            =   3885
         TabIndex        =   27
         Top             =   15
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-30"
         Height          =   180
         Index           =   3
         Left            =   2880
         TabIndex        =   26
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-50"
         Height          =   180
         Index           =   2
         Left            =   2235
         TabIndex        =   25
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-80"
         Height          =   180
         Index           =   1
         Left            =   1200
         TabIndex        =   24
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-110"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   23
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   510
      ScaleHeight     =   150
      ScaleWidth      =   4245
      TabIndex        =   21
      Top             =   135
      Width           =   4245
      Begin VB.Line Line1 
         Index           =   8
         X1              =   2765
         X2              =   2765
         Y1              =   60
         Y2              =   150
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   2425
         X2              =   2425
         Y1              =   60
         Y2              =   150
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   2085
         X2              =   2085
         Y1              =   60
         Y2              =   150
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   1745
         X2              =   1745
         Y1              =   60
         Y2              =   150
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   1405
         X2              =   1405
         Y1              =   60
         Y2              =   150
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1065
         X2              =   1065
         Y1              =   60
         Y2              =   150
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   725
         X2              =   725
         Y1              =   60
         Y2              =   150
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   385
         X2              =   385
         Y1              =   60
         Y2              =   150
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   45
         X2              =   45
         Y1              =   60
         Y2              =   150
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2850
      Top             =   450
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   6495
      Left            =   555
      ScaleHeight     =   6435
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   285
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "BCCH"
      Height          =   180
      Index           =   6
      Left            =   165
      TabIndex        =   29
      Top             =   6750
      Width           =   360
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   19
      Left            =   180
      TabIndex        =   20
      Top             =   6465
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   18
      Left            =   180
      TabIndex        =   19
      Top             =   6150
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   17
      Left            =   180
      TabIndex        =   18
      Top             =   5835
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   16
      Left            =   180
      TabIndex        =   17
      Top             =   5505
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   15
      Left            =   180
      TabIndex        =   16
      Top             =   5190
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   14
      Left            =   195
      TabIndex        =   15
      Top             =   4875
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   13
      Left            =   180
      TabIndex        =   14
      Top             =   4545
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   12
      Left            =   180
      TabIndex        =   13
      Top             =   4230
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   11
      Left            =   180
      TabIndex        =   12
      Top             =   3915
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   10
      Left            =   180
      TabIndex        =   11
      Top             =   3585
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   9
      Left            =   180
      TabIndex        =   10
      Top             =   3270
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   8
      Left            =   180
      TabIndex        =   9
      Top             =   2955
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   7
      Left            =   180
      TabIndex        =   8
      Top             =   2625
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   6
      Left            =   180
      TabIndex        =   7
      Top             =   2310
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   5
      Left            =   180
      TabIndex        =   6
      Top             =   1995
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   5
      Top             =   1665
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   4
      Top             =   1350
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   3
      Top             =   1035
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   705
      Width           =   540
   End
   Begin VB.Label ARFCN1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   390
      Width           =   540
   End
End
Attribute VB_Name = "Scan_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Play_Flag As Boolean
Dim go_step As Integer
Dim my_seltbl, my_lon, my_lat
Dim all_row, col_num
Dim start_rec As Long
Dim Minbsic As Integer, Maxbsic As Integer
Dim next_row As Long
Dim xpos, ypos
Dim my_msg As String
Dim UnLocal_Frag(0 To 19) As Boolean

Private Sub Command1_Click()
    Play_Flag = True
    Command1.Caption = "继续"
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Play_Flag = False
    Unload Me
    mapinfo.Do "set map redraw off"
    mapinfo.Do "delete  from cosmetic1 "
    mapinfo.Do "Set Map Layer 0 Editable Off  "
    mapinfo.Do "set map redraw on"
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Play_Flag = False
    go_step = 0
    next_row = start_rec
    mapinfo.Do "Fetch Rec " & start_rec & " FROM " & my_seltbl
    Command1.Caption = "开始"
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    Play_Flag = False
    
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    go_step = 1
    Play_graph

End Sub


Private Sub Form_Load()
    Dim my_string As String
    Dim sel_row
    Dim i As Integer
    Dim Local_Min As Integer, Local_Max As Integer
    
    On Error Resume Next
    mapinfo.Do "set map redraw off"
    mapinfo.Do "Set Map Layer 0 Editable On"
    mapinfo.Do "set map redraw on"
    Play_Flag = False
    go_step = 0
    my_seltbl = mapinfo.eval("selectionInfo(1)")
    my_lon = my_seltbl + ".lon"
    my_lat = my_seltbl + ".lat"
    sel_row = mapinfo.eval("selectionInfo(3)")
    If sel_row <> 0 Then
       all_row = mapinfo.eval("tableinfo(" & my_seltbl & ",8)")
       sel_row = mapinfo.eval("searchpoint(" & mapid & ",selection.lon,selection.lat)")
       sel_row = mapinfo.eval("SearchInfo(1, 2)")
       mapinfo.Do "Fetch Rec " & sel_row & " FROM " & my_seltbl
    End If
    next_row = sel_row
    start_rec = sel_row
    my_msg = "TableInfo(""" & my_seltbl & """, 4)"
    col_num = mapinfo.eval(my_msg)
    mapinfo.Do "select min(arfcn) from cell into temp"
    Local_Min = Val(mapinfo.eval("temp.col1"))
    mapinfo.Do "select max(arfcn) from cell into temp"
    Local_Max = Val(mapinfo.eval("temp.col1"))
    For i = 0 To 19
        UnLocal_Frag(i) = False
    Next
    For i = 4 To col_num Step 2
        my_msg = "Columninfo(""" & my_seltbl & """,""COL" & i & """, 1)"
        my_string = mapinfo.eval(my_msg)
        ARFCN1((i / 2) - 2).Caption = Mid(my_string, 7)
        If Val(Mid(my_string, 7)) > Local_Max Or Val(Mid(my_string, 7)) < Local_Min Then
           UnLocal_Frag(i / 2 - 2) = True
        End If
        If i = 4 Then
           Text1(0).Text = ARFCN1((i / 2) - 2).Caption
           Text1(2).Text = ARFCN1((i / 2) - 2).Caption
        End If
        If i / 2 - 2 = 19 Then
           Exit For
        End If
    Next
    For i = (col_num - 3) / 2 To 19
        ARFCN1(i).Caption = ""
    Next
    all_row = Val(mapinfo.eval("tableinfo(" & my_seltbl & ",8)"))
    mapinfo.Do "select Min(BSIC) From cell into Temp"
    Minbsic = Val(mapinfo.eval("Temp.col1"))
    mapinfo.Do "select Max(BSIC) From cell into Temp"
    Maxbsic = Val(mapinfo.eval("Temp.col1"))
    mapinfo.Do "close table temp"
End Sub

Private Sub Play_graph()
    Dim i As Integer, j As Integer
    Dim bsic_val
    Dim check_bsic As Boolean, check_rxlev As Boolean
    Dim bsic_col As Integer, rxlev_col As Integer
    Dim my_string As String
    
    On Error Resume Next
    If Check1.Value = 1 Then
       For i = 4 To col_num Step 2
           my_msg = "Columninfo(""" & my_seltbl & """,""COL" & i & """, 1)"
           my_string = mapinfo.eval(my_msg)
           If Trim(Text1(0).Text) = Trim(Mid(my_string, 7)) Then
              check_bsic = True
              bsic_col = i
              Exit For
           End If
       Next
    End If
    If Check2.Value = 1 Then
       For i = 4 To col_num Step 2
           my_msg = "Columninfo(""" & my_seltbl & """,""COL" & i & """, 1)"
           my_string = mapinfo.eval(my_msg)
           If Trim(Text1(2).Text) = Trim(Mid(my_string, 7)) Then
              check_rxlev = True
              rxlev_col = i
              Exit For
           End If
       Next
    End If
    If Play_Flag = True Or go_step = 1 Then
       go_step = 0
       If next_row > all_row Then
          Play_Flag = False
          Exit Sub
       End If
       For i = 4 To col_num Step 2
           my_msg = my_seltbl + ".col" & i
           ypos = ARFCN1((i / 2) - 2).Top - Picture1.Top - 38   'win95
           xpos = (110 - Val(mapinfo.eval(my_msg))) * 34
           If xpos < 0 Then
              xpos = 0
           End If
           Picture1.Line (0, ypos)-(Picture1.Width, ypos + 210), RGB(0, 0, 0), BF
           j = i + 1
           my_msg = my_seltbl + ".col" & j
           bsic_val = mapinfo.eval(my_msg)
           If bsic_val = 99 Then
              Picture1.Line (0, ypos)-(xpos, ypos + 210), RGB(255, 0, 0), BF
           Else
              If UnLocal_Frag(i / 2 - 2) = True Then
                 Picture1.Line (0, ypos)-(xpos, ypos + 210), RGB(0, 0, 255), BF
              Else
                 Picture1.Line (0, ypos)-(xpos, ypos + 210), RGB(0, 255, 0), BF
              End If
           End If
           Picture1.Line (578, 0)-(578, Picture1.Height), RGB(255, 0, 0), B
           my_msg = bsic_val
           Picture1.CurrentX = xpos - 300    'win95
           Picture1.CurrentY = ypos + 13
           Picture1.Print my_msg
           If check_bsic = True And i = 4 Then
              If mapinfo.eval(my_seltbl & ".col" & (bsic_col + 1)) = Trim(Text1(1).Text) Then
                 Play_Flag = False
              End If
           End If
           If check_rxlev = True And i = 4 Then
'              If Val(mapinfo.eval(my_seltbl & ".col" & rxlev_col)) < Val(Text1(3).Text) + 110 Then
              If Val(mapinfo.eval(my_seltbl & ".col" & rxlev_col)) > Abs(Val(Text1(3).Text)) Then
                 Play_Flag = False
              End If
           End If
           If i / 2 - 2 = 19 Then
              Exit For
           End If
       Next
       mapinfo.Do "Set Style Symbol MakeSymbol(33,255,4)"
       my_msg = "Create Point(" & my_lon & "," & my_lat & ")"
       mapinfo.Do my_msg
       mapinfo.Do "Fetch next from " & my_seltbl
       next_row = next_row + 1
          
    End If
End Sub

Private Sub HScroll1_Change()
    On Error Resume Next
    'If HScroll1.Value < 95 Then
    '   Timer1.Interval = (100 - HScroll1.Value) * 10
    'Else
    '   Timer1.Interval = 50
    'End If
End Sub

Private Sub HScroll1_Scroll()
    On Error Resume Next
    'If HScroll1.Value < 95 Then
    '   Timer1.Interval = (100 - HScroll1.Value) * 10
    'Else
    '   Timer1.Interval = 30
    'End If
End Sub


Private Sub SpinButton1_SpinDown(Index As Integer)
    On Error Resume Next
    If Index < 3 Then
       If Val(Text1(Index).Text) <= 0 Then
          Exit Sub
       End If
    End If
    If Index = 1 Then
       If Val(Text1(1).Text) = 99 Then
          Text1(1).Text = 77
          Exit Sub
       End If
    End If
    Text1(Index).Text = Trim(str(Val(Text1(Index).Text) - 1))
End Sub

Private Sub SpinButton1_SpinUp(Index As Integer)
    On Error Resume Next
    If Index = 0 Or Index = 2 Then
       If Val(Text1(Index).Text) >= 124 Then
          Exit Sub
       End If
    End If
    If Index = 1 Then
       If Val(Text1(1).Text) >= 77 Then
          Text1(1).Text = 99
          Exit Sub
       End If
    End If
    Text1(Index).Text = Trim(str(Val(Text1(Index).Text) + 1))
End Sub

Private Sub Slider1_Change()
    On Error Resume Next
    If Slider1.Value < 95 Then
       Timer1.Interval = (100 - Slider1.Value) * 10
    Else
       Timer1.Interval = 50
    End If
End Sub

Private Sub Slider1_Scroll()
    On Error Resume Next
    If Slider1.Value < 95 Then
       Timer1.Interval = (100 - Slider1.Value) * 10
    Else
       Timer1.Interval = 30
    End If
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If Play_Flag = True And go_step = 0 Then
        Play_graph
    End If
End Sub

Private Sub UpDown1_DownClick(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
       If Val(Text1(1).Text) = 98 Then
          Text1(1).Text = 77
          Exit Sub
       End If
    End If

End Sub

Private Sub UpDown1_UpClick(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
       If Val(Text1(1).Text) > 77 Then
          Text1(1).Text = 99
          Exit Sub
       End If
    End If
End Sub
