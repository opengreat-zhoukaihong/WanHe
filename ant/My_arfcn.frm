VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form My_ArfcnChanging 
   BackColor       =   &H00C0C0C0&
   Caption         =   "可服务信道变化图"
   ClientHeight    =   1725
   ClientLeft      =   3330
   ClientTop       =   2325
   ClientWidth     =   2985
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "My_arfcn.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1725
   ScaleWidth      =   2985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "条件定义"
      Height          =   1005
      Left            =   255
      TabIndex        =   3
      Top             =   90
      Width           =   2475
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   0
         Text            =   "93"
         Top             =   465
         Width           =   450
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   1530
         TabIndex        =   4
         Top             =   465
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Command2"
         BuddyDispid     =   196613
         OrigLeft        =   1950
         OrigTop         =   555
         OrigRight       =   2190
         OrigBottom      =   795
         Max             =   150
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxLev >="
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   6
         Top             =   495
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-dBm"
         Height          =   180
         Index           =   2
         Left            =   1845
         TabIndex        =   5
         Top             =   510
         Width           =   360
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   300
      Left            =   375
      TabIndex        =   1
      Top             =   1320
      Width           =   1080
   End
End
Attribute VB_Name = "My_ArfcnChanging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Long, j As Integer, col_num As Integer
    Dim row As Long
    Dim check_num As Integer
    Dim lon, lat, MyTableNum
    
    On Error Resume Next
    
    If Trim(Text1.Text) = "" Then
       Unload Me
       Exit Sub
    End If
    Screen.MousePointer = 11
    row = Val(mapinfo.eval("tableinfo(" & tblname & ",8)"))
    col_num = mapinfo.eval("TableInfo(""" & tblname & """, 4)")
    Gsm_FileName = Gsm_Path + "\gsm_temp.tab"
    MyTableNum = mapinfo.eval("NumTables()")
    For i = 1 To MyTableNum
        If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "GSM_TEMP" Then
           mapinfo.do "close table gsm_temp"
        End If
    Next
    mapinfo.do "Create Table " + Chr(34) + "gsm_temp" + Chr(34) + " (my_col Integer,lon Decimal(12,6),lat Decimal(12,6)) file " + Chr(34) + Gsm_FileName + Chr(34) + " TYPE NATIVE Charset " + Chr(34) + "WindowsSimpChinese" + Chr(34)
    mapinfo.do "open table " + Chr(34) + Gsm_FileName + Chr(34)
    mapinfo.do "Create Map For gsm_temp CoordSys Earth Projection 1, 0 "
    mapinfo.do "fetch first from " & tblname
    mapinfo.do "fetch first from gsm_temp"
    For i = 1 To row
        check_num = 0
        For j = 4 To col_num Step 2
            If Val(mapinfo.eval(tblname & ".col" & j)) <= Val(Text1.Text) And Val(mapinfo.eval(tblname & ".col" & (j + 1))) <> 99 Then
               check_num = check_num + 1
            End If
        Next
        mapinfo.do "x1=" & tblname & ".lon"
        mapinfo.do "y1=" & tblname & ".lat"
        'lat = mapinfo.eval(tblname & ".lat")
        mapinfo.do "create point into variable sts_mypoint (x1,y1) symbol(33,0,2)" '(34,7585792,2)"
        mapinfo.do "insert into gsm_temp (col1,col2,col3,obj) values (" & check_num & ",x1,y1,sts_mypoint )"
        mapinfo.do "fetch next from " & tblname
    Next
    mapinfo.do "commit table gsm_temp"
    mapinfo.do "Add Map Auto Layer gsm_temp"
'    msg = " shade window FrontWindow() gsm_temp With col1 values 0 Symbol (41,16711680,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,13671424,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,16776960,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) "
    Msg = " shade window FrontWindow() gsm_temp With col1 values 0 Symbol (41,16711680,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,16776960,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,16756832,8,""MapInfo Cartographic"",0,0) "
    Msg = Msg + ",7 Symbol (41,12632064,8,""MapInfo Cartographic"",0,0),8 Symbol (41,4243711,8,""MapInfo Cartographic"",0,0),9 Symbol (41,13684991,8,""MapInfo Cartographic"",0,0),10 Symbol (41,16765136,8,""MapInfo Cartographic"",0,0),11 Symbol (41,13672703,8,""MapInfo Cartographic"",0,0),12 Symbol (41,16771280,8,""MapInfo Cartographic"",0,0),13 Symbol (41,13693183,8,""MapInfo Cartographic"",0,0),14 Symbol (41,13696976,8,""MapInfo Cartographic"",0,0),15 Symbol (41,16744703,8,""MapInfo Cartographic"",0,0),16 Symbol (41,16777088,8,""MapInfo Cartographic"",0,0)"
    Msg = Msg + ",17 Symbol (41,8454143,8,""MapInfo Cartographic"",0,0),18 Symbol (41,7340256,8,""MapInfo Cartographic"",0,0),19 Symbol (41,16764992,8,""MapInfo Cartographic"",0,0),20 Symbol (41,13696976,8,""MapInfo Cartographic"",0,0) default Symbol (41,0,8,""MapInfo Cartographic"",0,0) "
    mapinfo.do Msg
    If legendid = 0 Then
       mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
       mapinfo.do "Create Legend From Window  Frontwindow()"
       legendid = mapinfo.eval("windowinfo(1009,12)")
    End If
    Msg = " Title " + Chr(34) + "可服务信道变化图 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev>=" + Text1.Text + "(dBm)" + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
    mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub
