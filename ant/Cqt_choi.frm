VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form CQT_choice 
   BackColor       =   &H00C0C0C0&
   Caption         =   "呼叫质量测试"
   ClientHeight    =   3345
   ClientLeft      =   2895
   ClientTop       =   2250
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Cqt_choi.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3345
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "提示"
      Height          =   2280
      Left            =   2400
      TabIndex        =   4
      Top             =   135
      Width           =   2775
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "每小区做 20 次呼叫测试"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   8
         Top             =   435
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "每次通话时长 > 2 分钟"
         Height          =   180
         Index           =   1
         Left            =   195
         TabIndex        =   7
         Top             =   885
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "移动用户呼叫固定用户占 60 %"
         Height          =   180
         Index           =   2
         Left            =   195
         TabIndex        =   6
         Top             =   1320
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "固定用户呼叫移动有户占 40 %"
         Height          =   180
         Index           =   3
         Left            =   195
         TabIndex        =   5
         Top             =   1740
         Width           =   2430
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "小区选择"
      Height          =   2265
      Left            =   150
      TabIndex        =   3
      Top             =   150
      Width           =   2160
      Begin VB.OptionButton Option3 
         Caption         =   "高掉话率"
         Height          =   240
         Left            =   225
         TabIndex        =   14
         Top             =   1230
         Width           =   1035
      End
      Begin VB.OptionButton Option2 
         Caption         =   "高拥塞率"
         Height          =   240
         Left            =   225
         TabIndex        =   13
         Top             =   825
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "高每线话务量"
         Height          =   240
         Left            =   225
         TabIndex        =   12
         Top             =   405
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1260
         TabIndex        =   10
         Text            =   "5"
         Top             =   1725
         Width           =   435
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   1695
         TabIndex        =   9
         Top             =   1725
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text1"
         BuddyDispid     =   196615
         OrigLeft        =   2235
         OrigTop         =   2370
         OrigRight       =   2475
         OrigBottom      =   2625
         Increment       =   5
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "选择小区数:"
         Height          =   180
         Left            =   225
         TabIndex        =   11
         Top             =   1755
         Width           =   990
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   320
      Left            =   2790
      TabIndex        =   1
      Top             =   2910
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   320
      Left            =   1530
      TabIndex        =   0
      Top             =   2910
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "选择结果将保存在 "
      Height          =   180
      Left            =   450
      TabIndex        =   2
      Top             =   2550
      Width           =   1530
   End
End
Attribute VB_Name = "CQT_choice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tab_name As String

Private Sub Command1_Click()
    Dim mytablenum
    Dim i As Integer
    Dim mymsg As String, use_name As String
    Dim Is_motorola As Boolean, Is_open As Boolean
    Dim row_num, lon, lat
    Dim center_point, center_lon, center_lat
    Dim WinId
    
    On Error Resume Next
    Is_motorola = False
    Is_open = False
    If Val(Text1.Text) < 1 Then
       MsgBox "选择的小区数无效！", 64, "提示"
       Unload Me
       Exit Sub
    End If
    Gsm_FileName = Gsm_Path + "\sts\tch_sts.tab"
    If UCase(dir(Gsm_FileName, 0)) <> "TCH_STS.TAB" Then
       MsgBox " TCH_STS.TAB 不存在！", 64, "提示"
       Unload Me
       Exit Sub
    End If
    mytablenum = mapinfo.eval("NumTables()")
    For i = 1 To mytablenum
        If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "TCH_STS" Then
           Is_open = True
           GoTo opened
        End If
    Next
    mapinfo.do "open table " + Chr(34) + Gsm_FileName + Chr(34)
opened:
    If mapinfo.eval("tableinfo(tch_sts,4)") = 18 Then
       Is_motorola = False
    Else
       Is_motorola = True
    End If
    If Option1.Value = True Then
       If Is_motorola = False Then
          mymsg = "col6"
       Else
          mymsg = "col7"
       End If
    End If
    If Option2.Value = True Then
       If Is_motorola = False Then
          mymsg = "col7"
       Else
          mymsg = "col16"
       End If
    End If
    If Option3.Value = True Then
       If Is_motorola = False Then
          mymsg = "col9"
       Else
          mymsg = "col19"
       End If
    End If
On Error GoTo 0
    mapinfo.do "Select * from TCH_STS order by " & mymsg & " into CQT_Temp"
    mytablenum = mapinfo.eval("NumTables()")
    use_name = Left(tab_name, Len(tab_name) - 4)
    For i = 1 To mytablenum
        If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = UCase(use_name) Then
           mapinfo.do "close table " & use_name
           Exit For
        End If
    Next
    
    Gsm_FileName = Gsm_Path + "\cqt.dbf"
    Gsm_File2 = Gsm_Path + "\sts\" + use_name + ".dbf"
    FileCopy Gsm_FileName, Gsm_File2
    mapinfo.do "Register Table  " + Chr(34) + Gsm_File2 + Chr(34) + "Type " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + Gsm_Path + "\sts\" + tab_name + Chr(34)
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\" + tab_name + Chr(34)
    mapinfo.do "Create Map For " & use_name & " CoordSys Earth Projection 1, 0 "
    If Is_motorola = True Then
       mapinfo.do "insert into " + use_name + "(col1,col2,col3,col4,col5,col14,col15,obj) select col1,col2,col7,col16,col19,col24,col25,obj from CQT_Temp"
    Else
       mapinfo.do "insert into " + use_name + "(col1,col2,col3,col4,col5,col14,col15,obj) select col1,col2,col6,col7,col9,col16,col17,obj from CQT_Temp"
    End If
    mapinfo.do "commit table " & use_name
'    mapinfo.do "commit table CQT_Temp as " + Chr(34) + Gsm_Path + "\sts\" + tab_name + Chr(34)
    mapinfo.do "close table CQT_Temp"
    If Is_open = False Then
       mapinfo.do "close table tch_sts"
    End If
'    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\" + tab_name + Chr(34)
    mapinfo.do "fetch first from " & use_name
    row_num = mapinfo.eval("tableinfo(" + use_name + ",8)")
    If row_num > Val(Text1.Text) Then
       For i = 1 To row_num - Val(Text1.Text)
           mapinfo.do "delete from " + Chr(34) + use_name + Chr(34) + "where rowid=" & i
       Next
    End If
    mapinfo.do "commit table " & use_name
    mapinfo.do "pack table " + use_name + " Graphic Data Data Interactive"
    mapinfo.do "fetch first from " & use_name
    If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
       MapForm.WindowState = 0
    End If
    MapForm.Move 0, 10, 12000, 4000
    If InStr(MapForm.Caption, "CQT_Select") = 0 Then
       MapForm.Caption = MapForm.Caption + ",CQT_Select"
    End If
    mapinfo.do "set next document parent " & MapForm.hwnd & "style 1"
    mytablenum = mapinfo.eval("NumTables()")
    If mytablenum > 1 Then
       msg = "Add Map Auto Layer" + Chr(34) + use_name + Chr(34)
       mapinfo.do msg
       msg = Chr(34) + "km" + Chr(34)
       mapinfo.do "set map zoom 6 units " & msg
    Else
       msg = "Map from " + Chr(34) + use_name + Chr(34)
       mapinfo.do msg
       thereIsAMap = True
       mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
           WinId = mapinfo.eval("windowid(" & i & ")")
           If WinId = mapinfo.eval("frontwindow()") Then
              Exit For
           End If
        End If
    Next
    mapinfo.do "shade window " + WinId + use_name + " with col3 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value 1 vary size by " + Chr(34) + "SQRT" + Chr(34) + " border Pen (1,2,8245248)  position center center style Brush (2,8245248,16777215)  # max 1 color 0 #"
    mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " CQT数据选择饼状图" + Chr(34) + " Font(""宋体"",0,9,0) ascending on ranges Font(""宋体"",0,9,0) """" display off ," + Chr(34) + "每线话务量(erl)" + Chr(34) + " display on"
    thereIsAMap = True
    If mapid = 0 Then
       mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
   
'   msg = "set map Center(x1,y1) Smart redraw zoom 4.5 units " + Chr(34) + "km" + Chr(34)
    center_point = mapinfo.eval("tableinfo(" + use_name + ",8)")
    mapinfo.do "fetch first from " & use_name
    For i = 1 To center_point
        center_lon = mapinfo.eval(use_name & ".lon")
        center_lat = mapinfo.eval(use_name & ".lat")
        If center_lon <> 0 And center_lat Then
           Exit For
        Else
           mapinfo.do "fetch next from " & use_name
        End If
    Next
    mapinfo.do "set map Center(" & center_lon & "," & center_lat & ") "
    mapinfo.runmenucommand 610
    MDIMain.SUB_23.Enabled = 1
    MDIMain.SUB_24.Enabled = 1
    MDIMain.SUB_25.Enabled = 1
    MDIMain.SUB_26.Enabled = 1
    MDIMain.SUB_28.Enabled = 1
    mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
    mapinfo.do "set paper units ""pt"""
    mapinfo.do "browse * from " & use_name
    mapinfo.do "set window Frontwindow() Position(0,250) Width 600 Height 160 "
    
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    tab_name = "CQT_" + Format(DATE, "mm") + Format(DATE, "dd") + ".TAB"
    Label3.Caption = Label3.Caption + tab_name
End Sub
