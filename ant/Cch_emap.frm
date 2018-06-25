VERSION 5.00
Begin VB.Form Cch_emap_choice 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CCH 地图显示选择"
   ClientHeight    =   2340
   ClientLeft      =   3135
   ClientTop       =   2100
   ClientWidth     =   3450
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Cch_emap.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2340
   ScaleWidth      =   3450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "下列组合之一"
      Height          =   240
      Left            =   1620
      TabIndex        =   6
      Top             =   225
      Width           =   1380
   End
   Begin VB.CheckBox Check1 
      Caption         =   "取线率"
      Height          =   240
      Left            =   300
      TabIndex        =   5
      Top             =   225
      Value           =   1  'Checked
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "组合选择"
      Height          =   1140
      Left            =   255
      TabIndex        =   2
      Top             =   615
      Width           =   2955
      Begin VB.OptionButton Option2 
         Caption         =   "申请数,分配数"
         Enabled         =   0   'False
         Height          =   240
         Left            =   270
         TabIndex        =   4
         Top             =   780
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "拥塞率,掉话率,信令接通率"
         Enabled         =   0   'False
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   390
         Value           =   -1  'True
         Width           =   2520
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   320
      Left            =   555
      TabIndex        =   1
      Top             =   1965
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   320
      Left            =   1770
      TabIndex        =   0
      Top             =   1965
      Width           =   1080
   End
End
Attribute VB_Name = "Cch_emap_choice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check2_Click()
    On Error Resume Next
    If Check2.Value = 1 Then
       Option1.Enabled = True
       Option2.Enabled = True
    Else
       Option1.Enabled = False
       Option2.Enabled = False
    End If

End Sub

Private Sub Command1_Click()
    Dim max_val, WinId
    Dim center_point, center_lon, center_lat
    Dim i As Integer
    
    On Error Resume Next
    If Check1.Value = 0 And Check2.Value = 0 Then
       Unload Me
       Exit Sub
    End If
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\cch_sts.tab" + Chr(34)
    mapinfo.do "set next document parent " & MapForm.hwnd & "style 1"
    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
       msg = "Add Map Auto Layer" + Chr(34) + "cch_sts" + Chr(34)
       mapinfo.do msg
       msg = Chr(34) + "km" + Chr(34)
'       mapinfo.do "set map zoom 6 units " & msg
       mapinfo.do "set map zoom 10 units " & msg
    Else
       msg = "Map from " + Chr(34) + "cch_sts" + Chr(34)
       mapinfo.do msg
       thereIsAMap = True
       mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    MapForm.Caption = MapForm.Caption + ",CCH"
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
           WinId = mapinfo.eval("windowid(" & i & ")")
           If WinId = mapinfo.eval("frontwindow()") Then
              Exit For
           End If
        End If
    Next
    If Check1.Value = 1 Then
       mapinfo.do "select max(col3) from cch_sts into mytemp"
       max_val = mapinfo.eval("mytemp.col1")
       max_val = Int(max_val)
       mapinfo.do "close table mytemp"
       mapinfo.do "Set Style Pen MakePen(1,60,0)"
       mapinfo.do "set style brush  makebrush(2,7585792,7585792) "
'       mapinfo.do "shade window Frontwindow() cch_sts with col6 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value 200 vary size by " + Chr(34) + "SQRT" + Chr(34) + " border Pen (1,2,8245248)  position center center style Brush (2,8245248,16777215)  # max 200 color 0 #"
'       mapinfo.do "shade window Frontwindow() cch_sts with col6 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value 200 vary size by " + Chr(34) + "LOG" + Chr(34) + " border Pen (1,2,8245248)  position center center style Brush (2,16711935,16777215)  # max 200 color 0 #"
'       mapinfo.do "shade window Frontwindow() cch_sts with col3 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value " & max_val & " vary size by " + Chr(34) + "LOG" + Chr(34) + " border Pen (1,1,0)  position center center style Brush (2,16711935,16777215)  # max " & max_val & " color 0 #"
'       mapinfo.do "set legend window Frontwindow() layer prev display on shades on symbols off lines off count off title " + Chr(34) + " CCH饼状图(对数分级)" + Chr(34) + " Font (" + Chr(34) + "System" + Chr(34) + ",0,12,0) subtitle " + Chr(34) + Chr(34) + " Font (" + Chr(34) + "System" + Chr(34) + ",0,12,0) ascending on ranges Font (" + Chr(34) + "System" + Chr(34) + ",0,12,0) " + Chr(34) + Chr(34) + " display off ," + Chr(34) + "取线率 %" + Chr(34) + " display on"
       mapinfo.do "shade window " + WinId + " cch_sts with col3 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value " & max_val & " vary size by " + Chr(34) + "LOG" + Chr(34) + " border Pen (1,1,0)  position center center style Brush (2,16711935,16777215)  # max " & max_val & " color 0 #"
       mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " CCH饼状图(对数分级)" + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) """" display off ," + Chr(34) + "取线率 %" + Chr(34) + " display on"
    End If
    If Check2.Value = 1 Then
       If Option1.Value = True Then
'          mapinfo.do "shade window Frontwindow() cch_sts with col7,col8,col6 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.635 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215) ,Brush (2,255,16777215)  # max 100 color 0 #"
'          mapinfo.do "set legend window Frontwindow() layer prev display on shades on symbols off lines off count off title " + Chr(34) + " CCH直方图" + Chr(34) + " Font (" + Chr(34) + "System" + Chr(34) + ",0,12,0) subtitle " + Chr(34) + Chr(34) + " Font (" + Chr(34) + "System" + Chr(34) + ",0,12,0) ascending on ranges Font (" + Chr(34) + "System" + Chr(34) + ",0,12,0) " + Chr(34) + Chr(34) + " display off ," + Chr(34) + "拥塞率 %" + Chr(34) + " display on ," + Chr(34) + "掉话率 %" + Chr(34) + " display on ," + Chr(34) + "信令接通率 %" + Chr(34) + " display on"
          mapinfo.do "shade window " + WinId + " cch_sts with col7,col8,col6 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.505 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215) ,Brush (2,255,16777215)  # max 100 color 0 #"
          mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " CCH直方图" + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) """" display off ," + Chr(34) + "拥塞率 %" + Chr(34) + " display on ," + Chr(34) + "掉话率 %" + Chr(34) + " display on ," + Chr(34) + "信令接通率 %" + Chr(34) + " display on"
       End If
       If Option2.Value = True Then
'          mapinfo.do "shade window Frontwindow() cch_sts with col4,col5 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.388 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215)  # max 100 color 0 #"
'          mapinfo.do "set legend window Frontwindow() layer prev display on shades on symbols off lines off count off title " + Chr(34) + " CCH直方图" + Chr(34) + " Font (" + Chr(34) + "System" + Chr(34) + ",0,12,0) subtitle " + Chr(34) + Chr(34) + " Font (" + Chr(34) + "System" + Chr(34) + ",0,12,0) ascending on ranges Font (" + Chr(34) + "System" + Chr(34) + ",0,12,0) " + Chr(34) + Chr(34) + " display off ," + Chr(34) + "申请数" + Chr(34) + " display on ," + Chr(34) + "分配数" + Chr(34) + " display on "
          mapinfo.do "shade window " + WinId + " cch_sts with col4,col5 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.340 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215)  # max 100 color 0 #"
          mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " CCH直方图" + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) """" display off ," + Chr(34) + "申请数" + Chr(34) + " display on ," + Chr(34) + "分配数" + Chr(34) + " display on "
       End If
    End If
    thereIsAMap = True
    If mapid = 0 Then
       mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    MDIMain.SUB_23.Enabled = 1
    MDIMain.SUB_24.Enabled = 1
    MDIMain.SUB_25.Enabled = 1
    MDIMain.SUB_26.Enabled = 1
    MDIMain.SUB_28.Enabled = 1
    
    center_point = mapinfo.eval("tableinfo(cch_sts,8)")
    mapinfo.do "fetch first from cch_sts"
    For i = 1 To center_point
        center_lon = mapinfo.eval("cch_sts.lon")
        center_lat = mapinfo.eval("cch_sts.lat")
        If center_lon <> 0 And center_lat <> 0 Then
           Exit For
        Else
           mapinfo.do "fetch next from cch_sts"
        End If
    Next
    mapinfo.do "set map Center(" & center_lon & "," & center_lat & ") "
    mapinfo.runmenucommand 610
    Unload Me

End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

