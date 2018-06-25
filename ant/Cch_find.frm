VERSION 5.00
Begin VB.Form Cch_data_find 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CCH 数据查询"
   ClientHeight    =   3360
   ClientLeft      =   2220
   ClientTop       =   1755
   ClientWidth     =   4275
   Icon            =   "Cch_find.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3360
   ScaleWidth      =   4275
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   165
      TabIndex        =   7
      Top             =   75
      Width           =   3945
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   2835
         TabIndex        =   3
         Top             =   1875
         Width           =   645
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   2835
         TabIndex        =   2
         Top             =   1485
         Width           =   645
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   2835
         TabIndex        =   1
         Top             =   1095
         Width           =   645
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   2835
         TabIndex        =   0
         Top             =   690
         Width           =   645
      End
      Begin VB.ComboBox Cond 
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
         Index           =   3
         Left            =   1635
         TabIndex        =   16
         Text            =   " "
         Top             =   1875
         Width           =   870
      End
      Begin VB.ComboBox Cond 
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
         Index           =   2
         Left            =   1635
         TabIndex        =   15
         Text            =   " "
         Top             =   1485
         Width           =   870
      End
      Begin VB.ComboBox Cond 
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
         Index           =   1
         Left            =   1635
         TabIndex        =   14
         Text            =   " "
         Top             =   1095
         Width           =   870
      End
      Begin VB.ComboBox Cond 
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
         Index           =   0
         Left            =   1635
         TabIndex        =   13
         Text            =   " "
         Top             =   690
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Caption         =   "信令接通率"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   255
         TabIndex        =   11
         Top             =   1920
         Width           =   1260
      End
      Begin VB.CheckBox Check1 
         Caption         =   "掉话率"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   255
         TabIndex        =   10
         Top             =   1530
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "拥塞率"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   255
         TabIndex        =   9
         Top             =   1110
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "取线率"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   8
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   3570
         TabIndex        =   22
         Top             =   1905
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   3585
         TabIndex        =   21
         Top             =   1515
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   3570
         TabIndex        =   20
         Top             =   1125
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   3570
         TabIndex        =   19
         Top             =   720
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "条件值:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   2835
         TabIndex        =   18
         Top             =   315
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "条件符号:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   1635
         TabIndex        =   17
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "数据项:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   255
         TabIndex        =   12
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "查询结果保存在 CCH_FIND.TAB 中"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   420
      TabIndex        =   4
      Top             =   2595
      Width           =   3105
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   855
      TabIndex        =   5
      Top             =   2985
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   2100
      TabIndex        =   6
      Top             =   2985
      Width           =   1080
   End
End
Attribute VB_Name = "Cch_data_find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim moto As Boolean

Private Sub Command1_Click()
    Dim MyMsg(3) As String
    Dim domsg As String
    Dim i As Integer, cond_no As Integer
    Dim cond_text(3) As String
    Dim mysubtitle As String
    Dim max_val, WinId
    Dim center_point, center_lon, center_lat
    
    On Error Resume Next
    cond_no = 0
    If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 Then
       Unload Me
       Exit Sub
    End If
    cond_text(0) = Trim(str(Val(Text1(0).Text)))
    For i = 1 To 3
        cond_text(i) = Trim(str(Val(Text1(i).Text)))
    Next
    For i = 0 To 3
        If Mid(cond_text(i), 1, 1) = "." Then
           cond_text(i) = "0" + cond_text(i)
        End If
    Next
    
    For i = 0 To 3
        If Check1(i).Value = 1 Then
           If Cond(i).Text <> "" And Text1(i).Text <> "" Then
              cond_no = cond_no + 1
              If i = 0 Then
                 If moto = True Then
                    MyMsg(i) = "col7" + Cond(i).Text + cond_text(i)
                 Else
                    MyMsg(i) = "col3" + Cond(i).Text + cond_text(i)
                 End If
              End If
              If i = 1 Then
                 If moto = True Then
                    MyMsg(i) = "col15" + Cond(i).Text + cond_text(i)
                 Else
                    MyMsg(i) = "col7" + Cond(i).Text + cond_text(i)
                 End If
              End If
              If i = 2 Then
                 If moto = True Then
                    MyMsg(i) = "col18" + Cond(i).Text + cond_text(i)
                 Else
                    MyMsg(i) = "col8" + Cond(i).Text + cond_text(i)
                 End If
              End If
              If i = 3 Then
                 If moto = True Then
                    MyMsg(i) = "col16" + Cond(i).Text + cond_text(i)
                 Else
                    MyMsg(i) = "col6" + Cond(i).Text + cond_text(i)
                 End If
              End If
           Else
              MyMsg(i) = ""
           End If
        Else
           MyMsg(i) = ""
        End If
    Next
    domsg = ""
    If cond_no = 1 Then
       For i = 0 To 3
           If MyMsg(i) <> "" Then
              domsg = MyMsg(i)
              Exit For
           End If
       Next
    Else
       For i = 0 To 3
           If MyMsg(i) <> "" Then
              domsg = domsg + "(" + MyMsg(i) + ")" + " and "
          End If
       Next
       domsg = Left(domsg, Len(domsg) - 5)
    End If
    If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
       MapForm.WindowState = 0
    End If
    MapForm.Move 0, 10, 12000, 4000
    MapForm.Caption = MapForm.Caption + ",CCH_Find"
    TableNum = Val(mapinfo.eval("NumTables()"))
    mapinfo.do "select * from cch_sts where " + domsg + " Into Cch_Find"
    mapinfo.do "set next document parent " & MapForm.hwnd & "style 1"
    
    If TableNum > 1 Then
       msg = "Add Map Auto Layer" + Chr(34) + "Cch_Find" + Chr(34)
       mapinfo.do msg
       msg = Chr(34) + "km" + Chr(34)
       mapinfo.do "set map zoom 6 units " & msg
    Else
       msg = "Map from " + Chr(34) + "Cch_Find" + Chr(34)
       mapinfo.do msg
       thereIsAMap = True
       mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    mysubtitle = ""
    If moto = True Then
       mysubtitle = "每线话务量" + Trim(Cond(0).Text) + cond_text(0) + "erl "
    Else
       mysubtitle = "取线率" + Trim(Cond(0).Text) + cond_text(0) + "% "
    End If
    mysubtitle = mysubtitle + "拥塞率" + Trim(Cond(1).Text) + cond_text(1) + "% "
    mysubtitle = mysubtitle + "掉话率" + Trim(Cond(2).Text) + cond_text(2) + "% "
    If moto = True Then
       mysubtitle = mysubtitle + "呼叫成功率" + Trim(Cond(3).Text) + cond_text(3) + "%"
    Else
       mysubtitle = mysubtitle + "信令接通率" + Trim(Cond(3).Text) + cond_text(3) + "%"
    End If
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
           WinId = mapinfo.eval("windowid(" & i & ")")
           If WinId = mapinfo.eval("frontwindow()") Then
              Exit For
           End If
        End If
    Next
    If moto = True Then
       mapinfo.do "shade window " + WinId + " cch_find with col7 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value 1 vary size by " + Chr(34) + "SQRT" + Chr(34) + " border Pen (1,2,8245248)  position center center style Brush (2,8245248,16777215)  # max 1 color 0 #"
       mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " CCH数据查询饼状图" + Chr(34) + " Font (""宋体"",0,9,0) subtitle " + Chr(34) + mysubtitle + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) """" display off ," + Chr(34) + "每线话务量(erl)" + Chr(34) + " display on"
    Else
       mapinfo.do "select max(col3) from cch_find into mytemp"
       max_val = mapinfo.eval("mytemp.col1")
       max_val = Int(max_val)
       mapinfo.do "close table mytemp"
'       mapinfo.do "shade window Frontwindow() cch_find with col3 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value 200 vary size by " + Chr(34) + "SQRT" + Chr(34) + " border Pen (1,2,8245248)  position center center style Brush (2,8245248,16777215)  # max 200 color 0 #"
       mapinfo.do "shade window " + WinId + " cch_find with col3 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value " & max_val & " vary size by " + Chr(34) + "LOG" + Chr(34) + " border Pen (1,1,0)  position center center style Brush (2,16711935,16777215)  # max " & max_val & " color 0 #"
       mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " CCH数据查询饼状图" + Chr(34) + " Font (""宋体"",0,9,0) subtitle " + Chr(34) + mysubtitle + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) """" display off ," + Chr(34) + "取线率" + Chr(34) + " display on"
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
    
    center_point = mapinfo.eval("tableinfo(cch_find,8)")
    mapinfo.do "fetch first from cch_find"
    For i = 1 To center_point
        center_lon = mapinfo.eval("cch_find.lon")
        center_lat = mapinfo.eval("cch_find.lat")
        If center_lon <> 0 And center_lat Then
           Exit For
        Else
           mapinfo.do "fetch next from cch_find"
        End If
    Next
    mapinfo.do "set map Center(" & center_lon & "," & center_lat & ") "
    mapinfo.runmenucommand 610
    mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
    mapinfo.do "set paper units ""pt"""
    mapinfo.do "browse * from Cch_Find"
    
    mapinfo.do "set window Frontwindow() Position(0,250) Width 600 Height 160 "
    If Check2.Value = 1 Then
       mapinfo.do "commit table Cch_Find as " + Chr(34) + Gsm_Path + "\sts\cch_find.tab" + Chr(34)
    End If
    Unload Me

End Sub

Private Sub Command2_Click()
    On Error Resume Next
    mapinfo.do "close table cch_sts"
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    On Error Resume Next
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\cch_sts.tab" + Chr(34)
    If mapinfo.eval("tableinfo(cch_sts,4)") = 26 Then
       Check1(0).Caption = "每线话务量"
       Label1(3).Caption = "erl"
       Check1(3).Caption = "呼叫成功率"
       moto = True
    End If
    For i = 0 To 3
        Cond(i).AddItem " > "
        Cond(i).AddItem " < "
        Cond(i).AddItem " = "
        Cond(i).AddItem " >= "
        Cond(i).AddItem " <= "
    Next

End Sub
