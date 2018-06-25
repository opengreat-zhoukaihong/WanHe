VERSION 5.00
Begin VB.Form Tch_data_find 
   BackColor       =   &H00C0C0C0&
   Caption         =   "TCH 数据查询"
   ClientHeight    =   3435
   ClientLeft      =   2730
   ClientTop       =   1860
   ClientWidth     =   4245
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tch_find.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3435
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "查询结果保存在 TCH_FIND.TAB 中"
      Height          =   240
      Left            =   420
      TabIndex        =   22
      Top             =   2655
      Width           =   3030
   End
   Begin VB.Frame Frame1 
      Height          =   2430
      Left            =   150
      TabIndex        =   2
      Top             =   45
      Width           =   3930
      Begin VB.ComboBox Cond 
         Height          =   300
         Index           =   0
         Left            =   1680
         TabIndex        =   17
         Text            =   " "
         Top             =   705
         Width           =   795
      End
      Begin VB.ComboBox Cond 
         Height          =   300
         Index           =   1
         Left            =   1680
         TabIndex        =   16
         Text            =   " "
         Top             =   1125
         Width           =   795
      End
      Begin VB.ComboBox Cond 
         Height          =   300
         Index           =   2
         Left            =   1680
         TabIndex        =   15
         Text            =   " "
         Top             =   1545
         Width           =   795
      End
      Begin VB.ComboBox Cond 
         Height          =   300
         Index           =   3
         Left            =   1680
         TabIndex        =   14
         Text            =   " "
         Top             =   1965
         Width           =   795
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   0
         Left            =   2835
         TabIndex        =   13
         Text            =   "0.4"
         Top             =   720
         Width           =   645
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   1
         Left            =   2835
         TabIndex        =   12
         Top             =   1140
         Width           =   660
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   2
         Left            =   2835
         TabIndex        =   11
         Top             =   1560
         Width           =   660
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   3
         Left            =   2835
         TabIndex        =   10
         Top             =   1980
         Width           =   645
      End
      Begin VB.CheckBox Check1 
         Caption         =   "话音接通率"
         Height          =   240
         Index           =   3
         Left            =   255
         TabIndex        =   9
         Top             =   1995
         Width           =   1200
      End
      Begin VB.CheckBox Check1 
         Caption         =   "掉话率"
         Height          =   240
         Index           =   2
         Left            =   255
         TabIndex        =   8
         Top             =   1575
         Width           =   840
      End
      Begin VB.CheckBox Check1 
         Caption         =   "拥塞率"
         Height          =   240
         Index           =   1
         Left            =   255
         TabIndex        =   7
         Top             =   1155
         Width           =   840
      End
      Begin VB.CheckBox Check1 
         Caption         =   "每线话务量"
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   6
         Top             =   735
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "erl"
         Height          =   180
         Index           =   3
         Left            =   3525
         TabIndex        =   21
         Top             =   750
         Width           =   270
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
         Left            =   3555
         TabIndex        =   20
         Top             =   1170
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
         Left            =   3555
         TabIndex        =   19
         Top             =   1605
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
         Index           =   6
         Left            =   3540
         TabIndex        =   18
         Top             =   2025
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "数据项:"
         Height          =   180
         Index           =   0
         Left            =   315
         TabIndex        =   5
         Top             =   315
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "条件符号:"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   4
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "条件值:"
         Height          =   180
         Index           =   2
         Left            =   2835
         TabIndex        =   3
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   320
      Left            =   2100
      TabIndex        =   1
      Top             =   3045
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   320
      Left            =   870
      TabIndex        =   0
      Top             =   3045
      Width           =   1080
   End
End
Attribute VB_Name = "Tch_data_find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moto As Boolean
Private Sub Command1_Click()
    Dim mymsg(3) As String
    Dim domsg As String
    Dim i As Integer, cond_no As Integer
    Dim cond_text(3) As String
    Dim mysubtitle As String
    Dim center_point, center_lon, center_lat
    Dim WinId
    
    On Error Resume Next
    cond_no = 0
    If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 Then
       Unload Me
       Exit Sub
    End If
    cond_text(0) = Trim(str(Val(Text1(0).Text)))
    For i = 1 To 3
'        cond_text(i) = Trim(str(Val(Text1(i).Text) / 100))
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
                    mymsg(i) = "col7" + Cond(i).Text + cond_text(i)
                 Else
                    mymsg(i) = "col6" + Cond(i).Text + cond_text(i)
                 End If
              End If
              If i = 1 Then
                 If moto = True Then
                    mymsg(i) = "col16" + Cond(i).Text + cond_text(i)
                 Else
                    mymsg(i) = "col7" + Cond(i).Text + cond_text(i)
                 End If
              End If
              If i = 2 Then
                 If moto = True Then
                    mymsg(i) = "col19" + Cond(i).Text + cond_text(i)
                 Else
                    mymsg(i) = "col9" + Cond(i).Text + cond_text(i)
                 End If
              End If
              If i = 3 Then
                 If moto = True Then
                    mymsg(i) = "col17" + Cond(i).Text + cond_text(i)
                 Else
                    mymsg(i) = "col5" + Cond(i).Text + cond_text(i)
                 End If
              End If
           Else
              mymsg(i) = ""
           End If
        Else
           mymsg(i) = ""
        End If
    Next
    domsg = ""
    If cond_no = 1 Then
       For i = 0 To 3
           If mymsg(i) <> "" Then
              domsg = mymsg(i)
              Exit For
           End If
       Next
    Else
       For i = 0 To 3
           If mymsg(i) <> "" Then
              domsg = domsg + "(" + mymsg(i) + ")" + " and "
           End If
       Next
       domsg = Left(domsg, Len(domsg) - 5)
    End If
    If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
       MapForm.WindowState = 0
    End If
    MapForm.Move 0, 10, 12000, 4000
    MapForm.Caption = MapForm.Caption + ",TCH_Find"
    TableNum = Val(mapinfo.eval("NumTables()"))
    mapinfo.do "select * from tch_sts where " + domsg + " Into Tch_Find"
    mapinfo.do "set next document parent " & MapForm.hwnd & "style 1"
    If TableNum > 1 Then
       msg = "Add Map Auto Layer" + Chr(34) + "Tch_Find" + Chr(34)
       mapinfo.do msg
       msg = Chr(34) + "km" + Chr(34)
       mapinfo.do "set map zoom 6 units " & msg
    Else
       msg = "Map from " + Chr(34) + "Tch_Find" + Chr(34)
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
    If moto = True Then
       mapinfo.do "shade window " + WinId + " tch_find with col7 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value 1 vary size by " + Chr(34) + "SQRT" + Chr(34) + " border Pen (1,2,8245248)  position center center style Brush (2,8245248,16777215)  # max 1 color 0 #"
    Else
       mapinfo.do "shade window " + WinId + " tch_find with col6 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value 1 vary size by " + Chr(34) + "SQRT" + Chr(34) + " border Pen (1,2,8245248)  position center center style Brush (2,8245248,16777215)  # max 1 color 0 #"
    End If
    mysubtitle = ""
    If Check1(0).Value = 1 Then
       mysubtitle = "每线话务量" + Trim(Cond(0).Text) + cond_text(0) + "erl "
    End If
    If Check1(1).Value = 1 Then
       mysubtitle = mysubtitle + "拥塞率" + Trim(Cond(1).Text) + cond_text(1) + "% "
    End If
    If Check1(2).Value = 1 Then
       mysubtitle = mysubtitle + "掉话率" + Trim(Cond(2).Text) + cond_text(2) + "% "
    End If
    If Check1(3).Value = 1 Then
       If moto = True Then
          mysubtitle = mysubtitle + "呼叫成功率" + Trim(Cond(3).Text) + cond_text(3) + "%"
       Else
          mysubtitle = mysubtitle + "话音接通率" + Trim(Cond(3).Text) + cond_text(3) + "%"
       End If
    End If
    mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " TCH数据查询饼状图" + Chr(34) + " Font(""宋体"",0,9,0) subtitle " + Chr(34) + mysubtitle + Chr(34) + " Font(""宋体"",0,9,0) ascending on ranges Font(""宋体"",0,9,0) """" display off ," + Chr(34) + "每线话务量(erl)" + Chr(34) + " display on"
    
    thereIsAMap = True
    If mapid = 0 Then
       mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    MDIMain.SUB_23.Enabled = 1
    MDIMain.SUB_24.Enabled = 1
    MDIMain.SUB_25.Enabled = 1
    MDIMain.SUB_26.Enabled = 1
    MDIMain.SUB_28.Enabled = 1
    center_point = mapinfo.eval("tableinfo(Tch_Find,8)")
    mapinfo.do "fetch first from Tch_Find"
    For i = 1 To center_point
        center_lon = mapinfo.eval("Tch_Find.lon")
        center_lat = mapinfo.eval("Tch_Find.lat")
        If center_lon <> 0 And center_lat <> 0 Then
           Exit For
        Else
           mapinfo.do "fetch next from Tch_Find"
        End If
    Next
    mapinfo.do "set map Center(" & center_lon & "," & center_lat & ") "
    mapinfo.runmenucommand 610
    
    mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
    mapinfo.do "set paper units ""pt"""
    mapinfo.do "browse * from Tch_Find"
    
    mapinfo.do "set window Frontwindow() Position(0,250) Width 600 Height 160 "
    If Check2.Value = 1 Then
       mapinfo.do "commit table tch_find as " + Chr(34) + Gsm_Path + "\sts\tch_find.tab" + Chr(34)
    End If
    Unload Me
    
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    mapinfo.do "close table tch_sts"
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    On Error Resume Next
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\tch_sts.tab" + Chr(34)
    If mapinfo.eval("tableinfo(tch_sts,4)") = 26 Then
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

