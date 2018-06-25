VERSION 5.00
Begin VB.Form frmDataGraph 
   Caption         =   "数据统计图"
   ClientHeight    =   3780
   ClientLeft      =   5835
   ClientTop       =   1740
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataGraph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "统计范围（X轴）"
      Height          =   1575
      Left            =   225
      TabIndex        =   8
      Top             =   105
      Width           =   4185
      Begin VB.OptionButton Option2 
         Caption         =   "做主、邻小区"
         Height          =   255
         Index           =   2
         Left            =   1965
         TabIndex        =   12
         Top             =   705
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IntegralHeight  =   0   'False
         ItemData        =   "frmDataGraph.frx":000C
         Left            =   450
         List            =   "frmDataGraph.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1065
         Width           =   2160
      End
      Begin VB.OptionButton Option2 
         Caption         =   "做主小区"
         Height          =   255
         Index           =   1
         Left            =   195
         TabIndex        =   10
         Top             =   705
         Width           =   1245
      End
      Begin VB.OptionButton Option2 
         Caption         =   "以时间分布"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   9
         Top             =   345
         Value           =   -1  'True
         Width           =   1290
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      DragIcon        =   "frmDataGraph.frx":0010
      Height          =   320
      Left            =   2280
      TabIndex        =   7
      Top             =   3390
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      DragIcon        =   "frmDataGraph.frx":0162
      Height          =   320
      Left            =   1080
      TabIndex        =   6
      Top             =   3390
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "统计内容（Y轴）"
      Height          =   1515
      Left            =   225
      TabIndex        =   0
      Top             =   1740
      Width           =   4185
      Begin VB.OptionButton Option1 
         Caption         =   "TA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2310
         TabIndex        =   13
         Top             =   1095
         Width           =   1290
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tx_Power"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   495
         TabIndex        =   5
         Top             =   1095
         Width           =   1290
      End
      Begin VB.OptionButton Option1 
         Caption         =   "RxQual_s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   495
         TabIndex        =   4
         Top             =   735
         Width           =   1290
      End
      Begin VB.OptionButton Option1 
         Caption         =   "RxQual_f"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2310
         TabIndex        =   3
         Top             =   735
         Width           =   1290
      End
      Begin VB.OptionButton Option1 
         Caption         =   "RxLev_s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   495
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1290
      End
      Begin VB.OptionButton Option1 
         Caption         =   "RxLev_f"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2310
         TabIndex        =   1
         Top             =   360
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmDataGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelCIValue() As String
'Dim AllCIValue() As String
'Dim mySelTbl As String
Dim QueryName As String

Private Sub Command1_Click()
    Dim MyParameter As String
    Dim i As Integer
    Dim WinId As Long
    Dim ViewAllCi As Boolean
    Dim MyColor As Long
    Dim MyTitle As String
    Dim MyType As String, MySubtitle As String
    
    On Error Resume Next
    If MapGraphflag Then
       Unload frmMapGraph
    End If
    If Option1(0).Value Then
       MyParameter = "RxLev_f"
       MySubtitle = "RxLev_f"
       MyColor = 32896
       MyType = "Line"
    ElseIf Option1(1).Value Then
       MyParameter = "RxLev_s"
       MySubtitle = "RxLev_s"
       MyColor = 32896
       MyType = "Line"
    ElseIf Option1(2).Value Then
                  If mapinfo.eval("ColumnInfo(" & QueryName & ", ""rxqual_s"", 3)") = 1 Then    'Character
                     MyParameter = "val(rxqual_f)"
                  Else
                     MyParameter = "rxqual_f"
                  End If
    
       'MyParameter = "RxQual_f"
       MySubtitle = "RxQual_f"
       MyColor = 16744448
       MyType = "Bar"
    ElseIf Option1(3).Value Then
                  If mapinfo.eval("ColumnInfo(" & QueryName & ", ""rxqual_s"", 3)") = 1 Then    'Character
                     MyParameter = "val(rxqual_s)"
                  Else
                     MyParameter = "rxqual_s"
                  End If
       
       'MyParameter = "RxQual_s"
       MySubtitle = "RxQual_s"
       MyColor = 16744448
       MyType = "Bar"
    ElseIf Option1(4).Value Then
       MyParameter = "val(Tx_Power)"
       MySubtitle = "Tx_Power"
       MyColor = 12632064
       'MyType = "Line"
       MyType = "Bar"
    Else
       MyParameter = "val(Ta)"
       MySubtitle = "TA"
       MyColor = &H808080
       'MyType = "Line"
       MyType = "Bar"
    End If
    mapinfo.do "Set Next Document Parent " & frmMapGraph.hWnd & " Style 1"
    If Option2(1).Value Then
       If Combo1.Text <> "全部" Then
          mapinfo.do "select * from " & QueryName & " where ci_serv = """ & Left(Combo1.Text, InStr(Combo1.Text, "[") - 1) & """ into " & QueryName
          MyTitle = Combo1.Text & "做主小区分布" & MySubtitle & "统计图"
       Else
          mapinfo.do "select ci_serv,avg(" & MyParameter & ") from " & QueryName & " where ci_serv <> """" group by ci_serv into " & QueryName
          MyTitle = "全部主小区" & MySubtitle & "对比统计图"
          ViewAllCi = True
          MyType = "Bar"
       End If
    ElseIf Option2(2).Value Then
       'Option2(2) is invisible
    Else
       MyTitle = "以时间分布" & MySubtitle & "统计图"
    End If
    mapinfo.do "select * from " & QueryName
    If ViewAllCi Then
       mapinfo.do "Graph ci_serv,col2 from " & QueryName & " width 2 Units ""in"" height 1 Units ""in"""
       If Option1(5).Value Then
          mapinfo.do "Set Graph Type Bar Series 2 Pen(1, 1," & Format(MyColor) & ") Brush (2," & Format(MyColor) & "," & Format(MyColor) & ") Line (1,2,0," & Format(MyColor) & ") Symbol(34," & Format(MyColor) & ",6) Stacked Off Overlapped Off Droplines Off Rotated Off Show3d Off Gutter 60 Title """ & MyTitle & """ Value Axis Major Unit 1 Minor Unit 1 Title """ & MySubtitle & """"
       Else
          mapinfo.do "Set Graph Type Bar Series 2 Pen(1, 1," & Format(MyColor) & ") Brush (2," & Format(MyColor) & "," & Format(MyColor) & ") Line (1,2,0," & Format(MyColor) & ") Symbol(34," & Format(MyColor) & ",6) Stacked Off Overlapped Off Droplines Off Rotated Off Show3d Off Gutter 60 Title """ & MyTitle & """ Value Axis Title """ & MySubtitle & """"
       End If
       mapinfo.do "Set Window FrontWindow() Width 10 Units ""cm"" Height 5 units ""cm"""
    ElseIf Option2(2).Value Then
       'Option2(2) is invisible
    Else
       mapinfo.do "Graph Time," & MyParameter & " from " & QueryName & " width 2 height 1"
       If Option1(5).Value Then
          mapinfo.do "Set Graph Type " & MyType & " Series 2 Pen(1, 1," & Format(MyColor) & ") Brush (2," & Format(MyColor) & "," & Format(MyColor) & ") Line (1,2,0," & Format(MyColor) & ") Symbol(34," & Format(MyColor) & ",6) Stacked Off Overlapped Off Droplines Off Rotated Off Show3d Off Gutter 0 Title """ & MyTitle & """ Value Axis Major Unit 1 Minor Unit 1 Title """ & MySubtitle & """"
       Else
          mapinfo.do "Set Graph Type " & MyType & " Series 2 Pen(1, 1," & Format(MyColor) & ") Brush (2," & Format(MyColor) & "," & Format(MyColor) & ") Line (1,2,0," & Format(MyColor) & ") Symbol(34," & Format(MyColor) & ",6) Stacked Off Overlapped Off Droplines Off Rotated Off Show3d Off Gutter 0 Title """ & MyTitle & """ Value Axis Title """ & MySubtitle & """"
       End If
       mapinfo.do "Set Window FrontWindow() Width 10 Units ""cm"" Height 5 units ""cm"""
    End If
    frmMapGraph.WindowState = 0
    frmMapGraph.Height = 3000
    frmMapGraph.Width = 5000
    frmMapGraph.Move 6300, 4300, 5000, 3000
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim SelCIRows As Integer
    Dim i As Integer
    Dim MyCellName As String
    Dim CellIsOpen As Boolean
    Dim MyTableNum As Integer
    
    On Error Resume Next
    'mySelTbl = mapinfo.eval("selectionInfo(1)")
    MyTableNum = mapinfo.eval("NumTables()")
    For i = 1 To MyTableNum
        If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "CELL" Then
           CellIsOpen = True
           Exit For
        End If
    Next
    QueryName = mapinfo.eval("selectionInfo(2)")
    mapinfo.do "Select * from " & QueryName & " where Ci_serv<>"""" group by ci_serv into mytemp"
    SelCIRows = mapinfo.eval("tableinfo(mytemp,8)")
    mapinfo.do "fetch first from mytemp"
    ReDim SelCIValue(0 To SelCIRows - 1) As String
    For i = 0 To SelCIRows - 1
        SelCIValue(i) = mapinfo.eval("mytemp.ci_serv")
        If CellIsOpen Then
           MyCellName = Findcell(SelCIValue(i))
        End If
        Combo1.AddItem SelCIValue(i) & "[" & MyCellName & "]"
        mapinfo.do "fetch next from mytemp"
    Next
    Combo1.AddItem "全部"
    mapinfo.do "close table mytemp"
    Combo1.ListIndex = 0
    'Combo1.Text = SelCIValue(0)

End Sub

Private Sub Option2_Click(Index As Integer)
    
    On Error Resume Next
    If Option2(0).Value Then
       Combo1.Enabled = False
    Else
        Combo1.Enabled = True
    End If
    
End Sub
