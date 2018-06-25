VERSION 5.00
Begin VB.Form frmTACover 
   Caption         =   "覆盖合理性统计"
   ClientHeight    =   2805
   ClientLeft      =   4110
   ClientTop       =   1545
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
   Icon            =   "frmTACover.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   180
      TabIndex        =   2
      Top             =   150
      Width           =   4305
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1245
         TabIndex        =   5
         Text            =   "1"
         Top             =   1125
         Width           =   420
      End
      Begin VB.OptionButton Option1 
         Caption         =   "只显示"
         Height          =   375
         Index           =   1
         Left            =   345
         TabIndex        =   4
         Top             =   1095
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "显示占用小区的覆盖分布(Timing Advance)"
         Height          =   375
         Index           =   0
         Left            =   330
         TabIndex        =   3
         Top             =   540
         Value           =   -1  'True
         Width           =   3720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公里外的占用小区覆盖分布"
         Height          =   180
         Left            =   1740
         TabIndex        =   6
         Top             =   1185
         Width           =   2160
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      DragIcon        =   "frmTACover.frx":000C
      Height          =   320
      Left            =   2460
      TabIndex        =   1
      Top             =   2355
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      DragIcon        =   "frmTACover.frx":015E
      Height          =   320
      Left            =   1275
      TabIndex        =   0
      Top             =   2355
      Width           =   1080
   End
End
Attribute VB_Name = "frmTACover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim MyMsgs As String
    Dim MyRows As Integer, i As Integer
    Dim MyArray() As String
    Dim j As Integer
    Dim nn As Integer
    
    On Error Resume Next
    Me.Hide
    If Option1(0).Value Then
        mapinfo.do "select ta from " & tblname & " group by ta order by ta into mytemp"
        MyRows = mapinfo.eval("tableinfo(mytemp,8)")
        mapinfo.do "fetch first from mytemp"
        ReDim MyArray(MyRows - 1) As String
        For i = 0 To MyRows - 1
            MyArray(i) = mapinfo.eval("Mytemp.TA")
            mapinfo.do "fetch next from mytemp"
        Next
        mapinfo.do "close table mytemp"
        MyMsgs = "shade window FrontWindow() " + tblname + " With ta "
    Else
        mapinfo.do "select * from " & tblname & " where val(ta)>" & Format(Val(Text1.Text) * 1000 / 500) & " into TACover"
        MyRows = mapinfo.eval("tableinfo(TACover,8)")
        If MyRows = 0 Then
            MsgBox "该路段不存在" & Text1.Text & "公里以外的小区覆盖区域", 64, "提示"
            mapinfo.do "close table TACover"
            Unload Me
            Exit Sub
        End If
        mapinfo.do "select ta from TACover group by ta order by ta into mytemp"
        MyRows = mapinfo.eval("tableinfo(mytemp,8)")
        mapinfo.do "fetch first from mytemp"
        ReDim MyArray(MyRows - 1) As String
        For i = 0 To MyRows - 1
            MyArray(i) = mapinfo.eval("Mytemp.TA")
            mapinfo.do "fetch next from mytemp"
        Next
        mapinfo.do "close table mytemp"
        mapinfo.do "Add Map window FrontWindow() Layer TACover"
        MyMsgs = "shade window FrontWindow() TACover With ta "
    End If
    MyMsgs = MyMsgs + "values """" Symbol (63,14737632,8,""MapInfo Cartographic"",0,0) ,"
    MyMsgs = MyMsgs + "0 Symbol (63,65280,8,""MapInfo Cartographic"",0,0) , 1 Symbol (63,7585792,8,""MapInfo Cartographic"",0,0) ,"
    MyMsgs = MyMsgs + "2 Symbol (63,8421376,8,""MapInfo Cartographic"",0,0) ,3 Symbol (63,8388736,8,""MapInfo Cartographic"",0,0) ,"
    MyMsgs = MyMsgs + "4 Symbol (63,255,8,""MapInfo Cartographic"",0,0) ,5 Symbol (63,8432639,8,""MapInfo Cartographic"",0,0) ,"
    MyMsgs = MyMsgs + "6 Symbol (63,65535,8,""MapInfo Cartographic"",0,0) ,7 Symbol (63,16750640,8,""MapInfo Cartographic"",0,0) ,"
    MyMsgs = MyMsgs + "8 Symbol (63,16765088,8,""MapInfo Cartographic"",0,0),9 Symbol (63,16711935,8,""MapInfo Cartographic"",0,0),10 Symbol (63,16756952,8,""MapInfo Cartographic"",0,0),11 Symbol (63,8421376,8,""MapInfo Cartographic"",0,0),"
    
                   MyMsgs = MyMsgs + "12 Symbol (63,32896,8,""MapInfo Cartographic"",0,0),13 Symbol (63,16744576,8,""MapInfo Cartographic"",0,0),14 Symbol (63,8454016,8,""MapInfo Cartographic"",0,0),15 Symbol (63,8421631,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "16 Symbol (63,16744703,8,""MapInfo Cartographic"",0,0),17 Symbol (63,16777088,8,""MapInfo Cartographic"",0,0),18 Symbol (63,8454143,8,""MapInfo Cartographic"",0,0),19 Symbol (63,8405056,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "20 Symbol (63,4227136,8,""MapInfo Cartographic"",0,0),21 Symbol (63,4210816,8,""MapInfo Cartographic"",0,0),22 Symbol (63,8405120,8,""MapInfo Cartographic"",0,0),23 Symbol (63,8421440,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "24 Symbol (63,4227200,8,""MapInfo Cartographic"",0,0),25 Symbol (63,16761024,8,""MapInfo Cartographic"",0,0),26 Symbol (63,12648384,8,""MapInfo Cartographic"",0,0),27 Symbol (63,12632319,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "28 Symbol (63,16761087,8,""MapInfo Cartographic"",0,0),29 Symbol (63,16777152,8,""MapInfo Cartographic"",0,0),30 Symbol (63,12648447,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "31 Symbol (63,8413280,8,""MapInfo Cartographic"",0,0),32 Symbol (63,6324320,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "33 Symbol (63,6316160,8,""MapInfo Cartographic"",0,0),34 Symbol (63,8413312,8,""MapInfo Cartographic"",0,0),35 Symbol (63,8421472,8,""MapInfo Cartographic"",0,0),36 Symbol (63,6324352,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "37 Symbol (63,16711680,8,""MapInfo Cartographic"",0,0),38 Symbol (63,65280,8,""MapInfo Cartographic"",0,0),39 Symbol (63,255,8,""MapInfo Cartographic"",0,0),40 Symbol (63,16711935,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "41 Symbol (63,16776960,8,""MapInfo Cartographic"",0,0),42 Symbol (63,65535,8,""MapInfo Cartographic"",0,0),43 Symbol (63,8388608,8,""MapInfo Cartographic"",0,0),44 Symbol (63,32768,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "45 Symbol (63,128,8,""MapInfo Cartographic"",0,0),46 Symbol (63,8388736,8,""MapInfo Cartographic"",0,0),47 Symbol (63,8421376,8,""MapInfo Cartographic"",0,0),48 Symbol (63,32896,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "49 Symbol (63,16744576,8,""MapInfo Cartographic"",0,0),50 Symbol (63,8454016,8,""MapInfo Cartographic"",0,0),51 Symbol (63,8421631,8,""MapInfo Cartographic"",0,0),52 Symbol (63,16744703,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "53 Symbol (63,16777088,8,""MapInfo Cartographic"",0,0),54 Symbol (63,8454143,8,""MapInfo Cartographic"",0,0),55 Symbol (63,8405056,8,""MapInfo Cartographic"",0,0),56 Symbol (63,4227136,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "57 Symbol (63,4210816,8,""MapInfo Cartographic"",0,0),58 Symbol (63,8405120,8,""MapInfo Cartographic"",0,0),59 Symbol (63,8421440,8,""MapInfo Cartographic"",0,0),60 Symbol (63,4227200,8,""MapInfo Cartographic"",0,0),"
                   MyMsgs = MyMsgs + "61 Symbol (63,16761024,8,""MapInfo Cartographic"",0,0),62 Symbol (63,12648384,8,""MapInfo Cartographic"",0,0),63 Symbol (63,12632319,8,""MapInfo Cartographic"",0,0)"
    
    MyMsgs = MyMsgs + "default Symbol (63,16777215,8,""MapInfo Cartographic"",0,0)"
    mapinfo.do MyMsgs
    If legendid = 0 Then
        mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
        mapinfo.do "Create Legend From Window  Frontwindow()"
        legendid = mapinfo.eval("windowinfo(1009,12)")
    End If
    If Menu_Flag = 317 Then
        MyMsgs = " Title " + Chr(34) + "覆盖合理性统计 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
    Else
        MyMsgs = " Title " + Chr(34) + "Timing Advance分析 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
    End If
    
    If MyArray(0) = "" Then
        MyMsgs = MyMsgs + ",""IDLE"" display on"
        j = 1
    Else
        MyMsgs = MyMsgs + ",""IDLE"" display off"
        j = 0
    End If
    For i = 0 To 10
        If j > UBound(MyArray) Then
            Exit For
        End If
        If Format(i) = MyArray(j) Then
            If MyArray(j) = "0" Then
                MyMsgs = MyMsgs + ",""" & MyArray(j) & """ display on"
            Else
                MyMsgs = MyMsgs + ",""" & MyArray(j) & "  [" & Format(i * 500 / 1000, "0.0") & "公里以内]"" display on"
            End If
            j = j + 1
        Else
            MyMsgs = MyMsgs + ","""" display off"
        End If
        If i = 1 Then
            For nn = 10 To 19
                If j > UBound(MyArray) Then
                    Exit For
                End If
                If Format(nn) = MyArray(j) Then
                    MyMsgs = MyMsgs + ",""" & MyArray(j) & "  [" & Format(nn * 500 / 1000, "0.0") & "公里]"" display on"
                    j = j + 1
                Else
                    MyMsgs = MyMsgs + ","""" display off"
                End If
            Next
        ElseIf i = 2 Then
            For nn = 20 To 29
                If j > UBound(MyArray) Then
                    Exit For
                End If
                If Format(nn) = MyArray(j) Then
                    MyMsgs = MyMsgs + ",""" & MyArray(j) & "  [" & Format(nn * 500 / 1000, "0.0") & "公里]"" display on"
                    j = j + 1
                Else
                    MyMsgs = MyMsgs + ","""" display off"
                End If
            Next
        ElseIf i = 3 Then
            For nn = 30 To 39
                If j > UBound(MyArray) Then
                    Exit For
                End If
                If Format(nn) = MyArray(j) Then
                    MyMsgs = MyMsgs + ",""" & MyArray(j) & "  [" & Format(nn * 500 / 1000, "0.0") & "公里]"" display on"
                    j = j + 1
                Else
                    MyMsgs = MyMsgs + ","""" display off"
                End If
            Next
        ElseIf i = 4 Then
            For nn = 40 To 49
                If j > UBound(MyArray) Then
                    Exit For
                End If
                If Format(nn) = MyArray(j) Then
                    MyMsgs = MyMsgs + ",""" & MyArray(j) & "  [" & Format(nn * 500 / 1000, "0.0") & "公里]"" display on"
                    j = j + 1
                Else
                    MyMsgs = MyMsgs + ","""" display off"
                End If
            Next
        ElseIf i = 5 Then
            For nn = 50 To 59
                If j > UBound(MyArray) Then
                    Exit For
                End If
                If Format(nn) = MyArray(j) Then
                    MyMsgs = MyMsgs + ",""" & MyArray(j) & "  [" & Format(nn * 500 / 1000, "0.0") & "公里]"" display on"
                    j = j + 1
                Else
                    MyMsgs = MyMsgs + ","""" display off"
                End If
            Next
        ElseIf i = 6 Then
            For nn = 60 To 63
                If j > UBound(MyArray) Then
                    Exit For
                End If
                If Format(nn) = MyArray(j) Then
                    MyMsgs = MyMsgs + ",""" & MyArray(j) & "  [" & Format(nn * 500 / 1000, "0.0") & "公里]"" display on"
                    j = j + 1
                Else
                    MyMsgs = MyMsgs + ","""" display off"
                End If
            Next
        End If
    Next
    mapinfo.do "set legend window FrontWindow() Layer prev " & MyMsgs
    
    mapinfo.do "close table selection"
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If Menu_Flag = 991121 Then
        Caption = "Timing Advance分析"
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    On Error Resume Next
    If Option1(1).Value Then
        Text1.Enabled = True
    Else
        Text1.Enabled = False
    End If
End Sub
