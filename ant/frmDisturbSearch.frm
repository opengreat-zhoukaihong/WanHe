VERSION 5.00
Begin VB.Form frmDisturbSearch 
   Caption         =   "指定范围内同邻频查找"
   ClientHeight    =   2400
   ClientLeft      =   4200
   ClientTop       =   3225
   ClientWidth     =   3705
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDisturbSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3705
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      DragIcon        =   "frmDisturbSearch.frx":000C
      Height          =   320
      Left            =   690
      TabIndex        =   3
      Top             =   1920
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      DragIcon        =   "frmDisturbSearch.frx":015E
      Height          =   320
      Left            =   1935
      TabIndex        =   2
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3165
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   675
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "请输入主小区载频："
         Height          =   180
         Index           =   6
         Left            =   315
         TabIndex        =   1
         Top             =   735
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmDisturbSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim QueryName As String
Dim SelBsic() As String
Dim SelCi() As String
Dim mySelTbl As String
Dim MyDisturbBcch() As String
Dim MyDisturbBsic() As String

Private Sub Command1_Click()
    Dim CellIsOpen As Boolean
    Dim MyTableNum As Integer
    Dim MyRows As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim MyCellName As String, MyCival As String
    Dim itmX As ListItem
    Dim MyBcchtemp As Integer, Mybsictemp As Integer
    Dim MyComboText As String
    Dim MyComboIndex As Integer
    Dim Non1 As Boolean, Non2 As Boolean
    Dim MyAddFlag As Boolean
    Dim mm As Integer
    
    On Error Resume Next
    If Combo1.Text = "" Then
       Unload Me
       Exit Sub
    End If
    MyComboText = Combo1.Text
    MyComboIndex = Combo1.ListIndex
    Unload Me
    MyTableNum = mapinfo.eval("NumTables()")
    For i = 1 To MyTableNum
        If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "CELL" Then
           CellIsOpen = True
           Exit For
        End If
    Next
    mapinfo.do "select * from " & QueryName & " where bcch_serv = " & MyComboText & " and bsic_serv <> " & SelBsic(MyComboIndex) & " or bcch_n1=" & MyComboText & " and bsic_n1<> " & SelBsic(MyComboIndex) & " and bsic_n1<>99 or bcch_n2=" & MyComboText & " and bsic_n2<> " & SelBsic(MyComboIndex) & " and bsic_n2<>99 or bcch_n3=" & MyComboText & " and bsic_n3<> " & SelBsic(MyComboIndex) & " and bsic_n3<>99 or bcch_n4=" & MyComboText & " and bsic_n4<> " & SelBsic(MyComboIndex) & " and bsic_n4<>99 or bcch_n5=" & MyComboText & " and bsic_n5<> " & SelBsic(MyComboIndex) & " and bsic_n5<>99 or bcch_n6=" & MyComboText & " and bsic_n6<> " & SelBsic(MyComboIndex) & " and bsic_n6<>99 into Search1"
    MyRows = mapinfo.eval("tableinfo(Search1,8)")
    If MyRows > 0 Then
       ReDim MyDisturbBcch(MyRows) As String
       ReDim MyDisturbBsic(MyRows) As String
       mapinfo.do "Add Map window Frontwindow() Layer Search1"
       mapinfo.do "shade window FrontWindow() Search1 with bcch_serv ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 1: 1800 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0)"
                 If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                 End If
                 'Msg = " Title " + Chr(34) + "本网同频干扰查找 " + mySelTbl + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "主小区载频：" + MyComboText + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""同频干扰源"" display on "
                 Msg = " Title " + Chr(34) + "指定范围内同频查找 " + mySelTbl + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "主小区载频：" + MyComboText + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""存在地点"" display on "
                 mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
       
       If MessageId1 = 0 Then
          'Load frmMessage1
          frmMessage1.Show
          'frmMessage1.Move 6300, 4300, 3540, 3135
          SetWindowPos frmMessage1.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
       Else
          frmMessage1.ListView1.ListItems.Clear
       End If
       mapinfo.do "select * from Search1 where bcch_serv=" & MyComboText & " and bsic_serv <> " & SelBsic(MyComboIndex) & " group by ci_serv into mytemp1"
       MyRows = mapinfo.eval("tableinfo(mytemp1,8)")
       mapinfo.do "fetch first from mytemp1"
       MyCival = SelCi(MyComboIndex)
       If CellIsOpen Then
          Call SearchCellName(0, 0, 0, 0, MyCellName, MyCival, "")
       Else
          MyCellName = ""
       End If
       frmMessage1.Label1(2).Caption = MyCellName
       frmMessage1.Label1(3).Caption = MyComboText
       frmMessage1.Label1(4).Caption = SelBsic(MyComboIndex)
       k = 0
       If MyRows > 0 Then
          For i = 1 To MyRows
              MyAddFlag = False
              MyCival = mapinfo.eval("mytemp1.ci_serv")
              If CellIsOpen Then
                 Call SearchCellName(0, 0, 0, 0, MyCellName, MyCival, "")
              Else
                 MyCellName = ""
              End If
              If MyCellName <> "" Then
                 mapinfo.do "select * from cell where cell_name = """ & MyCellName & """ into Lin"
                 If mapinfo.eval("mytemp1.bcch_serv") = mapinfo.eval("Lin.arfcn") And mapinfo.eval("mytemp1.bsic_serv") = mapinfo.eval("Lin.bsic") Then
                    MyAddFlag = True
                 End If
                 mapinfo.do "close table Lin"
              Else
                 If k > 0 Then
                    If mapinfo.eval("mytemp1.bcch_serv") <> MyDisturbBcch(k - 1) Or mapinfo.eval("mytemp1.bsic_serv") <> MyDisturbBsic(k - 1) Then
                        MyAddFlag = True
                    End If
                 Else
                    MyAddFlag = True
                 End If
              End If
              If MyAddFlag Then
                Set itmX = frmMessage1.ListView1.ListItems.ADD(, , CStr(MyCellName))
                itmX.SubItems(1) = mapinfo.eval("mytemp1.bcch_serv")
                itmX.SubItems(2) = mapinfo.eval("mytemp1.bsic_serv")
                MyDisturbBcch(k) = mapinfo.eval("mytemp1.bcch_serv")
                MyDisturbBsic(k) = mapinfo.eval("mytemp1.bsic_serv")
                k = k + 1
              End If
              mapinfo.do "fetch next from mytemp1"
          Next
       End If
       mapinfo.do "select * from Search1 where not(bcch_serv=" & MyComboText & " and bsic_serv <> " & SelBsic(MyComboIndex) & ") into mytemp2"
       MyRows = mapinfo.eval("tableinfo(mytemp2,8)")
       mapinfo.do "fetch first from mytemp2"
       If MyRows > 0 Then
          For i = 1 To MyRows
              For j = 1 To 6
                  If mapinfo.eval("mytemp2.bcch_n" & Format(j)) = MyComboText And mapinfo.eval("mytemp2.bsic_n" & Format(j)) <> SelBsic(MyComboIndex) And mapinfo.eval("mytemp2.bsic_n" & Format(j)) <> 99 Then
                     MyBcchtemp = mapinfo.eval("mytemp2.bcch_n" & Format(j))
                     Mybsictemp = mapinfo.eval("mytemp2.bsic_n" & Format(j))
                     For k = 0 To UBound(MyDisturbBcch)
                         If MyDisturbBcch(k) = "" Then
                            Exit For
                         End If
                         If MyBcchtemp = MyDisturbBcch(k) And Mybsictemp = MyDisturbBsic(k) Then
                            GoTo PPP
                         End If
                     Next
                    MyDisturbBcch(k) = MyBcchtemp
                    MyDisturbBsic(k) = Mybsictemp
                    If CellIsOpen Then
                       Call SearchCellName(Mybsictemp, MyBcchtemp, mapinfo.eval("mytemp2.lon"), mapinfo.eval("mytemp2.lat"), MyCellName, "", "")
                    Else
                       MyCellName = ""
                    End If
                    Set itmX = frmMessage1.ListView1.ListItems.ADD(, , CStr(MyCellName))
                    itmX.SubItems(1) = MyBcchtemp
                    itmX.SubItems(2) = Mybsictemp
                     Exit For
                  End If
              Next
PPP:
             mapinfo.do "fetch next from mytemp2"
          Next
       End If
       
                 
'                 If MessageId1 = 0 Then
'                    mapinfo.do "Set Application Window " & MDIMain.hwnd
'                    mapinfo.do "Open Window Message"
'                    mapinfo.do "Set Window Message Title ""Lee"""
  'Font ("Helv", 1, 10, BLUE)    ' Helvetica bold...
  'Position (0.25, 0.25)        ' place in upper left
  'Width 3.0   ' make window 3" wide
  'Height 1#   ' make window 1" high
'                      mapinfo.do "Print Chr$(12)"
'                      mapinfo.do "print ""3255"""
'                      MessageId1 = mapinfo.eval("windowinfo(1003,12)")
'                 End If
    Else
        Non1 = True
        mapinfo.do "close table search1"
    End If
    

'---------------------邻频---------------------
    
    mapinfo.do "select * from " & QueryName & " where bcch_serv = " & Format(Val(MyComboText) - 1) & " or bcch_serv= " & Format(Val(MyComboText) + 1) & " or bsic_n1<>99 and (bcch_n1=" & Format(Val(MyComboText) - 1) & " or bcch_n1=" & Format(Val(MyComboText) + 1) & ") or bsic_n2<>99 and (bcch_n2=" & Format(Val(MyComboText) - 1) & " or bcch_n2=" & Format(Val(MyComboText) + 1) & ") or bsic_n3<>99 and (bcch_n3=" & Format(Val(MyComboText) - 1) & " or bcch_n3=" & Format(Val(MyComboText) + 1) & ") or bsic_n4<>99 and (bcch_n4=" & Format(Val(MyComboText) - 1) & " or bcch_n4=" & Format(Val(MyComboText) + 1) & ") or bsic_n5<>99 and (bcch_n5=" & Format(Val(MyComboText) - 1) & " or bcch_n5=" & Format(Val(MyComboText) + 1) & ") or bsic_n6<>99 and (bcch_n6=" & Format(Val(MyComboText) - 1) & " or bcch_n6=" & Format(Val(MyComboText) + 1) & ") into Search2"
    MyRows = mapinfo.eval("tableinfo(Search2,8)")
    If MyRows > 0 Then
       ReDim MyDisturbBcch(MyRows) As String
       ReDim MyDisturbBsic(MyRows) As String
       mapinfo.do "Add Map window Frontwindow() Layer Search2"
       mapinfo.do "shade window FrontWindow() Search2 with bcch_serv ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 1: 1800 Symbol (39,255,8,""MapInfo Cartographic"",0,0)"
                 If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                 End If
                 'Msg = " Title " + Chr(34) + "本网邻频干扰查找 " + mySelTbl + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "主小区载频：" + MyComboText + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""邻频干扰源"" display on "
                 Msg = " Title " + Chr(34) + "指定范围内邻频查找 " + mySelTbl + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "主小区载频：" + MyComboText + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""存在地点"" display on "
                 mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
       
       If MessageId2 = 0 Then
          frmMessage2.Show
          'frmMessage2.Move 6300, 4300, 3540, 3135
          SetWindowPos frmMessage2.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
       Else
          frmMessage2.ListView1.ListItems.Clear
       End If
       
       mapinfo.do "select * from Search2 where bcch_serv = " & Format(Val(MyComboText) - 1) & " or bcch_serv= " & Format(Val(MyComboText) + 1) & " group by ci_serv into mytemp1"
       MyRows = mapinfo.eval("tableinfo(mytemp1,8)")
       mapinfo.do "fetch first from mytemp1"
       MyCival = SelCi(MyComboIndex)
       If CellIsOpen Then
          Call SearchCellName(0, 0, 0, 0, MyCellName, MyCival, "")
       Else
          MyCellName = ""
       End If
       frmMessage2.Label1(2).Caption = MyCellName
       frmMessage2.Label1(3).Caption = MyComboText
       frmMessage2.Label1(4).Caption = SelBsic(MyComboIndex)
       mm = 0
       If MyRows > 0 Then
          For i = 1 To MyRows
              MyAddFlag = False
              MyCival = mapinfo.eval("mytemp1.ci_serv")
              MyBcchtemp = mapinfo.eval("mytemp1.bcch_serv")
              Mybsictemp = mapinfo.eval("mytemp1.bsic_serv")
              
              If CellIsOpen Then
                 Call SearchCellName(0, 0, 0, 0, MyCellName, MyCival, "")
              Else
                 MyCellName = ""
              End If
              If MyCellName <> "" Then
                 mapinfo.do "select * from cell where cell_name = """ & MyCellName & """ into Lin"
                 If mapinfo.eval("mytemp1.bcch_serv") = mapinfo.eval("Lin.arfcn") And mapinfo.eval("mytemp1.bsic_serv") = mapinfo.eval("Lin.bsic") Then
                    MyAddFlag = True
                 End If
                 mapinfo.do "close table Lin"
              Else
                 If mm > 0 Then
                    If mapinfo.eval("mytemp1.bcch_serv") <> MyDisturbBcch(mm - 1) Or mapinfo.eval("mytemp1.bsic_serv") <> MyDisturbBsic(mm - 1) Then
                        MyAddFlag = True
                    End If
                 Else
                    MyAddFlag = True
                 End If
              End If
              If MyAddFlag Then
                     For k = 0 To UBound(MyDisturbBcch)
                         If MyDisturbBcch(k) = "" Then
                            Exit For
                         End If
                         If MyBcchtemp = MyDisturbBcch(k) And Mybsictemp = MyDisturbBsic(k) Then
                            GoTo PPP2
                         End If
                     Next
              End If
              Set itmX = frmMessage2.ListView1.ListItems.ADD(, , CStr(MyCellName))
              itmX.SubItems(1) = mapinfo.eval("mytemp1.bcch_serv")
              itmX.SubItems(2) = mapinfo.eval("mytemp1.bsic_serv")
              MyDisturbBcch(i - 1) = mapinfo.eval("mytemp1.bcch_serv")
              MyDisturbBsic(i - 1) = mapinfo.eval("mytemp1.bsic_serv")
              mm = mm + 1
PPP2:
              mapinfo.do "fetch next from mytemp1"
          Next
       End If
       mapinfo.do "select * from Search2 where not(bcch_serv = " & Format(Val(MyComboText) - 1) & " or bcch_serv= " & Format(Val(MyComboText) + 1) & ") into mytemp2"
       MyRows = mapinfo.eval("tableinfo(mytemp2,8)")
       mapinfo.do "fetch first from mytemp2"
       If MyRows > 0 Then
          For i = 1 To MyRows
              For j = 1 To 6
                  If (mapinfo.eval("mytemp2.bcch_n" & Format(j)) = Val(MyComboText - 1) Or mapinfo.eval("mytemp2.bcch_n" & Format(j)) = Val(MyComboText + 1)) And mapinfo.eval("mytemp2.bsic_n" & Format(j)) <> 99 Then
                     MyBcchtemp = mapinfo.eval("mytemp2.bcch_n" & Format(j))
                     Mybsictemp = mapinfo.eval("mytemp2.bsic_n" & Format(j))
                     For k = 0 To UBound(MyDisturbBcch)
                         If MyDisturbBcch(k) = "" Then
                            Exit For
                         End If
                         If MyBcchtemp = MyDisturbBcch(k) And Mybsictemp = MyDisturbBsic(k) Then
                            GoTo PPP1
                         End If
                     Next
                    MyDisturbBcch(k) = MyBcchtemp
                    MyDisturbBsic(k) = Mybsictemp
                    If CellIsOpen Then
                       Call SearchCellName(Mybsictemp, MyBcchtemp, mapinfo.eval("mytemp2.lon"), mapinfo.eval("mytemp2.lat"), MyCellName, "", "")
                    Else
                       MyCellName = ""
                    End If
                    Set itmX = frmMessage2.ListView1.ListItems.ADD(, , CStr(MyCellName))
                    itmX.SubItems(1) = MyBcchtemp
                    itmX.SubItems(2) = Mybsictemp
                     Exit For
                  End If
              Next
PPP1:
             mapinfo.do "fetch next from mytemp2"
          Next
       End If
    Else
       Non2 = True
       mapinfo.do "close table search2"
    End If
    If Non1 And Non2 Then
       MsgBox "不存在本网同、邻频干扰", 64, "提示"
    ElseIf Non1 Then
       MsgBox "不存在本网同频干扰", 64, "提示"
    ElseIf Non2 Then
       MsgBox "不存在本网邻频干扰", 64, "提示"
    End If
    'Unload Me
    
End Sub

Private Sub Command2_Click()
    
    On Error Resume Next
    Unload Me

End Sub

Private Sub Form_Load()
    Dim SelCIRows As Integer
    Dim i As Integer
    
    On Error Resume Next
    mySelTbl = mapinfo.eval("selectionInfo(1)")
    QueryName = mapinfo.eval("selectionInfo(2)")
    mapinfo.do "Select * from " & QueryName & " where bcch_serv>0 group by bcch_serv into mytemp"
    SelCIRows = mapinfo.eval("tableinfo(mytemp,8)")
    If SelCIRows = 0 Then
       Exit Sub
    End If
    ReDim SelBsic(SelCIRows - 1) As String
    ReDim SelCi(SelCIRows - 1) As String
    mapinfo.do "fetch first from mytemp"
    For i = 0 To SelCIRows - 1
        Combo1.AddItem mapinfo.eval("mytemp.bcch_serv")
        SelBsic(i) = mapinfo.eval("mytemp.bsic_serv")
        SelCi(i) = mapinfo.eval("mytemp.ci_serv")
        mapinfo.do "fetch next from mytemp"
    Next
    mapinfo.do "close table mytemp"
    Combo1.ListIndex = 0

End Sub
