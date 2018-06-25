VERSION 5.00
Begin VB.Form frmStrongNcell 
   Caption         =   "功率预算切换统计"
   ClientHeight    =   3345
   ClientLeft      =   6945
   ClientTop       =   555
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStrongNcell.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   320
      Left            =   2025
      TabIndex        =   2
      Top             =   2925
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   320
      Left            =   810
      TabIndex        =   1
      Top             =   2925
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Height          =   2640
      Left            =   225
      TabIndex        =   0
      Top             =   90
      Width           =   3420
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   1725
         TabIndex        =   16
         Text            =   "30"
         Top             =   1875
         Width           =   480
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   1725
         TabIndex        =   13
         Text            =   "5"
         Top             =   1575
         Width           =   480
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   1
         Left            =   1725
         TabIndex        =   10
         Text            =   "5"
         Top             =   1275
         Width           =   480
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   0
         Left            =   1725
         TabIndex        =   6
         Text            =   "5"
         Top             =   975
         Width           =   480
      End
      Begin VB.CheckBox Check2 
         Caption         =   "DCS"
         Height          =   300
         Left            =   1965
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "GSM"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   255
         Value           =   1  'Checked
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tx_Power最大"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   9
         Left            =   570
         TabIndex        =   18
         Top             =   2265
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dBm"
         Enabled         =   0   'False
         Height          =   180
         Index           =   8
         Left            =   2280
         TabIndex        =   17
         Top             =   1905
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DCS 到 GSM"
         Enabled         =   0   'False
         Height          =   180
         Index           =   7
         Left            =   735
         TabIndex        =   15
         Top             =   1905
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dBm"
         Enabled         =   0   'False
         Height          =   180
         Index           =   6
         Left            =   2280
         TabIndex        =   14
         Top             =   1605
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DCS 到 DCS"
         Enabled         =   0   'False
         Height          =   180
         Index           =   5
         Left            =   735
         TabIndex        =   12
         Top             =   1605
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dBm"
         Height          =   180
         Index           =   3
         Left            =   2280
         TabIndex        =   11
         Top             =   1305
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GSM 到 DCS"
         Height          =   180
         Index           =   1
         Left            =   735
         TabIndex        =   9
         Top             =   1305
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "相邻小区场强>主小区场强"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   4
         Left            =   570
         TabIndex        =   8
         Top             =   675
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "dBm"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   7
         Top             =   1005
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GSM 到 GSM"
         Height          =   180
         Index           =   0
         Left            =   735
         TabIndex        =   3
         Top             =   1005
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmStrongNcell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Dim i As Integer
    
    On Error Resume Next
    If Check1.Value = 1 Then
        For i = 0 To 3
            Label1(i).Enabled = True
        Next
        Text1(0).Enabled = True
        Text1(1).Enabled = True
    Else
        For i = 0 To 3
            Label1(i).Enabled = False
        Next
        Text1(0).Enabled = False
        Text1(1).Enabled = False
    End If
End Sub

Private Sub Check2_Click()
    Dim i As Integer
    
    On Error Resume Next
    If Check2.Value = 1 Then
        For i = 5 To 8
            Label1(i).Enabled = True
        Next
        Text1(2).Enabled = True
        Text1(3).Enabled = True
    Else
        For i = 5 To 8
            Label1(i).Enabled = False
        Next
        Text1(2).Enabled = False
        Text1(3).Enabled = False
    End If

End Sub

Private Sub Command1_Click()
    Dim MyValue As Integer
    Dim StrongNcellRow As Integer
    Dim i As Integer, j As Integer
    Dim mytemp1 As Integer, mytemp2 As Boolean, mytemp3 As Integer
    Dim NonGSM As Boolean, NonDCS As Boolean
    
    On Error Resume Next
    Me.Hide
    For i = 1 To mapinfo.eval("NumTables()")
        If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "STRONGNCELL_1" Then
            mapinfo.do "close table strongncell_1"
            Exit For
        End If
    Next
    For i = 1 To mapinfo.eval("NumTables()")
        If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "STRONGNCELL_2" Then
            mapinfo.do "close table StrongNcell_2"
            Exit For
        End If
    Next
    If Check1.Value = 1 Then
        'mapinfo.do "select * from " & tblname & " where tx_power<6 and (bcch_serv<124 and (bsic_n1<>99 and bcch_n1<124 and (rxlev_n1-rxlev_s>" & Text1(0).Text & ") or bsic_n1<>99 and bcch_n1>123 and (rxlev_n1-rxlev_s>" & Text1(1).Text & ") or bsic_n2<>99 and bcch_n2<124 and (rxlev_n2-rxlev_s>" & Text1(0).Text & ") or bsic_n2<>99 and bcch_n2>123 and (rxlev_n2-rxlev_s>" & Text1(1).Text & ") or bsic_n3<>99 and bcch_n3<124 and (rxlev_n3-rxlev_s>" & Text1(0).Text & ") or bsic_n3 <>99 and bcch_n3>123 and (rxlev_n3-rxlev_s>" & Text1(1).Text & ") or bsic_n4 <>99 and bcch_n4<124 and (rxlev_n4-rxlev_s>" & Text1(0).Text & ") or bsic_n4 <>99 and bcch_n4>123 and (rxlev_n4-rxlev_s>" & Text1(1).Text & ") or bsic_n5 <>99 and bcch_n5<124 and (rxlev_n5-rxlev_s>" & Text1(0).Text & ") or bsic_n5 <>99 and bcch_n5>123 and (rxlev_n5-rxlev_s>" & Text1(1).Text & ") or bsic_n6 <>99 and bcch_n6<124 and (rxlev_n6-rxlev_s>" & Text1(0).Text & ") or bsic_n6<>99 and bcch_n6>123 and (rxlev_n6-rxlev_s>" & Text1(1).Text & ")) into MYTEMP"
        mapinfo.do "select * from " & tblname & " where tx_power <>"""" and val(tx_power)<6 and bcch_serv<124 into Mytemp"
        mapinfo.do "select * from Mytemp where bcch_n1<124 and (rxlev_n1-rxlev_s>" & Text1(0).Text & ") or bcch_n1>123 and (rxlev_n1-rxlev_s>" & Text1(1).Text & ") or bcch_n2<124 and (rxlev_n2-rxlev_s>" & Text1(0).Text & ") or bcch_n2>123 and (rxlev_n2-rxlev_s>" & Text1(1).Text & ") or bcch_n3<124 and (rxlev_n3-rxlev_s>" & Text1(0).Text & ") or bcch_n3>123 and (rxlev_n3-rxlev_s>" & Text1(1).Text & ") or bcch_n4<124 and (rxlev_n4-rxlev_s>" & Text1(0).Text & ") or bcch_n4>123 and (rxlev_n4-rxlev_s>" & Text1(1).Text & ") or bcch_n5<124 and (rxlev_n5-rxlev_s>" & Text1(0).Text & ") or bcch_n5>123 and (rxlev_n5-rxlev_s>" & Text1(1).Text & ") or bcch_n6<124 and (rxlev_n6-rxlev_s>" & Text1(0).Text & ") or bcch_n6>123 and (rxlev_n6-rxlev_s>" & Text1(1).Text & ") into MYTEMP"
        mapinfo.do "commit table mytemp as " + Chr(34) + Gsm_Path + "\user\strongncell_1.tab" + Chr(34)
        mapinfo.do "open table " + Chr(34) + Gsm_Path + "\user\strongncell_1.tab" + Chr(34)
        StrongNcellRow = Val(mapinfo.eval("tableinfo(strongncell_1,8)"))
        If StrongNcellRow = 0 Then
            NonGSM = True
            mapinfo.do "close table strongncell_1"
            GoTo next1
            
        End If
        mapinfo.do "fetch first from strongncell_1"
        For i = 1 To StrongNcellRow
            mytemp1 = mapinfo.eval("strongncell_1.rxlev_s")
            mytemp2 = False
            mytemp3 = 0
            For j = 1 To 6
            If ((mapinfo.eval("strongncell_1.rxlev_n" & Format(j)) - mytemp1 > Val(Text1(0))) And mapinfo.eval("strongncell_1.bsic_n" & Format(j)) <> 99 And mapinfo.eval("strongncell_1.bcch_n" & Format(j)) < 124) Or ((mapinfo.eval("strongncell_1.rxlev_n" & Format(j)) - mytemp1 > Val(Text1(1))) And mapinfo.eval("strongncell_1.bsic_n" & Format(j)) <> 99 And mapinfo.eval("strongncell_1.bcch_n" & Format(j)) > 124) Then
               mytemp3 = mytemp3 + 1
               If Not mytemp2 Then
                  mapinfo.do "UPDATE strongncell_1 set rxlev_f_2= " & Format(mapinfo.eval("strongncell_1.rxlev_n" & Format(j)) - mytemp1) & " WHERE ROWID=" & i
                  mytemp2 = True
               End If
            End If
        Next
        If mytemp3 = 0 Then
           mapinfo.do "UPDATE strongncell_1 set rxlev_f_2= 0 WHERE ROWID=" & i
        Else
           mapinfo.do "UPDATE strongncell_1 set ncell_num= " & Format(mytemp3) & " WHERE ROWID=" & i
        End If
        mapinfo.do "fetch next from strongncell_1"
    Next
            mapinfo.do "Create Map For strongncell_1 CoordSys Earth Projection 1, 0"
            mapinfo.do "Set Style Symbol MakeSymbol(33,0,2)" '
            mapinfo.do "update strongncell_1 set Obj= CreatePoint(Lon, Lat)"
            mapinfo.do "select * from strongncell_1 where rxlev_f_2>0 into strongncell1"
            mapinfo.do "Add Map window FrontWindow() Layer strongncell1"
    
            msg = " shade window FrontWindow() strongncell1 With rxlev_f_2 "
            msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 120: 66 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,66: 60 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,60: 54 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,54: 48 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,48: 42 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,42: 36 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,36: 30 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,30: 24 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,24: 18 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,18: 12 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,12: 6 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,6: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
            mapinfo.do msg

                 If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                 End If
                 'If Check1.Value = 1 And Check2.Value = 0 Then
                    'msg = " Title " + Chr(34) + "“该切不切”分析 (GSM) " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：相邻小区场强>主小区场强 且 Tx_Power最大   标注：场强大于主小区的邻小区数" + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off"
                    msg = " Title " + Chr(34) + "功率预算切换统计 (GSM) " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：场强大于主小区的邻小区数" + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off"
                    mapinfo.do "set legend window FrontWindow()  Layer prev " & msg

                  mapinfo.do "set map redraw off"
                  mapinfo.do "Set Map Layer strongncell1 Label Visibility Font (""Arial"",257,8,8388736,16777215) With ncell_num Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
                  mapinfo.do "set map redraw on"
        mapinfo.do "close table mytemp"
    End If
next1:
    If Check2.Value = 1 Then
        'mapinfo.do "select * from " & tblname & " where tx_power<2 and (bcch_serv>123 and (bsic_n1<>99 and bcch_n1<124 and (rxlev_n1-rxlev_s>" & Text1(3).Text & ") or bsic_n1<>99 and bcch_n1>123 and (rxlev_n1-rxlev_s>" & Text1(2).Text & ") or bsic_n2<>99 and bcch_n2<124 and (rxlev_n2-rxlev_s>" & Text1(3).Text & ") or bsic_n2<>99 and bcch_n2>123 and (rxlev_n2-rxlev_s>" & Text1(2).Text & ") or bsic_n3<>99 and bcch_n3<124 and (rxlev_n3-rxlev_s>" & Text1(3).Text & ") or bsic_n3 <>99 and bcch_n3>123 and (rxlev_n3-rxlev_s>" & Text1(2).Text & ") or bsic_n4 <>99 and bcch_n4<124 and (rxlev_n4-rxlev_s>" & Text1(3).Text & ") or bsic_n4 <>99 and bcch_n4>123 and (rxlev_n4-rxlev_s>" & Text1(2).Text & ") or bsic_n5 <>99 and bcch_n5<124 and (rxlev_n5-rxlev_s>" & Text1(3).Text & ") or bsic_n5 <>99 and bcch_n5>123 and (rxlev_n5-rxlev_s>" & Text1(2).Text & ") or bsic_n6 <>99 and bcch_n6<124 and (rxlev_n6-rxlev_s>" & Text1(3).Text & ") or bsic_n6<>99 and bcch_n6>123 and (rxlev_n6-rxlev_s>" & Text1(2).Text & ")) into MYTEMP"
        mapinfo.do "select * from " & tblname & " where tx_power<>"""" and val(tx_power)<2 and bcch_serv>123 into mytemp"
        mapinfo.do "select * from mytemp where bcch_n1<124 and (rxlev_n1-rxlev_s>" & Text1(3).Text & ") or bcch_n1>123 and (rxlev_n1-rxlev_s>" & Text1(2).Text & ") or bcch_n2<124 and (rxlev_n2-rxlev_s>" & Text1(3).Text & ") or bcch_n2>123 and (rxlev_n2-rxlev_s>" & Text1(2).Text & ") or bcch_n3<124 and (rxlev_n3-rxlev_s>" & Text1(3).Text & ") or bcch_n3>123 and (rxlev_n3-rxlev_s>" & Text1(2).Text & ") or bcch_n4<124 and (rxlev_n4-rxlev_s>" & Text1(3).Text & ") or bcch_n4>123 and (rxlev_n4-rxlev_s>" & Text1(2).Text & ") or bcch_n5<124 and (rxlev_n5-rxlev_s>" & Text1(3).Text & ") or bcch_n5>123 and (rxlev_n5-rxlev_s>" & Text1(2).Text & ") or bcch_n6<124 and (rxlev_n6-rxlev_s>" & Text1(3).Text & ") or bcch_n6>123 and (rxlev_n6-rxlev_s>" & Text1(2).Text & ") into MYTEMP"
    For i = 1 To mapinfo.eval("NumTables()")
        If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "STRONGNCELL_2" Then
            mapinfo.do "close table strongncell_2"
            Exit For
        End If
    Next
        
        mapinfo.do "commit table mytemp as " + Chr(34) + Gsm_Path + "\user\strongncell_2.tab" + Chr(34)
        mapinfo.do "open table " + Chr(34) + Gsm_Path + "\user\strongncell_2.tab" + Chr(34)
        StrongNcellRow = Val(mapinfo.eval("tableinfo(strongncell_2,8)"))
        If StrongNcellRow = 0 Then
            NonGSM = True
            mapinfo.do "close table strongncell_2"
            GoTo Next2
        End If
        mapinfo.do "fetch first from strongncell_2"
        For i = 1 To StrongNcellRow
            mytemp1 = mapinfo.eval("strongncell_2.rxlev_s")
            mytemp2 = False
            mytemp3 = 0
            For j = 1 To 6
                If ((mapinfo.eval("strongncell_2.rxlev_n" & Format(j)) - mytemp1 > Val(Text1(3))) And mapinfo.eval("strongncell_2.bsic_n" & Format(j)) <> 99 And mapinfo.eval("strongncell_2.bcch_n" & Format(j)) < 124) Or ((mapinfo.eval("strongncell_2.rxlev_n" & Format(j)) - mytemp1 > Val(Text1(2))) And mapinfo.eval("strongncell_2.bsic_n" & Format(j)) <> 99 And mapinfo.eval("strongncell_2.bcch_n" & Format(j)) > 123) Then
                   mytemp3 = mytemp3 + 1
                   If Not mytemp2 Then
                      mapinfo.do "UPDATE strongncell_2 set rxlev_f_2= " & Format(mapinfo.eval("strongncell_2.rxlev_n" & Format(j)) - mytemp1) & " WHERE ROWID=" & i
                      mytemp2 = True
                   End If
                End If
            Next
        
            If mytemp3 = 0 Then
               mapinfo.do "UPDATE strongncell_2 set rxlev_f_2= 0 WHERE ROWID=" & i
            Else
               mapinfo.do "UPDATE strongncell_2 set ncell_num= " & Format(mytemp3) & " WHERE ROWID=" & i
            End If
            mapinfo.do "fetch next from strongncell_2"
        Next
            mapinfo.do "Create Map For strongncell_2 CoordSys Earth Projection 1, 0"
            mapinfo.do "Set Style Symbol MakeSymbol(33,0,2)" '
            mapinfo.do "update strongncell_2 set Obj= CreatePoint(Lon, Lat)"
            mapinfo.do "select * from strongncell_2 where rxlev_f_2>0 into strongncell2"
            
            mapinfo.do "Add Map window FrontWindow() Layer strongncell2"
    
            msg = " shade window FrontWindow() strongncell2 With rxlev_f_2 "
            msg = msg + " ignore 0 ranges apply all use all Symbol (41,16711680,8,""MapInfo Cartographic"",0,0) 120: 66 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,66: 60 Symbol (41,7585792,8,""MapInfo Cartographic"",0,0) ,60: 54 Symbol (41,8388736,8,""MapInfo Cartographic"",0,0) ,54: 48 Symbol (41,16750640,8,""MapInfo Cartographic"",0,0) ,48: 42 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,42: 36 Symbol (41,8421376,8,""MapInfo Cartographic"",0,0) ,36: 30 Symbol (41,8432639,8,""MapInfo Cartographic"",0,0) ,30: 24 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,24: 18 Symbol (41,9584,8,""MapInfo Cartographic"",0,0) ,18: 12 Symbol (41,16744576,8,""MapInfo Cartographic"",0,0) ,12: 6 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,6: 0 Symbol (41,16711680,8,""MapInfo Cartographic"",0,0)"
            mapinfo.do msg

                 If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                 End If
                 'If Check1.Value = 1 And Check2.Value = 0 Then
                    'msg = " Title " + Chr(34) + "“该切不切”分析 (DCS) " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：相邻小区场强>主小区场强 且 Tx_Power最大   标注：场强大于主小区的邻小区数" + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off"
                    msg = " Title " + Chr(34) + "功率预算切换统计 (DCS) " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：场强大于主小区的邻小区数" + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off"
                    mapinfo.do "set legend window FrontWindow()  Layer prev " & msg

                  mapinfo.do "set map redraw off"
                  mapinfo.do "Set Map Layer strongncell2 Label Visibility Font (""Arial"",257,8,8388736,16777215) With ncell_num Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
                  mapinfo.do "set map redraw on"
                  mapinfo.do "close table mytemp"
    End If
Next2:
    If NonGSM And NonDCS Then
       'MsgBox "该路段不存在该切不切问题", 64, "提示"
       'MsgBox "该路段不能进行功率预算切换统计", 64, "提示"
       MsgBox "该路段不能进行功率预算切换统计", 64, "提示"
    ElseIf NonGSM Then
       MsgBox "该路段的GSM网不能进行功率预算切换统计", 64, "提示"
    ElseIf NonDCS Then
       MsgBox "该路段的DCS网不能进行功率预算切换统计", 64, "提示"
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

