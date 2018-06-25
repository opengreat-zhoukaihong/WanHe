VERSION 5.00
Begin VB.Form frmDurbSelBcch 
   Caption         =   "选频覆盖统计"
   ClientHeight    =   2730
   ClientLeft      =   6045
   ClientTop       =   870
   ClientWidth     =   3780
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDurbSelBcch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3780
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   330
      Left            =   1950
      TabIndex        =   2
      Top             =   2325
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   345
      Left            =   780
      TabIndex        =   1
      Top             =   2325
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   285
      TabIndex        =   0
      Top             =   135
      Width           =   3210
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1725
         TabIndex        =   7
         Top             =   420
         Width           =   570
      End
      Begin VB.CheckBox Check1 
         Caption         =   "标注BSIC"
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   5
         Top             =   1665
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CheckBox Check1 
         Caption         =   "邻小区"
         Height          =   315
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   1275
         Width           =   885
      End
      Begin VB.CheckBox Check1 
         Caption         =   "主小区"
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   885
         Value           =   1  'Checked
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "所选ARFCN："
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   705
         TabIndex        =   6
         Top             =   480
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmDurbSelBcch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim MyRows As Integer
    Dim i As Integer, j As Integer
    Dim ServingBsic() As String, NeighborBsic() As String
    Dim MyMsgs As String
    Dim OpenTableNum As Integer
    Dim Non_match As Boolean
    Dim Layers As Integer
    
    On Error Resume Next
    If Trim(Text1.Text) = "" Then
        Exit Sub
    End If
    Me.Hide
    
    If Check1(0).Value = 0 Then
       GoTo NeighborSearch
    End If
    mapinfo.do "select * from " & tblname & " where bcch_serv = " & Text1.Text & " into ServingBcch"
    If mapinfo.eval("tableinfo(ServingBcch,8)") = 0 Then
        Non_match = True
        mapinfo.do "close table ServingBcch"
        GoTo NeighborSearch
    End If
    
    mapinfo.do "Add Map window Frontwindow() Layer ServingBcch"
    MyMsgs = "shade window FrontWindow() ServingBcch With rxlev_s "
             'Symbol (169,65280,10,"Monotype Sorts",0,0)   'Serving
             'Symbol (64,255,10,"Monotype Sorts",0,0)      'Neighbor
             'Symbol (224,16711680,10,"Wingdings 2",0,0)   'Neighbor
    'For i = 0 To MyRows - 1
        ''MyMsgs = MyMsgs + Chr(34) + ServingBsic(i) + Chr(34) + " Symbol (169," + Format(MyRndColor(i)) + ",12,""Monotype Sorts"",0,0),"
         'MyMsgs = MyMsgs + Chr(34) + ServingBsic(i) + Chr(34) + " Symbol (41," + Format(MyRndColor(i)) + ",7,""MapInfo Cartographic"",0,0),"
    'Next
    'MyMsgs = Left(MyMsgs, Len(MyMsgs) - 1)
    
    If Legend_Tog = 0 Then
        MyMsgs = MyMsgs + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
    Else
        MyMsgs = MyMsgs + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 120: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
    End If
    mapinfo.do MyMsgs
                 
                 If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                 End If
                 
    If Legend_Tog = 0 Then
        MyMsgs = " Title " + Chr(34) + "选频覆盖统计 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "BCCH=" & Text1.Text & "的覆盖分布  标注：BSIC(蓝色)" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
    Else
        MyMsgs = " Title " + Chr(34) + "选频覆盖统计 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "BCCH=" & Text1.Text & "的覆盖分布  标注：BSIC(蓝色)" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
    End If
                 
               mapinfo.do "set legend window FrontWindow()  Layer prev " & MyMsgs
               mapinfo.do "set map redraw off"
               mapinfo.do "Set Map Layer ServingBcch Label Visibility Font (""Arial"",257,8,255,16777215) With bsic_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
               mapinfo.do "set map redraw on"
               
       'mapinfo.do "select * from ServingBcch into DupLabel"
       'mapinfo.do "Add Map window FrontWindow() Layer DupLabel"
       'Layers = mapinfo.eval("mapperinfo(frontwindow(),9)")
       'MyMsgs = "set map order "
       'For i = 2 To Layers
       '    MyMsgs = MyMsgs + Format(i) + ","
       'Next
       'MyMsgs = MyMsgs + "1"
       'mapinfo.do MyMsgs
       '        mapinfo.do "set map redraw off"
       '        mapinfo.do "Set Map Layer DupLabel Label Visibility Font (""Arial"",257,8,255,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position below Auto On Offset 2"
       '        mapinfo.do "set map redraw on"
               
NeighborSearch:
        If Check1(1).Value = 0 Then
            If Non_match And Check1(0).Value = 1 Then
                MsgBox "不存在主小区载频为" & Text1.Text & "的覆盖", 64, "提示"
            End If
            Unload Me
            Exit Sub
        End If
          
       OpenTableNum = mapinfo.eval("NumTables()")
       For i = 1 To OpenTableNum
           If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "NEIGHBORBCCH" Then
              mapinfo.do "close table NeighborBcch"
              Exit For
           End If
       Next

    mapinfo.do "select * from " & tblname & " where Bcch_N1= " & Text1.Text & " or Bcch_N2= " & Text1.Text & " or Bcch_N3= " & Text1.Text & " or Bcch_N4= " & Text1.Text & " or Bcch_N5= " & Text1.Text & " or Bcch_N6= " & Text1.Text & " into NeighborBcch"
        
    If mapinfo.eval("tableinfo(NeighborBcch,8)") = 0 Then
        If Non_match Then
            If Check1(0).Value = 1 Then
                MsgBox "不存在主、邻小区载频为" & Text1.Text & "的覆盖", 64, "提示"
            Else
                MsgBox "不存在邻小区载频为" & Text1.Text & "的覆盖", 64, "提示"
            End If
        Else
            MsgBox "不存在邻小区载频为" & Text1.Text & "的覆盖", 64, "提示"
        End If
        mapinfo.do "close table NeighborBcch"
        Unload Me
        Exit Sub
    End If
    MyRows = mapinfo.eval("tableinfo(NeighborBcch,8)")
    If MyRows > 0 Then
        mapinfo.do "commit table NeighborBcch as " + Chr(34) + Gsm_Path + "\User\NeighborBcch.tab" + Chr(34)
        mapinfo.do "close table NeighborBcch"
        mapinfo.do "open table " + Chr(34) + Gsm_Path + "\User\NeighborBcch.tab" + Chr(34)
        'mapinfo.do "Alter Table ""NeighborBcch"" ( add NcellBcch Decimal(3,0),NcellBsic Decimal(3,0)) Interactive"
        mapinfo.do "Alter Table ""NeighborBcch"" ( add NcellBcch Decimal(3,0),NcellBsic Char(3),NcellRxlev Decimal(3,0)) Interactive"
        mapinfo.do "fetch first from NeighborBcch"
        For i = 1 To MyRows
            For j = 1 To 6
                If Abs(mapinfo.eval("NeighborBcch.bcch_n" & Format(j)) - Val(Text1.Text)) <= 1 Then
                    If mapinfo.eval("NeighborBcch.bsic_n" & Format(j)) = 99 Then
                        mapinfo.do "UPDATE NeighborBcch set NcellBcch = NeighborBcch.bcch_n" & Format(j) & ",NcellBsic = ""**"",NcellRxlev=NeighborBcch.rxlev_n" & Format(j) & " where rowid = " & Format(i)
                    Else
                        mapinfo.do "UPDATE NeighborBcch set NcellBcch = NeighborBcch.bcch_n" & Format(j) & ",NcellBsic = Str$(NeighborBcch.bsic_n" & Format(j) & "),NcellRxlev=NeighborBcch.rxlev_n" & Format(j) & " where rowid = " & Format(i)
                    End If
                    Exit For
                End If
            Next
            mapinfo.do "fetch next from NeighborBcch"
        Next
        mapinfo.do "commit table NeighborBcch"
        mapinfo.do "Add Map window Frontwindow() Layer NeighborBcch"
        MyMsgs = "shade window FrontWindow() NeighborBcch With NcellRxlev "
    If Legend_Tog = 0 Then
        MyMsgs = MyMsgs + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
    Else
        MyMsgs = MyMsgs + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 120: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
    End If
    mapinfo.do MyMsgs
                                  
    If Legend_Tog = 0 Then
        MyMsgs = " Title " + Chr(34) + "选频覆盖统计 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "NCELL ARFCN=" & Text1.Text & "的覆盖分布  标注：BSIC(粉红色)" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
    Else
        MyMsgs = " Title " + Chr(34) + "选频覆盖统计 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "NCELL ARFCN=" & Text1.Text & "的覆盖分布  标注：BSIC(粉红色)" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
    End If
                 
                 'MyMsgs = " Title " + Chr(34) + "干扰选频分析 " + mySelTbl + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "同频主小区   所选载频：" + text1.Text + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off"
                 mapinfo.do "set legend window FrontWindow()  Layer prev " & MyMsgs
                 
               mapinfo.do "set map redraw off"
               mapinfo.do "Set Map Layer NeighborBcch Label Visibility Font (""Arial"",257,8,16711935,16777215) With ncellbsic Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
               mapinfo.do "set map redraw on"
    
    Else
        mapinfo.do "close table NeighborBcch"
    End If
    mapinfo.do "close table selection"
    mapinfo.runmenucommand 610
    If Non_match Then
        MsgBox "不存在主小区载频为" & Text1.Text & "的覆盖", 64, "提示"
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

