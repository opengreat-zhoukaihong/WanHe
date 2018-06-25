VERSION 5.00
Begin VB.Form frmGDCover 
   Caption         =   "G网与D网覆盖显示"
   ClientHeight    =   3060
   ClientLeft      =   4110
   ClientTop       =   1545
   ClientWidth     =   3990
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGDCover.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3990
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   375
      TabIndex        =   2
      Top             =   180
      Width           =   3240
      Begin VB.CheckBox Check1 
         Caption         =   "邻小区覆盖"
         Height          =   405
         Index           =   1
         Left            =   1770
         TabIndex        =   6
         Top             =   1545
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "主小区覆盖"
         Height          =   405
         Index           =   0
         Left            =   315
         TabIndex        =   5
         Top             =   1545
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1800兆GSM网覆盖"
         Height          =   375
         Index           =   1
         Left            =   825
         TabIndex        =   4
         Top             =   855
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "900兆GSM网覆盖"
         Height          =   375
         Index           =   0
         Left            =   825
         TabIndex        =   3
         Top             =   405
         Width           =   1650
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      DragIcon        =   "frmGDCover.frx":000C
      Height          =   320
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      DragIcon        =   "frmGDCover.frx":015E
      Height          =   320
      Left            =   855
      TabIndex        =   0
      Top             =   2640
      Width           =   1080
   End
End
Attribute VB_Name = "frmGDCover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim MyPoint As Long
    
    On Error Resume Next
    Me.Hide
    If Option1(0).Value Then   '900
        If Check1(0) = 1 And Check1(1) = 1 Then
            mapinfo.do "select * from " & tblname & " where bcch_serv<>0 and (bcch_serv<=124 or bcch_n1<=124 or bcch_n2<=124 or bcch_n3<=124 or bcch_n4<=124 or bcch_n5<=124 or bcch_n6<=124) into GSMCover"
        ElseIf Check1(0) = 1 Then
            mapinfo.do "select * from " & tblname & " where bcch_serv<>0 and bcch_serv<=124 into GSMCover"
        Else
            mapinfo.do "select * from " & tblname & " where bcch_n1<=124 or bcch_n2<=124 or bcch_n3<=124 or bcch_n4<=124 or bcch_n5<=124 or bcch_n6<=124 into GSMCover"
        End If
        MyPoint = mapinfo.eval("tableinfo(GSMCover,8)")
        If MyPoint = 0 Then
            MsgBox "该路段不存在900兆GSM网覆盖", 64, "提示"
            mapinfo.do "close table GSMCover"
            Unload Me
            Exit Sub
        End If
        mapinfo.do "Add Map window FrontWindow() Layer GSMCover"
        If Legend_Tog = 0 Then
            mapinfo.do "shade window FrontWindow() GSMCover With RXLEV_s ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
        Else
            mapinfo.do "shade window FrontWindow() GSMCover With RXLEV_s ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 120: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
        End If
        If legendid = 0 Then     'win95
            mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"     'win95
            mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
            legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
        End If     'win95
        If Legend_Tog = 0 Then
            mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "900兆GSM网覆盖显示 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH（蓝色）" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
        Else
            mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "900兆GSM网覆盖显示 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH（蓝色）" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
        End If
        mapinfo.do "set map redraw off"
        mapinfo.do "Set Map Layer ""GSMCover"" Label Visibility Font (""Arial"",257,8,255,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
        mapinfo.do "set map redraw on"
    Else
        If Check1(0) = 1 And Check1(1) = 1 Then
            mapinfo.do "select * from " & tblname & " where bcch_serv>124 or bcch_n1>124 or bcch_n2>124 or bcch_n3>124 or bcch_n4>124 or bcch_n5>124 or bcch_n6>124 into DCSCover"
        ElseIf Check1(0) = 1 Then
            mapinfo.do "select * from " & tblname & " where bcch_serv>124 into DCSCover"
        Else
            mapinfo.do "select * from " & tblname & " where bcch_n1>124 or bcch_n2>124 or bcch_n3>124 or bcch_n4>124 or bcch_n5>124 or bcch_n6>124 into DCSCover"
        End If
        MyPoint = mapinfo.eval("tableinfo(DCSCover,8)")
        If MyPoint = 0 Then
            MsgBox "该路段不存在1800兆GSM网覆盖", 64, "提示"
            mapinfo.do "close table DCSCover"
            Unload Me
            Exit Sub
        End If
        mapinfo.do "Add Map window FrontWindow() Layer DCSCover"
        If Legend_Tog = 0 Then
            mapinfo.do "shade window FrontWindow() DCSCover With RXLEV_s ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
        Else
            mapinfo.do "shade window FrontWindow() DCSCover With RXLEV_s ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 120: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
        End If
        If legendid = 0 Then     'win95
            mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"     'win95
            mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
            legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
        End If     'win95
        If Legend_Tog = 0 Then
            mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "1800兆GSM网覆盖显示 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH（杏黄色）" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
        Else
            mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "1800兆GSM网覆盖显示 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH（杏黄色）" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
        End If
        mapinfo.do "set map redraw off"
        mapinfo.do "Set Map Layer ""DCSCover"" Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
        mapinfo.do "set map redraw on"
    End If
    mapinfo.do "close table selection"
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub
