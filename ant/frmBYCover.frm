VERSION 5.00
Begin VB.Form frmBYCover 
   Caption         =   "本网与异网覆盖显示"
   ClientHeight    =   2535
   ClientLeft      =   4110
   ClientTop       =   1545
   ClientWidth     =   3525
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBYCover.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3525
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   375
      TabIndex        =   2
      Top             =   180
      Width           =   2790
      Begin VB.OptionButton Option1 
         Caption         =   "联通网覆盖"
         Height          =   375
         Index           =   1
         Left            =   765
         TabIndex        =   4
         Top             =   960
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
         Caption         =   "电信网覆盖"
         Height          =   375
         Index           =   0
         Left            =   765
         TabIndex        =   3
         Top             =   510
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      DragIcon        =   "frmBYCover.frx":000C
      Height          =   320
      Left            =   1845
      TabIndex        =   1
      Top             =   2100
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      DragIcon        =   "frmBYCover.frx":015E
      Height          =   320
      Left            =   660
      TabIndex        =   0
      Top             =   2100
      Width           =   1080
   End
End
Attribute VB_Name = "frmBYCover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim MyPoint As Long
    
    On Error Resume Next
    Me.Hide
    If Option1(0).Value Then   '电信
        'mapinfo.do "select * from " & tblname & " where Ltrim$(Rtrim$(Mnc_serv))=""00"" into TelecomCover"
        mapinfo.do "select * from " & tblname & " where Ltrim$(Rtrim$(Mnc_serv))<>""01"" and Ltrim$(Rtrim$(Mnc_serv))<>"""" into TelecomCover"
        MyPoint = mapinfo.eval("tableinfo(TelecomCover,8)")
        If MyPoint = 0 Then
            MsgBox "该路段不存在电信网覆盖", 64, "提示"
            mapinfo.do "close table TelecomCover"
            Unload Me
            Exit Sub
        End If
        mapinfo.do "Add Map window FrontWindow() Layer TelecomCover"
        'If Legend_Tog = 0 Then
            mapinfo.do "shade window FrontWindow() TelecomCover With mnc_serv values ""00"" Symbol (39,255,8,""MapInfo Cartographic"",0,0) default Symbol (39,8388736,8,""MapInfo Cartographic"",0,0)"
        'Else
        '    mapinfo.do "shade window FrontWindow() TelecomCover With RXLEV_s ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 120: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
        'End If
        If legendid = 0 Then     'win95
            mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"     'win95
            mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
            legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
        End If     'win95
        'If Legend_Tog = 0 Then
            mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "电信网覆盖显示 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH（蓝色）" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其他"" display on ,""电信网"" display on"
        'Else
        '    mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "电信网覆盖显示 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH（蓝色）" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
        'End If
        mapinfo.do "set map redraw off"
        mapinfo.do "Set Map Layer ""TelecomCover"" Label Visibility Font (""Arial"",257,8,255,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
        mapinfo.do "set map redraw on"
    Else
        mapinfo.do "select * from " & tblname & " where Ltrim$(Rtrim$(Mnc_serv))<>""00"" and  Ltrim$(Rtrim$(Mnc_serv))<>"""" into UnitecomCover"
        MyPoint = mapinfo.eval("tableinfo(UnitecomCover,8)")
        If MyPoint = 0 Then
            MsgBox "该路段不存在联通网覆盖", 64, "提示"
            mapinfo.do "close table UnitecomCover"
            Unload Me
            Exit Sub
        End If
        mapinfo.do "Add Map window FrontWindow() Layer UnitecomCover"
        mapinfo.do "shade window FrontWindow() UnitecomCover With mnc_serv values ""01"" Symbol (39,56567,8,""MapInfo Cartographic"",0,0) default Symbol (39,8388736,8,""MapInfo Cartographic"",0,0)"
        'If Legend_Tog = 0 Then
        '    mapinfo.do "shade window FrontWindow() UnitecomCover With RXLEV_s ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
        'Else
        '    mapinfo.do "shade window FrontWindow() UnitecomCover With RXLEV_s ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 120: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
        'End If
        If legendid = 0 Then     'win95
            mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"     'win95
            mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
            legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
        End If     'win95
        'If Legend_Tog = 0 Then
        '    mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "联通网覆盖显示 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH（杏黄色）" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
        'Else
        '    mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "联通网覆盖显示 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH（杏黄色）" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
        'End If
        mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "联通网覆盖显示 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH（蓝色）" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其他"" display on ,""联通网"" display on"
        mapinfo.do "set map redraw off"
        mapinfo.do "Set Map Layer ""UnitecomCover"" Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
        mapinfo.do "set map redraw on"
    End If
    mapinfo.do "close table selection"
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub
