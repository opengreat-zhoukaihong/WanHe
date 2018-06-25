VERSION 5.00
Begin VB.Form frmGDNcell 
   Caption         =   "有效邻小区分布"
   ClientHeight    =   2535
   ClientLeft      =   4185
   ClientTop       =   3075
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
   Icon            =   "frmGDNcell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3525
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   375
      TabIndex        =   2
      Top             =   180
      Width           =   2790
      Begin VB.CheckBox Check1 
         Caption         =   "DCS网有效邻小区"
         Height          =   285
         Index           =   1
         Left            =   690
         TabIndex        =   4
         Top             =   990
         Width           =   1665
      End
      Begin VB.CheckBox Check1 
         Caption         =   "GSM网有效邻小区"
         Height          =   285
         Index           =   0
         Left            =   690
         TabIndex        =   3
         Top             =   510
         Value           =   1  'Checked
         Width           =   1755
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      DragIcon        =   "frmGDNcell.frx":000C
      Height          =   320
      Left            =   1845
      TabIndex        =   1
      Top             =   2100
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      DragIcon        =   "frmGDNcell.frx":015E
      Height          =   320
      Left            =   660
      TabIndex        =   0
      Top             =   2115
      Width           =   1080
   End
End
Attribute VB_Name = "frmGDNcell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    On Error Resume Next
    If Menu_Flag = 991109 Then
        If Check1(0).Value = 1 And Check1(1).Value = 1 Then
            HOParaFlag = 3
            mapinfo.do "select * from " & tblname & " where Left$(mark1,3)=""HSC"" or Left$(mark1,3)=""HFC"" into HOParameter"
        ElseIf Check1(0).Value = 1 Then
            HOParaFlag = 1
            mapinfo.do "select * from " & tblname & " where Left$(mark1,3)=""HFC"" into HOParameter"
        ElseIf Check1(1).Value = 1 Then
            HOParaFlag = 2
            mapinfo.do "select * from " & tblname & " where Left$(mark1,3)=""HSC"" into HOParameter"
        End If
        If mapinfo.eval("tableinfo(HOParameter,8)") = 0 Then
            MsgBox "不存在切换过程或该数据是旧数据，无法显示切换前后参数", 64, "提示"
            HOParaFlag = 0
        End If
        Unload Me
        Exit Sub
    End If
    Me.Hide
    If Check1(0).Value = 1 Then
        mapinfo.do " shade window FrontWindow() " + tblname + " With ncell_num values 0 Symbol (41,16711680,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,13671424,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,16776960,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) "
        If legendid = 0 Then
            mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
            mapinfo.do "Create Legend From Window  Frontwindow()"
            legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        mapinfo.do "set legend window FrontWindow()  Layer prev Title " + Chr(34) + "有效邻小区分布" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "GSM网   标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.do "set map redraw off"
        mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
        mapinfo.do "set map redraw on"
    End If
    If Check1(1).Value = 1 Then
        mapinfo.do "select * from " & tblname & " where bsic_same1>0 into mytemp"
        If mapinfo.eval("tableinfo(mytemp,8)") = 0 Then
            MsgBox "不存在DCS网有效邻小区", 64, "提示"
            mapinfo.do "close table mytemp"
            Unload Me
            Exit Sub
        End If
        mapinfo.do "close table mytemp"
        mapinfo.do " shade window FrontWindow() " + tblname + " With bsic_same1 values 0 Symbol (41,16711680,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,13671424,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,16776960,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) "
        If legendid = 0 Then
            mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
            mapinfo.do "Create Legend From Window  Frontwindow()"
            legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        mapinfo.do "set legend window FrontWindow()  Layer prev Title " + Chr(34) + "有效邻小区分布" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "DCS网   标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.do "set map redraw off"
        mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
        mapinfo.do "set map redraw on"
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If Menu_Flag = 991109 Then
        Me.Caption = "切换前后参数显示"
        Check1(0).Caption = "切换失败"
        Check1(1).Caption = "切换成功"
    End If
End Sub
