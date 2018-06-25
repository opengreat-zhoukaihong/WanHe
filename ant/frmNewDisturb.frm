VERSION 5.00
Begin VB.Form frmNewDisturb 
   Caption         =   "上行干扰查找条件"
   ClientHeight    =   2445
   ClientLeft      =   3645
   ClientTop       =   4110
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewDisturb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3720
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   165
      TabIndex        =   2
      Top             =   165
      Width           =   3405
      Begin VB.TextBox RxLevValue 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1920
         TabIndex        =   3
         Text            =   "70"
         Top             =   510
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "如果RxQual>3，可能存在下行干扰（BS->MS）"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   8
         Top             =   1275
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "如果RxQual<4，可能存在上行干扰（MS->BS）"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   1
         Left            =   315
         TabIndex        =   7
         Top             =   1125
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tx_Power最大(GSM 2W;DCS 0.6W)"
         Height          =   180
         Index           =   0
         Left            =   375
         TabIndex        =   6
         Top             =   1035
         Width           =   2610
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "下行场强RxLev >"
         Height          =   180
         Index           =   6
         Left            =   495
         TabIndex        =   5
         Top             =   555
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "-dBm"
         Height          =   180
         Left            =   2475
         TabIndex        =   4
         Top             =   555
         Width           =   360
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      DragIcon        =   "frmNewDisturb.frx":000C
      Height          =   320
      Left            =   1980
      TabIndex        =   1
      Top             =   2040
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      DragIcon        =   "frmNewDisturb.frx":015E
      Height          =   320
      Left            =   735
      TabIndex        =   0
      Top             =   2040
      Width           =   1080
   End
End
Attribute VB_Name = "frmNewDisturb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Integer, Layers As Integer
    Dim Mymsg As String
    Dim CellLayer As Integer
    
    On Error Resume Next
    Me.Hide
    'mapinfo.DO "Select * from " & tblname & " where tx_power<>"""" and bcch_serv >124 and val(tx_power)<2 and rxlev_s > " & Format(110 - Val(RxLevValue.Text)) & " or tx_power<>"""" and bcch_serv <125 and val(tx_power)<6 and rxlev_s > " & Format(110 - Val(RxLevValue.Text)) & " into Disturb"
    If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
       'mapinfo.do "Select * from " & tblname & " where tx_power<>"""" and val(rxqual_s) <4 and (bcch_serv >124 and val(tx_power)<2 and rxlev_s > " & Format(110 - Val(RxLevValue.Text)) & " or tx_power<>"""" and bcch_serv <125 and val(tx_power)<6 and rxlev_s > " & Format(110 - Val(RxLevValue.Text)) & ") into Disturb"
       mapinfo.do "Select * from " & tblname & " where tx_power<>"""" and (bcch_serv >124 and val(tx_power)<2 and rxlev_s > " & Format(110 - Val(RxLevValue.Text)) & " or tx_power<>"""" and bcch_serv <125 and val(tx_power)<6 and rxlev_s > " & Format(110 - Val(RxLevValue.Text)) & ") into Disturb"
    Else
       'mapinfo.do "Select * from " & tblname & " where tx_power<>"""" and rxqual_s <4 and (bcch_serv >124 and val(tx_power)<2 and rxlev_s > " & Format(110 - Val(RxLevValue.Text)) & " or tx_power<>"""" and bcch_serv <125 and val(tx_power)<6 and rxlev_s > " & Format(110 - Val(RxLevValue.Text)) & ") into Disturb"
       mapinfo.do "Select * from " & tblname & " where tx_power<>"""" and and (bcch_serv >124 and val(tx_power)<2 and rxlev_s > " & Format(110 - Val(RxLevValue.Text)) & " or tx_power<>"""" and bcch_serv <125 and val(tx_power)<6 and rxlev_s > " & Format(110 - Val(RxLevValue.Text)) & ") into Disturb"
    End If
    If Val(mapinfo.eval("tableinfo(Disturb,8)")) = 0 Then
       MsgBox "该路段不存在上行干扰", 64, "提示"
       mapinfo.do "close table Disturb"
       Unload Me
       Exit Sub
    End If
    mapinfo.do "Add Map Auto Layer Disturb"
    
    msg = " shade window FrontWindow() Disturb With val(rxqual_s) "
    msg = msg + " ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 0: 7 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) " ' , 3: 7 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) "
    mapinfo.do msg
    
                  If legendid = 0 Then     'win95
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"     'win95
                      mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
                      legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
                  End If     'win95
    
                   'msg = " Title " + Chr(34) + "干扰分析 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "标注 蓝色：RxQual_s 粉红色：FER" + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""上行干扰"" display on ,""下行干扰"" display on "
    
    msg = " Title " + Chr(34) + "干扰分析 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "标注 蓝色：RxQual_s 粉红色：FER" + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""上行干扰"" display on "
    'msg = " Title " + Chr(34) + "干扰分析 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev>" & RxLevValue.Text & "-dBm且Tx_Power最大   标注 蓝色：RxQual_s 粉红色：FER" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""上行干扰"" display on "
    
    mapinfo.do "set legend window FrontWindow() Layer prev " & msg

       mapinfo.do "select * from Disturb into MyDuplicate"
       mapinfo.do "Add Map window FrontWindow() Layer MyDuplicate"
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer ""Disturb"" Label Visibility Font (""Arial"",257,8,255,16777215) With rxqual_s Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
          mapinfo.do "set map redraw on"
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer ""MyDuplicate"" Label Visibility Font (""Arial"",257,8,16711935,16777215) With fer Auto On Overlap Off Duplicates On Position Below Auto On Offset 10"
          mapinfo.do "set map redraw on"
          
       Layers = mapinfo.eval("mapperinfo(frontwindow(),9)")
       Mymsg = "set map order "
       For i = 2 To Layers
           Mymsg = Mymsg + Format(i) + ","
       Next
       Mymsg = Mymsg + "1"
       mapinfo.do Mymsg

    Unload Me

End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub
