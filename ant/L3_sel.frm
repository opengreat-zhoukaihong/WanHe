VERSION 5.00
Begin VB.Form L3_Sel 
   BackColor       =   &H00C0C0C0&
   Caption         =   " 主要信令选择"
   ClientHeight    =   2985
   ClientLeft      =   2805
   ClientTop       =   2085
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "L3_sel.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2985
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "主要信令选择"
      Height          =   2190
      Left            =   180
      TabIndex        =   2
      Top             =   165
      Width           =   4290
      Begin VB.CheckBox Check1 
         Caption         =   "ASSIGNMENT COMPLETE"
         Height          =   240
         Index           =   7
         Left            =   2145
         TabIndex        =   10
         Top             =   1680
         Width           =   2010
      End
      Begin VB.CheckBox Check1 
         Caption         =   "IDEL REPORT"
         Height          =   240
         Index           =   6
         Left            =   2145
         TabIndex        =   9
         Top             =   1275
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CHANNEL RELEASE"
         Height          =   240
         Index           =   5
         Left            =   2145
         TabIndex        =   8
         Top             =   870
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CONNECT"
         Height          =   240
         Index           =   4
         Left            =   2145
         TabIndex        =   7
         Top             =   465
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "LOCATION UPDATE"
         Height          =   240
         Index           =   3
         Left            =   270
         TabIndex        =   6
         Top             =   1680
         Width           =   1680
      End
      Begin VB.CheckBox Check1 
         Caption         =   "SETUP"
         Height          =   240
         Index           =   2
         Left            =   270
         TabIndex        =   5
         Top             =   1275
         Width           =   795
      End
      Begin VB.CheckBox Check1 
         Caption         =   "RELEASE"
         Height          =   240
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   870
         Width           =   960
      End
      Begin VB.CheckBox Check1 
         Caption         =   "HANDOVER"
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   465
         Value           =   1  'Checked
         Width           =   1050
      End
   End
   Begin VB.CommandButton SSCommand2 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   2430
      TabIndex        =   1
      Top             =   2550
      Width           =   1080
   End
   Begin VB.CommandButton SSCommand3 
      Caption         =   "&O 确定"
      Height          =   320
      Left            =   1200
      TabIndex        =   0
      Top             =   2550
      Width           =   1080
   End
End
Attribute VB_Name = "L3_Sel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim select1(1 To 8) As String * 1

Private Sub Form_Load()

    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\ant.cfg"
    Open Gsm_FileName For Binary As #1
    Seek #1, 18
    For j = 1 To 8
       Get #1, 18 + j, select1(j)
       Check1(j - 1).Value = Val(select1(j))
    Next
    Close
End Sub

Private Sub Check1_Click(Index As Integer)

    On Error Resume Next
    If Check1(Index).Value = 1 Then
       select1(Index + 1) = "1"
    Else
       select1(Index + 1) = "0"
    End If
End Sub


Private Sub SSCommand2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub SSCommand3_Click()
    Dim endeof As String * 1
    
    On Error Resume Next
    L3_Sel.Hide
    endeof = Chr$(26)
    Gsm_FileName = Gsm_Path + "\ant.cfg"
    Open Gsm_FileName For Binary As #1
    Seek #1, 18
    For j = 1 To 8
       If Check1(j - 1).Value = 1 Then
         select1(j) = "1"
       Else
         select1(j) = "0"
       End If
       Put #1, 18 + j, select1(j)
    Next
    
    Put #1, , endeof
    Close

      If Check1(0).Value = 1 Then
                  mapinfo.do "select * from " & tblname & " where ( MESSAGE= " + Chr(34) + "Handover complete" + Chr(34) + ") or (MESSAGE= " + Chr(34) + "Handover Failure" + Chr(34) + ") or (MESSAGE= " + Chr(34) + "Handover Command" + Chr(34) + ") into HANDOVER"
                  mapinfo.do "Add Map window FrontWindow() Layer  HANDOVER"

                  msg = "shade window FrontWindow() HANDOVER with MESSAGE values  " + Chr(34) + "Handover complete" + Chr(34) + " Symbol (""hand_c.bmp"",16776960,24,0),"
                  msg = msg + Chr(34) + "Handover Failure" + Chr(34) + " Symbol (""hand_f.bmp"",16776960,24,0),"
                  msg = msg + Chr(34) + "Handover Command" + Chr(34) + " Symbol (""hand_com.bmp"",255,24,0)"
                  mapinfo.do msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                      
                  End If
                  msg = " Title " + Chr(34) + "HANDOVER观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & msg
       End If

      If Check1(1).Value = 1 Then
                  mapinfo.do "select * from " & tblname & " where (MESSAGE= " + Chr(34) + "Release Complete" + Chr(34) + ") or (MESSAGE= " + Chr(34) + "Release Fail" + Chr(34) + ") or (MESSAGE= " + Chr(34) + "Release" + Chr(34) + ") into Release"
                  mapinfo.do "Add Map window FrontWindow() Layer  Release"
                  msg = "shade window FrontWindow() Release with MESSAGE values  " + Chr(34) + "Release Complete" + Chr(34) + " Symbol (""release.bmp"",65535,22,0), "
                  msg = msg + Chr(34) + "Release" + Chr(34) + " Symbol (""release.bmp"",19711765,24,0),"
                  msg = msg + Chr(34) + "Release Fail" + Chr(34) + " Symbol (""rele_f.bmp"",19711765,24,0)"
                  mapinfo.do msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  msg = " Title " + Chr(34) + "RELEASE观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & msg

       End If

      If Check1(2).Value = 1 Then
                  mapinfo.do "select * from " & tblname & " where MESSAGE= " + Chr(34) + "SETUP" + Chr(34) + " into SETUP"
                  mapinfo.do "Add Map window FrontWindow() Layer  SETUP"

                  mapinfo.do "shade window FrontWindow()  Setup with MESSAGE values  " + Chr(34) + "SETUP" + Chr(34) + " Symbol (""setup.bmp"",255,22,0) "
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  msg = " Title " + Chr(34) + "SETUP观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & msg

       End If

      If Check1(3).Value = 1 Then
                  mapinfo.do "select * from " & tblname & " where (MESSAGE= " + Chr(34) + "Location Updating Accept" + Chr(34) + ") or (MESSAGE= " + Chr(34) + "Location Updating Reject" + Chr(34) + ") into Locup"
                  mapinfo.do "Add Map window FrontWindow() Layer  Locup"
                  msg = "shade window FrontWindow()  Locup with MESSAGE values  " + Chr(34) + "Location Updating Accept" + Chr(34) + " Symbol (""loc_acc.bmp"",16711680,24,0), " + Chr(34) + "Location Updating Reject" + Chr(34) + " Symbol (""LOC_F.bmp"",16711680,24,0)"
                  mapinfo.do msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  msg = " Title " + Chr(34) + "LOCATION_UPDATE观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & msg


       End If

      If Check1(4).Value = 1 Then
                  mapinfo.do "select * from " & tblname & " where ( MESSAGE= " + Chr(34) + "CONNECT" + Chr(34) + ") Or  ( MESSAGE= " + Chr(34) + "Connect Fail" + Chr(34) + ")  Or  ( MESSAGE= " + Chr(34) + "disconnect" + Chr(34) + ") into CONNECT"
                  mapinfo.do "Add Map window FrontWindow() Layer  CONNECT"

                  msg = "shade window FrontWindow()  CONNECT with MESSAGE  values  " + Chr(34) + "Connect" + Chr(34) + " Symbol (""connect.bmp"",65280,22,0),"
                  msg = msg + Chr(34) + "Disconnect" + Chr(34) + " Symbol (""discon.bmp"",19711685,24,0)"
                  mapinfo.do msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  On Error Resume Next
                  msg = " Title " + Chr(34) + "CONNECT观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & msg

       End If

      If Check1(5).Value = 1 Then
                  mapinfo.do "select * from " & tblname & " where MESSAGE= " + Chr(34) + "Channel release" + Chr(34) + " into ch_rele"
                  mapinfo.do "Add Map window FrontWindow() Layer  ch_rele"

                  mapinfo.do "shade window FrontWindow() ch_rele with MESSAGE values  " + Chr(34) + "Channel release" + Chr(34) + " Symbol (""chan_rel.bmp"",255,22,0) "
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  msg = " Title " + Chr(34) + " channel release 观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & msg

       End If

      If Check1(6).Value = 1 Then
                  mapinfo.do "select * from " & tblname & " where MESSAGE= " + Chr(34) + "Idle mode report" + Chr(34) + " or MESSAGE= " + Chr(34) + "Idle mode" + Chr(34) + " into Idle"
                  mapinfo.do "Add Map window FrontWindow() Layer  Idle"

                  mapinfo.do "shade window FrontWindow() Idle with MESSAGE values  " + Chr(34) + "Idle mode report" + Chr(34) + " Symbol (""Idle.bmp"",255,22,0) ," + Chr(34) + "Idle mode" + Chr(34) + " Symbol (""Idle.bmp"",255,22,0)"
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  msg = " Title " + Chr(34) + " Idle 观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & msg
       End If

      If Check1(7).Value = 1 Then
                  mapinfo.do "select * from " & tblname & " where MESSAGE= " + Chr(34) + "assignment complete" + Chr(34) + " into ASIG_COM"
                  mapinfo.do "Add Map window FrontWindow() Layer  ASIG_COM"

                  mapinfo.do "shade window FrontWindow() ASIG_COM with MESSAGE values  " + Chr(34) + "assignment complete" + Chr(34) + " Symbol (""imm.bmp"",255,22,0) "
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  msg = " Title " + Chr(34) + " assignment complete 观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & msg
       End If

    Unload Me
End Sub

