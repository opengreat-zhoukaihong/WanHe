VERSION 5.00
Begin VB.Form frmMark 
   Caption         =   "采集事件标注"
   ClientHeight    =   2430
   ClientLeft      =   3780
   ClientTop       =   2655
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
   Icon            =   "frmMark.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   3780
   Begin VB.CommandButton SSCommand3 
      Caption         =   "&O 确定"
      Height          =   320
      Left            =   675
      TabIndex        =   0
      Top             =   2040
      Width           =   1080
   End
   Begin VB.CommandButton SSCommand2 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   1905
      TabIndex        =   1
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "事件选择"
      Height          =   1785
      Left            =   180
      TabIndex        =   8
      Top             =   60
      Width           =   3435
      Begin VB.CheckBox Check1 
         Caption         =   "Start Call"
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   465
         Value           =   1  'Checked
         Width           =   1260
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Blocked Call"
         Height          =   240
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   870
         Width           =   1380
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Dropped Call"
         Height          =   240
         Index           =   2
         Left            =   270
         TabIndex        =   6
         Top             =   1275
         Width           =   1380
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Good Call"
         Height          =   240
         Index           =   4
         Left            =   1980
         TabIndex        =   3
         Top             =   465
         Width           =   1140
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Noisy Call"
         Height          =   240
         Index           =   5
         Left            =   1980
         TabIndex        =   5
         Top             =   870
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No Service"
         Height          =   240
         Index           =   6
         Left            =   1980
         TabIndex        =   7
         Top             =   1275
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SSCommand2_Click()
    
    On Error Resume Next
    Unload Me
End Sub

Private Sub SSCommand3_Click()
    
    On Error Resume Next
    Me.Hide
    If Check1(0).Value = 1 Then    'Start Call
       mapinfo.do "select * from " & tblname & " where (mark= " + Chr(34) + "Start Call" + Chr(34) + ")  into StartC"
       mapinfo.do "Add Map window FrontWindow() Layer  StartC"
       mapinfo.do "shade window FrontWindow() StartC with mark values  " + Chr(34) + "Start Call" + Chr(34) + " Symbol (""start.bmp"",16776960,24,0)"
       If legendid = 0 Then
          mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
          mapinfo.do "Create Legend From Window  Frontwindow()"
          legendid = mapinfo.eval("windowinfo(1009,12)")
       End If
       msg = " Title " + Chr(34) + "Start Call 观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
       mapinfo.do "set legend window FrontWindow()  Layer prev " & msg
    End If
    If Check1(4).Value = 1 Then    'Good Call
       mapinfo.do "select * from " & tblname & " where (mark= " + Chr(34) + "Good Call" + Chr(34) + ")  into GoodC"
       mapinfo.do "Add Map window FrontWindow() Layer  GoodC"
       mapinfo.do "shade window FrontWindow() GoodC with mark values  " + Chr(34) + "Good Call" + Chr(34) + " Symbol (""good.bmp"",16776960,24,0)"
       If legendid = 0 Then
          mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
          mapinfo.do "Create Legend From Window  Frontwindow()"
          legendid = mapinfo.eval("windowinfo(1009,12)")
       End If
       msg = " Title " + Chr(34) + "Good Call 观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
       mapinfo.do "set legend window FrontWindow()  Layer prev " & msg
    End If
    
    If Check1(1).Value = 1 Then    'Blocked Call
       mapinfo.do "select * from " & tblname & " where (mark= " + Chr(34) + "Blocked Call" + Chr(34) + ")  into BlockedC"
       mapinfo.do "Add Map window FrontWindow() Layer  BlockedC"
       mapinfo.do "shade window FrontWindow() BlockedC with mark values  " + Chr(34) + "Blocked Call" + Chr(34) + " Symbol (""Blocked.bmp"",16776960,24,0)"
       If legendid = 0 Then
          mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
          mapinfo.do "Create Legend From Window  Frontwindow()"
          legendid = mapinfo.eval("windowinfo(1009,12)")
       End If
       msg = " Title " + Chr(34) + "Blocked Call 观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
       mapinfo.do "set legend window FrontWindow()  Layer prev " & msg
    End If
    If Check1(5).Value = 1 Then    'Noisy Call
       mapinfo.do "select * from " & tblname & " where (mark= " + Chr(34) + "Noisy Call" + Chr(34) + ")  into NoisyC"
       mapinfo.do "Add Map window FrontWindow() Layer  NoisyC"
       mapinfo.do "shade window FrontWindow() NoisyC with mark values  " + Chr(34) + "Noisy Call" + Chr(34) + " Symbol (""Noisy.bmp"",16776960,24,0)"
       If legendid = 0 Then
          mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
          mapinfo.do "Create Legend From Window  Frontwindow()"
          legendid = mapinfo.eval("windowinfo(1009,12)")
       End If
       msg = " Title " + Chr(34) + "Noisy Call 观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
       mapinfo.do "set legend window FrontWindow()  Layer prev " & msg
    End If
    If Check1(2).Value = 1 Then    'Dropped Call
       mapinfo.do "select * from " & tblname & " where (mark= " + Chr(34) + "Dropped Call" + Chr(34) + ")  into DroppedC"
       mapinfo.do "Add Map window FrontWindow() Layer  DroppedC"
       mapinfo.do "shade window FrontWindow() DroppedC with mark values  " + Chr(34) + "Dropped Call" + Chr(34) + " Symbol (""Dropped.bmp"",16776960,24,0)"
       If legendid = 0 Then
          mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
          mapinfo.do "Create Legend From Window  Frontwindow()"
          legendid = mapinfo.eval("windowinfo(1009,12)")
       End If
       msg = " Title " + Chr(34) + "Dropped Call 观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
       mapinfo.do "set legend window FrontWindow()  Layer prev " & msg
    End If
    If Check1(6).Value = 1 Then    'No Service
       mapinfo.do "select * from " & tblname & " where (mark= " + Chr(34) + "No Service" + Chr(34) + ")  into NoService"
       mapinfo.do "Add Map window FrontWindow() Layer  NoService"
       mapinfo.do "shade window FrontWindow() NoService with mark values  " + Chr(34) + "No Service" + Chr(34) + " Symbol (""NoService.bmp"",16776960,24,0)"
       If legendid = 0 Then
          mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
          mapinfo.do "Create Legend From Window  Frontwindow()"
          legendid = mapinfo.eval("windowinfo(1009,12)")
       End If
       msg = " Title " + Chr(34) + "No Service 观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
       mapinfo.do "set legend window FrontWindow()  Layer prev " & msg
    End If

End Sub
