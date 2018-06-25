VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form Cope_RxLev 
   BackColor       =   &H00C0C0C0&
   Caption         =   "条件选择"
   ClientHeight    =   2370
   ClientLeft      =   3240
   ClientTop       =   2970
   ClientWidth     =   4050
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Cope_rxl.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1530
      Left            =   345
      TabIndex        =   5
      Top             =   105
      Width           =   3345
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1935
         TabIndex        =   0
         Text            =   "0"
         Top             =   480
         Width           =   435
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Full"
         Height          =   240
         Left            =   615
         TabIndex        =   1
         Top             =   1005
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sub"
         Height          =   240
         Left            =   1785
         TabIndex        =   2
         Top             =   1020
         Width           =   585
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   2370
         TabIndex        =   6
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text1"
         BuddyDispid     =   196614
         OrigLeft        =   2640
         OrigTop         =   360
         OrigRight       =   2880
         OrigBottom      =   615
         Max             =   150
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxLev 异网 > 本网"
         Height          =   180
         Left            =   330
         TabIndex        =   8
         Top             =   510
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-dBm"
         Height          =   180
         Left            =   2685
         TabIndex        =   7
         Top             =   525
         Width           =   360
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   320
      Left            =   2025
      TabIndex        =   4
      Top             =   1890
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   320
      Left            =   810
      TabIndex        =   3
      Top             =   1890
      Width           =   1080
   End
End
Attribute VB_Name = "Cope_RxLev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim my_msg As String
    Dim WinId
    Dim i As Integer
    Dim select_name As String
    
    On Error Resume Next
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
           WinId = mapinfo.eval("windowid(" & i & ")")
           If WinId = mapinfo.eval("frontwindow()") Then
              Exit For
           End If
        End If
    Next
    Select Case Menu_Flag
        Case 888
             If M2_Local = True Then
                select_name = "LocalRxlev"
             Else
                select_name = "OtherRxlev"
             End If
             If Option1.Value = True Then
                If Val(Text1.Text) = 0 Then
                   mapinfo.do "select * From " + tblname + " where (RXLEV_F_2 - rxlev_f > 0 ) into " & select_name
                Else
                   mapinfo.do "select * From " + tblname + " where (RXLEV_F_2 - rxlev_f >= " & Format(Val(Text1.Text)) & " ) into " & select_name
                End If
             Else
                If Val(Text1.Text) = 0 Then
                   mapinfo.do "select * From " + tblname + " where (RXLEV_s_2 - rxlev_s > 0 ) into " & select_name
                Else
                   mapinfo.do "select * From " + tblname + " where (RXLEV_s_2 - rxlev_s >= " & Format(Val(Text1.Text)) & " ) into " & select_name
                End If
             End If
             If Val(mapinfo.eval("tableinfo(" & select_name & ",8)")) = 0 Then
                Unload Me
                Exit Sub
             End If
             mapinfo.do "Add Map window " + WinId + " Layer " & select_name
             If Option1.Value = True Then
                my_msg = "shade window " + WinId + select_name + " With RXLEV_F_2 - rxlev_f "
             Else
                my_msg = "shade window " + WinId + select_name + " With RXLEV_s_2 - rxlev_s "
             End If
             my_msg = my_msg + " ranges apply all use color Symbol(39,65280,8,""MapInfo Cartographic"",0,0) 0: 5 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,5: 10 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,10: 15 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,15: 20 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,20: 25 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,25: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 35 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,35: 40 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,40: 45 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,45: 50 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,50: 55 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,55: 60 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) default Symbol(39,0,8,""MapInfo Cartographic"",0,0)"
             mapinfo.do my_msg
             If legendid = 0 Then
                mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.do "Create Legend From Window  " & WinId
                legendid = mapinfo.eval("windowinfo(1009,12)")
             End If
             If M2_Local = True Then
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网优于异网路段场强差值" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（-dBm）FULL" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""大于 60"" display on"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网优于异网路段场强差值" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（-dBm）SUB" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""大于 60"" display on"
                End If
             Else
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网优于本网路段场强差值" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（-dBm）FULL" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""大于 60"" display on"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网优于本网路段场强差值" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（-dBm）SUB" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""大于 60"" display on"
                End If
             End If
        Case 885
              If M2_Local = False Then
                 select_name = "LocalRxlev"
              Else
                 select_name = "OtherRxlev"
              End If
              If Option1.Value = True Then
                If Val(Text1.Text) = 0 Then
                   mapinfo.do "select * From " + tblname + " where (RXLEV_F - rxlev_f_2 > 0 ) into " & select_name
                Else
                   mapinfo.do "select * From " + tblname + " where (RXLEV_F - rxlev_f_2 >= " & Format(Val(Text1.Text)) & " ) into " & select_name
                End If
             Else
                If Val(Text1.Text) = 0 Then
                   mapinfo.do "select * From " + tblname + " where (RXLEV_s - rxlev_s_2 > 0 ) into " & select_name
                Else
                   mapinfo.do "select * From " + tblname + " where (RXLEV_s - rxlev_s_2 >= " & Format(Val(Text1.Text)) & " ) into " & select_name
                End If
             End If
             If Val(mapinfo.eval("tableinfo(" & select_name & ",8)")) = 0 Then
                Unload Me
                Exit Sub
             End If
             mapinfo.do "Add Map window " + WinId + " Layer " & select_name
             If Option1.Value = True Then
                my_msg = "shade window " + WinId + select_name + " With RXLEV_F - rxlev_f_2 "
             Else
                my_msg = "shade window " + WinId + select_name + " With RXLEV_s - rxlev_s_2 "
             End If
             my_msg = my_msg + " ranges apply all use color Symbol(39,65280,8,""MapInfo Cartographic"",0,0) 0: 5 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,5: 10 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,10: 15 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,15: 20 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,20: 25 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,25: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 35 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,35: 40 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,40: 45 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,45: 50 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,50: 55 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,55: 60 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) default Symbol(39,0,8,""MapInfo Cartographic"",0,0)"
             mapinfo.do my_msg
             If legendid = 0 Then
                mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.do "Create Legend From Window  " & WinId
                legendid = mapinfo.eval("windowinfo(1009,12)")
             End If
             If M2_Local = True Then
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网优于本网路段场强差值" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（-dBm）FULL" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""大于 60"" display on"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网优于本网路段场强差值" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（-dBm）SUB" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""大于 60"" display on"
                End If
             Else
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网优于异网路段场强差值" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（-dBm）FULL" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""大于 60"" display on"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网优于异网路段场强差值" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（-dBm）SUB" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""大于 60"" display on"
                End If
             End If
    End Select
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If Menu_Flag = 888 Then
       If M2_Local = True Then
          Label1.Caption = "RxLev 本网 > 异网"
       End If
    Else
       If M2_Local = False Then
          Label1.Caption = "RxLev 本网 > 异网"
       End If
    End If
End Sub
