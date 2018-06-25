VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form Cope_Global 
   BackColor       =   &H00C0C0C0&
   Caption         =   "异网综合质量"
   ClientHeight    =   2550
   ClientLeft      =   3135
   ClientTop       =   3120
   ClientWidth     =   3435
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Cope_glo.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2550
   ScaleWidth      =   3435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   345
      TabIndex        =   6
      Top             =   105
      Width           =   2685
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1155
         TabIndex        =   0
         Text            =   "0"
         Top             =   375
         Width           =   450
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1425
         TabIndex        =   1
         Text            =   "3"
         Top             =   795
         Width           =   450
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Full"
         Height          =   240
         Left            =   420
         TabIndex        =   2
         Top             =   1290
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sub"
         Height          =   240
         Left            =   1560
         TabIndex        =   3
         Top             =   1290
         Width           =   585
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   1605
         TabIndex        =   7
         Top             =   375
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text1"
         BuddyDispid     =   196610
         OrigLeft        =   1815
         OrigTop         =   195
         OrigRight       =   2055
         OrigBottom      =   450
         Max             =   150
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   240
         Left            =   1875
         TabIndex        =   8
         Top             =   810
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   423
         _Version        =   327680
         BuddyControl    =   "Text2"
         BuddyDispid     =   196611
         OrigLeft        =   2145
         OrigTop         =   675
         OrigRight       =   2385
         OrigBottom      =   930
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxQual < = "
         Height          =   180
         Left            =   435
         TabIndex        =   11
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxLev > "
         Height          =   180
         Left            =   435
         TabIndex        =   10
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-dBm"
         Height          =   180
         Left            =   1905
         TabIndex        =   9
         Top             =   420
         Width           =   360
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   320
      Left            =   1725
      TabIndex        =   5
      Top             =   2100
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   320
      Left            =   525
      TabIndex        =   4
      Top             =   2100
      Width           =   1080
   End
End
Attribute VB_Name = "Cope_Global"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim WinId
    Dim i As Integer
    Dim my_msg As String, select_name As String
    
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
        Case 886
             If M2_Local = True Then
                select_name = "LocalGlobal"
             Else
                select_name = "OtherGlobal"
             End If
             If Option1.Value = True Then
                mapinfo.do "select * From " + tblname + " where (RXquql_F_2 <= " + Format(Val(Text2.Text)) + ") and (Rxlev_f_2 > " + Format(Val(Text1.Text)) + ") into " & select_name
             Else
                mapinfo.do "select * From " + tblname + " where (RXquql_s_2 <= " + Format(Val(Text2.Text)) + ") and (Rxlev_s_2 > " + Format(Val(Text1.Text)) + ") into " & select_name
             End If
             If Val(mapinfo.eval("tableinfo(" & select_name & ",8)")) = 0 Then
                Unload Me
                Exit Sub
             End If
             mapinfo.do "Add Map window " + WinId + " Layer " & select_name
             If Option1.Value = True Then
                my_msg = "shade window " + WinId + select_name + " With RXlev_F "
             Else
                my_msg = "shade window " + WinId + select_name + " With RXlev_s "
             End If
             If Legend_Tog = 0 Then
                'my_msg = my_msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,17: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                my_msg = my_msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 35 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
             Else
                my_msg = my_msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
             End If
             mapinfo.do my_msg
             If legendid = 0 Then
                mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                mapinfo.do "Create Legend From Window  " & WinId
                legendid = mapinfo.eval("windowinfo(1009,12)")
             End If
             If M2_Local = True Then
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网综合质量" + tblname + "（FULL）" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev > " + Format(Val(Text1.Text)) + " -dBm  RxQual <= " + Format(Val(Text2.Text)) + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网综合质量" + tblname + "（SUB）" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev > " + Format(Val(Text1.Text)) + " -dBm  RxQual <= " + Format(Val(Text2.Text)) + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                End If
             Else
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网综合质量" + tblname + "（FULL）" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev > " + Format(Val(Text1.Text)) + " -dBm  RxQual <= " + Format(Val(Text2.Text)) + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网综合质量" + tblname + "（SUB）" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev > " + Format(Val(Text1.Text)) + " -dBm  RxQual <= " + Format(Val(Text2.Text)) + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                End If
             End If
        Case 883
             If M2_Local = False Then
                select_name = "LocalGlobal"
             Else
                select_name = "OtherGlobal"
             End If
             If Option1.Value = True Then
                If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                   mapinfo.do "select * From " + tblname + " where (val(RXqual_f) <= " + Format(Val(Text2.Text)) + ") and (Rxlev_f > " + Format(Val(Text1.Text)) + ") into " & select_name
                Else
                   mapinfo.do "select * From " + tblname + " where (RXqual_f <= " + Format(Val(Text2.Text)) + ") and (Rxlev_f > " + Format(Val(Text1.Text)) + ") into " & select_name
                End If
             Else
                If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                   mapinfo.do "select * From " + tblname + " where (val(RXqual_s) <= " + Format(Val(Text2.Text)) + ") and (Rxlev_s > " + Format(Val(Text1.Text)) + ") into " & select_name
                Else
                   mapinfo.do "select * From " + tblname + " where (RXqual_s <= " + Format(Val(Text2.Text)) + ") and (Rxlev_s > " + Format(Val(Text1.Text)) + ") into " & select_name
                End If
             End If
             If Val(mapinfo.eval("tableinfo(" & select_name & ",8)")) = 0 Then
                Unload Me
                Exit Sub
             End If
             mapinfo.do "Add Map window " + WinId + " Layer " & select_name
             If Option1.Value = True Then
                my_msg = "shade window " + WinId + select_name + " With RXlev_F "
             Else
                my_msg = "shade window " + WinId + select_name + " With Rxlev_s "
             End If
             If Legend_Tog = 0 Then
                'my_msg = my_msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,17: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                my_msg = my_msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 35 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
             Else
                my_msg = my_msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
             End If
             mapinfo.do my_msg
             If legendid = 0 Then
                mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                mapinfo.do "Create Legend From Window  " & WinId
                legendid = mapinfo.eval("windowinfo(1009,12)")
             End If
             If M2_Local = True Then
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网综合质量" + tblname + "（FULL）" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev > " + Format(Val(Text1.Text)) + " -dBm  RxQual <= " + Format(Val(Text2.Text)) + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网综合质量" + tblname + "（SUB）" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev > " + Format(Val(Text1.Text)) + " -dBm  RxQual <= " + Format(Val(Text2.Text)) + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                End If
             Else
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网综合质量" + tblname + "（FULL）" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev > " + Format(Val(Text1.Text)) + " -dBm  RxQual <= " + Format(Val(Text2.Text)) + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网综合质量" + tblname + "（SUB）" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev > " + Format(Val(Text1.Text)) + " -dBm  RxQual <= " + Format(Val(Text2.Text)) + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
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
    If Menu_Flag = 886 Then
       If M2_Local = True Then
          Cope_Global.Caption = "本网综合质量"
       End If
    Else
       If M2_Local = False Then
          Cope_Global.Caption = "本网综合质量"
       End If
    End If
End Sub

