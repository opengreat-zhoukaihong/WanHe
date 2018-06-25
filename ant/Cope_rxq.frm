VERSION 5.00
Begin VB.Form Cope_RxQual 
   BackColor       =   &H00C0C0C0&
   Caption         =   "条件选择"
   ClientHeight    =   2640
   ClientLeft      =   3090
   ClientTop       =   2970
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Cope_rxq.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2640
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1785
      Left            =   420
      TabIndex        =   4
      Top             =   150
      Width           =   2610
      Begin VB.OptionButton Option1 
         Caption         =   "Full"
         Height          =   240
         Left            =   480
         TabIndex        =   0
         Top             =   1305
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sub"
         Height          =   240
         Left            =   1485
         TabIndex        =   1
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxQual（本网）> 3"
         Height          =   180
         Left            =   450
         TabIndex        =   6
         Top             =   405
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxQual（异网）< = 3"
         Height          =   180
         Left            =   450
         TabIndex        =   5
         Top             =   795
         Width           =   1710
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   320
      Left            =   555
      TabIndex        =   3
      Top             =   2175
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   320
      Left            =   1785
      TabIndex        =   2
      Top             =   2175
      Width           =   1080
   End
End
Attribute VB_Name = "Cope_RxQual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim WinId
    Dim i As Integer
    Dim my_msg As String
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
        Case 887
             If M2_Local = True Then
                select_name = "LocalRxqual"
             Else
                select_name = "OtherRxqual"
             End If
             If Option1.Value = True Then
                If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                   mapinfo.do "select * From " + tblname + " where (val(RXquql_F_2) <= 3) and (val(RXqual_f) > 3) into " & select_name
                Else
                   mapinfo.do "select * From " + tblname + " where (RXquql_F_2 <= 3) and (RXqual_f > 3) into " & select_name
                End If
             Else
                If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                   mapinfo.do "select * From " + tblname + " where (val(RXquql_s_2) <= 3) and (val(RXqual_s) > 3) into " & select_name
                Else
                   mapinfo.do "select * From " + tblname + " where (RXquql_s_2 <= 3) and (RXqual_s > 3) into " & select_name
                End If
             End If
             If Val(mapinfo.eval("tableinfo(" & select_name & ",8)")) = 0 Then
                Unload Me
                Exit Sub
             End If
             mapinfo.do "Add Map window " + WinId + " Layer " & select_name
             If Option1.Value = True Then
                my_msg = "shade window " + WinId + select_name + " With RXquql_F_2 "
             Else
                my_msg = "shade window " + WinId + select_name + " With RXquql_s_2 "
             End If
             my_msg = my_msg + " values 0 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,16776960,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,32768,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,7 Symbol (41,16719904,8,""MapInfo Cartographic"",0,0) default Symbol(41,0,8,""MapInfo Cartographic"",0,0)"
             mapinfo.do my_msg
             If legendid = 0 Then
                mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                mapinfo.do "Create Legend From Window  " & WinId
                legendid = mapinfo.eval("windowinfo(1009,12)")
             End If
             If M2_Local = True Then
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网优于异网路段品质观测" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（FULL）" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网优于异网路段品质观测" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（SUB）" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                End If
             Else
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网优于本网路段品质观测" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（FULL）" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网优于本网路段品质观测" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（SUB）" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                End If
             End If
        Case 884
             If M2_Local = False Then
                select_name = "LocalRxqual"
             Else
                select_name = "OtherRxqual"
             End If
             If Option1.Value = True Then
                If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                    mapinfo.do "select * From " + tblname + " where (val(RXqual_F) <= 3) and (val(RXquql_f_2) > 3) into " & select_name
                Else
                    mapinfo.do "select * From " + tblname + " where (RXqual_F <= 3) and (RXquql_f_2 > 3) into " & select_name
                End If
             Else
                If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                    mapinfo.do "select * From " + tblname + " where (val(RXqual_s) <= 3) and (val(RXquql_s_2) > 3) into " & select_name
                Else
                    mapinfo.do "select * From " + tblname + " where (RXqual_s <= 3) and (RXquql_s_2 > 3) into " & select_name
                End If
             End If
             If Val(mapinfo.eval("tableinfo(" & select_name & ",8)")) = 0 Then
                Unload Me
                Exit Sub
             End If
             mapinfo.do "Add Map window " + WinId + " Layer " & select_name
             'If Option1.Value = True Then
             '   my_msg = "shade window " + WinId + select_name + " With RXqual_F "
             'Else
             '   my_msg = "shade window " + WinId + select_name + " With RXqual_s "
             'End If
             If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                If Option1.Value = True Then
                    my_msg = " shade window " + WinId + select_name + " With RTrim$(LTrim$(RXqual_F)) values """" Symbol (41,14737632,8,""MapInfo Cartographic"",0,0) ,""0"" Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,""1"" Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,""2"" Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,""3"" Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,""4"" Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,""5"" Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,""6"" Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,""7"" Symbol (41,16719904,8,""MapInfo Cartographic"",0,0) "
                Else
                    my_msg = " shade window " + WinId + select_name + " With RTrim$(LTrim$(RXqual_s)) values """" Symbol (41,14737632,8,""MapInfo Cartographic"",0,0) ,""0"" Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,""1"" Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,""2"" Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,""3"" Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,""4"" Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,""5"" Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,""6"" Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,""7"" Symbol (41,16719904,8,""MapInfo Cartographic"",0,0) "
                End If
             Else
                If Option1.Value = True Then
                   Msg = " shade window " + WinId + select_name + " With RXqual_F values 0 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,7 Symbol (41,16719904,8,""MapInfo Cartographic"",0,0),9 Symbol (41,14737632,8,""MapInfo Cartographic"",0,0)"
                Else
                   Msg = " shade window " + WinId + select_name + " With RXqual_s values 0 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,7 Symbol (41,16719904,8,""MapInfo Cartographic"",0,0),9 Symbol (41,14737632,8,""MapInfo Cartographic"",0,0)"
                End If
             End If
             mapinfo.do my_msg
             If legendid = 0 Then
                mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                mapinfo.do "Create Legend From Window  " & WinId
                legendid = mapinfo.eval("windowinfo(1009,12)")
             End If
             If M2_Local = True Then
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网优于本网路段品质观测" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（FULL）" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "异网优于本网路段品质观测" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（SUB）" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                End If
             Else
                If Option1.Value = True Then
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网优于异网路段品质观测" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（FULL）" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                Else
                   mapinfo.do "set legend window " + WinId + " Layer prev Title " + Chr(34) + "本网优于异网路段品质观测" + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + tblname + "（SUB）" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
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
    If Menu_Flag = 887 Then
       If M2_Local = True Then
          Label1.Caption = "RxQual（异网）> 3"
          Label2.Caption = "RxQual（本网）< = 3"
       End If
    Else
       If M2_Local = False Then
          Label1.Caption = "RxQual（异网）> 3"
          Label2.Caption = "RxQual（本网）< = 3"
       End If
    End If
End Sub
