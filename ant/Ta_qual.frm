VERSION 5.00
Begin VB.Form Ta_Qual 
   BackColor       =   &H00C0C0C0&
   Caption         =   "孤岛定义条件"
   ClientHeight    =   2115
   ClientLeft      =   2835
   ClientTop       =   2175
   ClientWidth     =   3345
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Ta_qual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2115
   ScaleWidth      =   3345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Full/Sub选择"
      Height          =   765
      Left            =   180
      TabIndex        =   7
      Top             =   1245
      Width           =   1785
      Begin VB.OptionButton Option5 
         Caption         =   "Sub"
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   330
         Width           =   570
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Full"
         Height          =   300
         Left            =   210
         TabIndex        =   8
         Top             =   345
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "门限"
      Height          =   1080
      Left            =   180
      TabIndex        =   4
      Top             =   90
      Width           =   1770
      Begin VB.TextBox RxLevValue 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   990
         MaxLength       =   1
         TabIndex        =   0
         Text            =   "2"
         Top             =   285
         Width           =   450
      End
      Begin VB.TextBox RxQualValue 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   990
         MaxLength       =   1
         TabIndex        =   1
         Text            =   "4"
         Top             =   630
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TA> "
         Height          =   180
         Index           =   6
         Left            =   660
         TabIndex        =   6
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxQual<"
         Height          =   180
         Index           =   7
         Left            =   285
         TabIndex        =   5
         Top             =   675
         Width           =   630
      End
   End
   Begin VB.CommandButton SBSCancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   2160
      TabIndex        =   3
      Top             =   585
      Width           =   1080
   End
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   2160
      TabIndex        =   2
      Top             =   195
      Width           =   1080
   End
End
Attribute VB_Name = "Ta_Qual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Full_Sub As Integer


Private Sub Option4_Click()
    On Error Resume Next
    Full_Sub = 0
End Sub

Private Sub Option5_Click()
    On Error Resume Next
    Full_Sub = 1
End Sub

Private Sub SBSCancel_Click()
   On Error Resume Next
   Ta_Qual.Hide
   Unload Ta_Qual
End Sub

Private Sub SBSOK_Click()
  Dim Ta1, Rxqual1 As Integer
  Dim Iland As String
  
  Dim i, col_num, Maxbsic, Minbsic As Integer
  Dim Name, str As String
  Dim subtitle_str As String
  Dim ta_str As String, rxq_str As String
  Dim WindId As Variant
   
  On Error Resume Next
  Ta1 = Val(RxLevValue.Text)
  Rxqual1 = Val(RxQualValue.Text)
  
  Unload Me

  If tblname <> "" Then
  Select Case Menu_Flag
  Case 431
       Iland = "Iland"
       On Error Resume Next
       For i = 1 To mapinfo.eval("NumWindows()")
           If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
              WinId = mapinfo.eval("windowid(" & i & ")")
              If WinId = mapinfo.eval("frontwindow()") Then
                 Exit For
              End If
           End If
       Next
       mapinfo.do "fetch first from " & tblname
       If Full_Sub = 0 Then
          If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
             Msg = "select * From " + tblname + " where ( val(Ta) >  " & Ta1 & " ) And (val(RXQUAL_F) < " & Rxqual1 & " ) And (BCCH_SERV = " & Val(rmsg2) & ") And (ci_serv = " + Chr(34) + rmsg1 + Chr(34) + " ) into  " & Iland
          Else
             Msg = "select * From " + tblname + " where ( val(Ta) >  " & Ta1 & " ) And (RXQUAL_F < " & Rxqual1 & " ) And (BCCH_SERV = " & Val(rmsg2) & ") And (ci_serv = " + Chr(34) + rmsg1 + Chr(34) + " ) into  " & Iland
          End If
       Else
          If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
             Msg = "select * From " + tblname + " where ( val(Ta) >  " & Ta1 & " ) And (val(RXQUAL_S) < " & Rxqual1 & " ) And (BCCH_SERV = " & Val(rmsg2) & ") And (ci_serv = " + Chr(34) + rmsg1 + Chr(34) + " ) into  " & Iland
          Else
             Msg = "select * From " + tblname + " where ( val(Ta) >  " & Ta1 & " ) And (RXQUAL_S < " & Rxqual1 & " ) And (BCCH_SERV = " & Val(rmsg2) & ") And (ci_serv = " + Chr(34) + rmsg1 + Chr(34) + " ) into  " & Iland
          End If
       End If
       mapinfo.do Msg
       If Val(mapinfo.eval("tableinfo(Iland,8)")) = 0 Then
          MsgBox "不存在岛效应点", 64, "提示"
          mapinfo.do "close table Iland"
          Exit Sub
       End If
       mapinfo.do "Add Map window " & WinId & " Layer " & Iland

       If Full_Sub = 0 Then
          If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
             Msg = " shade window " & WinId + " " + Iland + " With RTrim$(LTrim$(RXQUAL_F)) values """" Symbol (41,14737632,8,""MapInfo Cartographic"",0,0) ,""0"" Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,""1"" Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,""2"" Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,""3"" Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,""4"" Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,""5"" Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,""6"" Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,""7"" Symbol (41,16719904,8,""MapInfo Cartographic"",0,0) "
          Else
             Msg = " shade window " & WinId + " " + Iland + " With RXQUAL_F values 0 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,7 Symbol (41,16719904,8,""MapInfo Cartographic"",0,0),9 Symbol (41,14737632,8,""MapInfo Cartographic"",0,0)"
          End If
       Else
          If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
             Msg = " shade window " & WinId + " " + Iland + " With RTrim$(LTrim$(RXQUAL_s)) values """" Symbol (41,14737632,8,""MapInfo Cartographic"",0,0) ,""0"" Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,""1"" Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,""2"" Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,""3"" Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,""4"" Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,""5"" Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,""6"" Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,""7"" Symbol (41,16719904,8,""MapInfo Cartographic"",0,0) "
          Else
             Msg = " shade window " & WinId + " " + Iland + " With RXQUAL_s values 0 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,7 Symbol (41,16719904,8,""MapInfo Cartographic"",0,0),9 Symbol (41,14737632,8,""MapInfo Cartographic"",0,0)"
          End If
       End If
       mapinfo.do Msg
       
       If legendid = 0 Then
          mapinfo.do "set next document parent " & MDIMain.hWnd & "style 0"
          mapinfo.do "Create Legend From Window  " & WinId
          legendid = mapinfo.eval("windowinfo(1009,12)")
       End If
       ta_str = CStr(Ta1 * 500)
       rxq_str = CStr(Rxqual1)
       If Full_Sub = 0 Then
          subtitle_str = "覆盖大于" + ta_str + "米且RxQual_f<" & RxQualValue.Text & "存在的孤岛"
       Else
          subtitle_str = "覆盖大于" + ta_str + "米且RxQual_s<" & RxQualValue.Text & "存在的孤岛"
       End If
       'msg = " Title " + Chr(34) + "岛效应点观测(RxQual) " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + subtitle_str + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
       Msg = " Title " + Chr(34) + "岛效应点观测" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + subtitle_str + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
       mapinfo.do "set legend window " & WinId & " Layer prev  display on shades off symbols on lines off count on" & Msg
    End Select
 End If
' Unload Ta_Qual
End Sub
