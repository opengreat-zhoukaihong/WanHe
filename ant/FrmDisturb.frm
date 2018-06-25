VERSION 5.00
Begin VB.Form FrmDisturb 
   Caption         =   "Ncell 对 BCCH 频率碰撞"
   ClientHeight    =   2835
   ClientLeft      =   6105
   ClientTop       =   2940
   ClientWidth     =   3945
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDisturb.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   3945
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   2730
      TabIndex        =   11
      Top             =   255
      Width           =   1080
   End
   Begin VB.CommandButton SBSCancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   2730
      TabIndex        =   10
      Top             =   645
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      Caption         =   "Full/Sub选择"
      Height          =   780
      Left            =   165
      TabIndex        =   1
      Top             =   1950
      Width           =   2385
      Begin VB.OptionButton Option5 
         Caption         =   "Sub"
         Height          =   315
         Left            =   1455
         TabIndex        =   3
         Top             =   315
         Width           =   570
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Full"
         Height          =   300
         Left            =   300
         TabIndex        =   2
         Top             =   330
         Value           =   -1  'True
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "分析选择"
      Height          =   1755
      Left            =   165
      TabIndex        =   0
      Top             =   135
      Width           =   2385
      Begin VB.TextBox RxLevValue 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   1
         Left            =   1455
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "-9"
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox RxLevValue 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   0
         Left            =   1455
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "12"
         Top             =   1350
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "邻频碰撞"
         Height          =   240
         Left            =   285
         TabIndex        =   5
         Top             =   375
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "同频碰撞"
         Height          =   240
         Left            =   285
         TabIndex        =   4
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C/A < "
         Height          =   180
         Index           =   0
         Left            =   885
         TabIndex        =   7
         Top             =   705
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C/I < "
         Height          =   180
         Index           =   6
         Left            =   885
         TabIndex        =   6
         Top             =   1380
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmDisturb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    If Menu_Flag = 442 Then
       Caption = "Ncell 对 TCH 频率碰撞"
    End If
End Sub

Private Sub SBSCancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub SBSOK_Click()
    Dim MyRow As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim Ci_Serv(1 To 100) As String
    Dim Bcch_Serv(1 To 100) As String
    Dim Ci_Name(1 To 100) As String
    Dim my_msg As Variant
    Dim My_Color As Variant
    Dim finds As Integer
    Dim SortTemp As String, SortBcch(1 To 100) As String
    Dim TempNum As Integer
    Dim SourceBcch(1 To 100) As String
    Dim SortSource(1 To 100) As String
    Dim SortCi(1 To 100) As String
    Dim MyRxlev As String
    Dim OpenTableNum As Integer
    Dim No_PZ1 As Boolean, No_PZ2 As Boolean
    
         On Error Resume Next
         Me.Hide
         If Check1.Value = 1 Then
       
       OpenTableNum = mapinfo.eval("NumTables()")
       For i = 1 To OpenTableNum
           If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "SARF_JAM" Then
              mapinfo.do "close table sARF_Jam"
              Exit For
           End If
       Next
         
            If Option4.Value = True Then
               If Menu_Flag = 441 Then    '上行干扰
                  mapinfo.do "select * From " + tblname + " where (BCCH_SERV= Bcch_N1 AND bsic_serv <> bsic_n1 and abs(RXLEV_F-Rxlev_n1) < " & Format(Val(RxLevValue(0))) & " ) OR (BCCH_SERV= Bcch_N2 AND bsic_serv <> bsic_n2 AND abs(RXLEV_F-Rxlev_n2) < " & Format(Val(RxLevValue(0))) & "  )OR (BCCH_SERV= Bcch_N3 AND bsic_serv <> bsic_n3 AND abs(RXLEV_F-Rxlev_n3) < " & Format(Val(RxLevValue(0))) & "  ) OR (BCCH_SERV= Bcch_N4 AND bsic_serv <> bsic_n4 AND abs(RXLEV_F-Rxlev_n4) < " & Format(Val(RxLevValue(0))) & "  ) OR (BCCH_SERV= Bcch_N5 AND bsic_serv <> bsic_n5 AND abs(RXLEV_F-Rxlev_n5) < " & Format(Val(RxLevValue(0))) & " )OR (BCCH_SERV= Bcch_N6 AND bsic_serv <> bsic_n6 AND abs(RXLEV_F-Rxlev_n6) < " & Format(Val(RxLevValue(0))) & "  ) Into SARF_Jam "
               Else                       '下行干扰
                  mapinfo.do "select * From " + tblname + " where (val(num_dch)>0 and val(num_dch)= Bcch_N1 AND bsic_serv <> bsic_n1 and abs(RXLEV_F-Rxlev_n1) < " & Format(Val(RxLevValue(0))) & " ) OR (val(num_dch)>0 and val(num_dch)= Bcch_N2 AND bsic_serv <> bsic_n2 AND abs(RXLEV_F-Rxlev_n2) < " & Format(Val(RxLevValue(0))) & "  )OR (val(num_dch)>0 and val(num_dch)= Bcch_N3 AND bsic_serv <> bsic_n3 AND abs(RXLEV_F-Rxlev_n3) < " & Format(Val(RxLevValue(0))) & "  ) OR (val(num_dch)>0 and val(num_dch)= Bcch_N4 AND bsic_serv <> bsic_n4 AND abs(RXLEV_F-Rxlev_n4) < " & Format(Val(RxLevValue(0))) & "  ) OR (val(num_dch)>0 and val(num_dch)= Bcch_N5 AND bsic_serv <> bsic_n5 AND abs(RXLEV_F-Rxlev_n5) < " & Format(Val(RxLevValue(0))) & " )OR (val(num_dch)>0 and val(num_dch)= Bcch_N6 AND bsic_serv <> bsic_n6 AND abs(RXLEV_F-Rxlev_n6) < " & Format(Val(RxLevValue(0))) & "  ) Into SARF_Jam "
               End If
            Else
               If Menu_Flag = 441 Then
                  mapinfo.do "select * From " + tblname + " where (BCCH_SERV= Bcch_N1 AND bsic_serv <> bsic_n1 AND abs(RXLEV_s-Rxlev_n1) < " & Format(Val(RxLevValue(0))) & " ) OR (BCCH_SERV= Bcch_N2 AND bsic_serv <> bsic_n2 AND abs(RXLEV_s-Rxlev_n2) < " & Format(Val(RxLevValue(0))) & "  )OR (BCCH_SERV= Bcch_N3 AND bsic_serv <> bsic_n3 AND abs(RXLEV_s-Rxlev_n3) < " & Format(Val(RxLevValue(0))) & "  ) OR (BCCH_SERV= Bcch_N4 AND bsic_serv <> bsic_n4 AND abs(RXLEV_s-Rxlev_n4) < " & Format(Val(RxLevValue(0))) & "  ) OR (BCCH_SERV= Bcch_N5 AND bsic_serv <> bsic_n5 AND abs(RXLEV_s-Rxlev_n5) < " & Format(Val(RxLevValue(0))) & " )OR (BCCH_SERV= Bcch_N6 AND bsic_serv <> bsic_n6 AND abs(RXLEV_s-Rxlev_n6) < " & Format(Val(RxLevValue(0))) & "  ) Into SARF_Jam "
               Else
                  mapinfo.do "select * From " + tblname + " where (val(num_dch)>0 and val(num_dch)= Bcch_N1 AND bsic_serv <> bsic_n1 AND abs(RXLEV_s-Rxlev_n1) < " & Format(Val(RxLevValue(0))) & " ) OR (val(num_dch)>0 and val(num_dch)= Bcch_N2 AND bsic_serv <> bsic_n2 AND abs(RXLEV_s-Rxlev_n2) < " & Format(Val(RxLevValue(0))) & "  )OR (val(num_dch)>0 and val(num_dch)= Bcch_N3 AND bsic_serv <> bsic_n3 AND abs(RXLEV_s-Rxlev_n3) < " & Format(Val(RxLevValue(0))) & "  ) OR (val(num_dch)>0 and val(num_dch)= Bcch_N4 AND bsic_serv <> bsic_n4 AND abs(RXLEV_s-Rxlev_n4) < " & Format(Val(RxLevValue(0))) & "  ) OR (val(num_dch)>0 and val(num_dch)= Bcch_N5 AND bsic_serv <> bsic_n5 AND abs(RXLEV_s-Rxlev_n5) < " & Format(Val(RxLevValue(0))) & " )OR (val(num_dch)>0 and val(num_dch)= Bcch_N6 AND bsic_serv <> bsic_n6 AND abs(RXLEV_s-Rxlev_n6) < " & Format(Val(RxLevValue(0))) & "  ) Into SARF_Jam "
               End If
            End If
            MyRow = Val(mapinfo.eval("tableinfo(sarf_jam,8)"))
            If MyRow = 0 Then
               No_PZ1 = True
               GoTo Non_1
            End If
            Gsm_FileName = Gsm_Path + "\sarf_jam.tab"
            mapinfo.do "commit table sarf_jam as " + Chr(34) + Gsm_FileName + Chr(34)
            mapinfo.do "close table sarf_jam"
            mapinfo.do "open table " + Chr(34) + Gsm_FileName + Chr(34)
            'mapinfo.do "Alter Table ""sarf_jam"" ( add C_I Decimal(3,0) ) Interactive"
            mapinfo.do "Alter Table ""sarf_jam"" ( add C_I Decimal(4,0) ) Interactive"
            mapinfo.do "fetch first from sarf_jam"
            
            If Menu_Flag = 441 Then
               j = 1
               For i = 1 To MyRow
                   If Val(mapinfo.eval("sarf_jam.bcch_serv")) > 0 Then
                      If j = 1 Then
                         Bcch_Serv(j) = mapinfo.eval("sarf_jam.bcch_serv")
                         Ci_Serv(j) = mapinfo.eval("sarf_jam.ci_serv")
                         j = j + 1
                      Else
                         If Bcch_Serv(j - 1) <> mapinfo.eval("sarf_jam.bcch_serv") Then
                            For k = 1 To j - 1
                                If Bcch_Serv(k) = mapinfo.eval("sarf_jam.bcch_serv") Then
                                   GoTo Next_Point1
                                End If
                            Next
                            Bcch_Serv(j) = mapinfo.eval("sarf_jam.bcch_serv")
                            Ci_Serv(j) = mapinfo.eval("sarf_jam.ci_serv")
                            j = j + 1
                         End If
                      End If
                   End If
Next_Point1:
                            For k = 1 To 6
                                If Option4.Value Then
                                   If Abs(Val(mapinfo.eval("sARF_Jam.bcch_n" & Format(k))) - Val(mapinfo.eval("sARF_Jam.bcch_serv"))) = 0 And Abs(Val(mapinfo.eval("sARF_Jam.rxlev_f")) - Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k)))) < Val(RxLevValue(0)) Then
                                      'MyRxlev = Format(Abs(Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k))) - Val(mapinfo.eval("sARF_Jam.rxlev_f"))))
                                      MyRxlev = Format(Val(mapinfo.eval("sARF_Jam.rxlev_f")) - Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k))))
                                      mapinfo.do "UPDATE sARF_Jam set c_i = " & MyRxlev & " where rowid = " & Format(i)
                                      Exit For
                                   End If
                                Else
                                   If Abs(Val(mapinfo.eval("sARF_Jam.bcch_n" & Format(k))) - Val(mapinfo.eval("sARF_Jam.bcch_serv"))) = 0 And Abs(Val(mapinfo.eval("sARF_Jam.rxlev_s")) - Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k)))) < Val(RxLevValue(0)) Then
                                      'MyRxlev = Format(Abs(Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k))) - Val(mapinfo.eval("sARF_Jam.rxlev_s"))))
                                      MyRxlev = Format(Val(mapinfo.eval("sARF_Jam.rxlev_s")) - Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k))))
                                      mapinfo.do "UPDATE sARF_Jam set c_i = " & MyRxlev & " where rowid = " & Format(i)
                                      Exit For
                                   End If
                                End If
                            Next

                   mapinfo.do "fetch next from sarf_jam"
               Next
               mapinfo.do "commit table sarf_jam"
               For i = 1 To j - 1
                   SortTemp = Bcch_Serv(1)
                   TempNum = 1
                   For k = 1 To j - 1
                       If SortTemp < Bcch_Serv(k) Then
                          SortTemp = Bcch_Serv(k)
                          TempNum = k
                       End If
                   Next
                   SortBcch(j - i) = SortTemp
                   Bcch_Serv(TempNum) = ""
                   SortCi(j - i) = Ci_Serv(TempNum)
               Next
               
            mapinfo.do "Add Map window Frontwindow() Layer SARF_Jam"
            
               my_msg = "shade window FrontWindow() sarf_jam With bcch_serv values "
               My_Color = "10535167"
               For i = 1 To j - 1
                   If i = j - 1 Then
                      my_msg = my_msg + SortBcch(i) + " Symbol (57," + My_Color + ",10,""MapInfo Symbols"",0,0) "
                      my_msg = my_msg + "default Symbol(57,0,10,""MapInfo Symbols"",0,0)"
                   Else
                      my_msg = my_msg + SortBcch(i) + " Symbol (57," + My_Color + ",10,""MapInfo Symbols"",0,0),"
                   End If
                   My_Color = Format(Val(My_Color) + 6000)
                   Ci_Name(i) = Findcell(SortCi(i))
                   finds = InStr(Ci_Name(i), Chr(0))
                   If finds > 0 Then
                      Ci_Name(i) = Trim(Left(Ci_Name(i), finds - 1))
                   End If
                   If Ci_Name(i) = "" Then
                      Ci_Name(i) = "    "
                   End If
               Next
               mapinfo.do my_msg
               my_msg = ""
               For i = 1 To j - 1
                   my_msg = my_msg + "," + Chr(34) + SortBcch(i) + " [" + Ci_Name(i) + "]" + Chr(34) + "display on"
               Next
               'my_msg = " Title " + Chr(34) + "上行同频干扰观测 OF " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "Ncell 对 BCCH 干扰" + Chr(34) + "Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""无服务小区"" display on" + my_msg
               If Option4.Value = True Then
                  my_msg = " Title " + Chr(34) + "Ncell 对 BCCH 同频碰撞分布" + Chr(34) + "Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "被碰撞小区（C/I<" & Trim(RxLevValue(0)) & "）标注：载干比" + Chr(34) + "Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off " + my_msg
               Else
                  my_msg = " Title " + Chr(34) + "Ncell 对 BCCH 同频碰撞分布" + Chr(34) + "Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "被碰撞小区（C/I<" & Trim(RxLevValue(0)) & "）标注：载干比" + Chr(34) + "Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off " + my_msg
               End If
               mapinfo.do "set legend window FrontWindow() Layer prev " & my_msg
               mapinfo.do "set map redraw off"
               'mapinfo.do "Set Map Layer sARF_Jam Label Visibility Font (""Arial"",257,8,16711680,16777215) With c_i Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
               mapinfo.do "Set Map Layer sARF_Jam Label Visibility Font (""Arial"",257,8,16711680,16777215) With c_i Auto On Overlap Off Duplicates On Position Below Auto On Offset 2"
               mapinfo.do "set map redraw on"
            Else
               j = 1
               For i = 1 To MyRow
                   If Val(mapinfo.eval("sarf_jam.num_dch")) > 0 Then
                      If j = 1 Then
                         Bcch_Serv(j) = mapinfo.eval("sarf_jam.num_dch")
                         Ci_Serv(j) = mapinfo.eval("sarf_jam.ci_serv")
                         j = j + 1
                      Else
                         If Bcch_Serv(j - 1) <> mapinfo.eval("sarf_jam.num_dch") Then
                            For k = 1 To j - 1
                                If Bcch_Serv(k) = mapinfo.eval("sarf_jam.num_dch") Then
                                   GoTo Next_Point2
                                End If
                            Next
                            Bcch_Serv(j) = mapinfo.eval("sarf_jam.num_dch")
                            Ci_Serv(j) = mapinfo.eval("sarf_jam.ci_serv")
                            j = j + 1
                         End If
                      End If
                   End If
Next_Point2:
                            For k = 1 To 6
                                If Option4.Value Then
                                   If Abs(Val(mapinfo.eval("sARF_Jam.bcch_n" & Format(k))) - Val(mapinfo.eval("sARF_Jam.num_dch"))) = 0 And Abs(Val(mapinfo.eval("sARF_Jam.rxlev_f")) - Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k)))) < Val(RxLevValue(0)) Then
                                      'MyRxlev = Format(Abs(Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k))) - Val(mapinfo.eval("sARF_Jam.rxlev_f"))))
                                      MyRxlev = Format(Val(mapinfo.eval("sARF_Jam.rxlev_f") - Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k)))))
                                      mapinfo.do "UPDATE sARF_Jam set c_i = " & MyRxlev & " where rowid = " & Format(i)
                                      Exit For
                                   End If
                                Else
                                   If Abs(Val(mapinfo.eval("sARF_Jam.bcch_n" & Format(k))) - Val(mapinfo.eval("sARF_Jam.num_dch"))) = 0 And Abs(Val(mapinfo.eval("sARF_Jam.rxlev_s")) - Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k)))) < Val(RxLevValue(0)) Then
                                      'MyRxlev = Format(Abs(Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k))) - Val(mapinfo.eval("sARF_Jam.rxlev_s"))))
                                      MyRxlev = Format(Val(mapinfo.eval("sARF_Jam.rxlev_s")) - Val(mapinfo.eval("sARF_Jam.rxlev_n" & Format(k))))
                                      mapinfo.do "UPDATE sARF_Jam set c_i = " & MyRxlev & " where rowid = " & Format(i)
                                      Exit For
                                   End If
                                End If
                            Next

                   mapinfo.do "fetch next from sarf_jam"
               Next
               mapinfo.do "commit table sarf_jam"
               For i = 1 To j - 1
                   SortTemp = Bcch_Serv(1)
                   TempNum = 1
                   For k = 1 To j - 1
                       If SortTemp < Bcch_Serv(k) Then
                          SortTemp = Bcch_Serv(k)
                          TempNum = k
                       End If
                   Next
                   SortBcch(j - i) = SortTemp
                   Bcch_Serv(TempNum) = ""
                   SortCi(j - i) = Ci_Serv(TempNum)
               Next
               
            mapinfo.do "Add Map window Frontwindow() Layer SARF_Jam"
            
               my_msg = "shade window FrontWindow() sarf_jam With val(num_dch) values "
               My_Color = "10535167"
               For i = 1 To j - 1
                   If i = j - 1 Then
                      my_msg = my_msg + SortBcch(i) + " Symbol (57," + My_Color + ",10,""MapInfo Symbols"",0,0) "
                      my_msg = my_msg + "default Symbol(57,0,10,""MapInfo Symbols"",0,0)"
                   Else
                      my_msg = my_msg + SortBcch(i) + " Symbol (57," + My_Color + ",10,""MapInfo Symbols"",0,0),"
                   End If
                   My_Color = Format(Val(My_Color) + 6000)
                   Ci_Name(i) = Findcell(SortCi(i))
                   finds = InStr(Ci_Name(i), Chr(0))
                   If finds > 0 Then
                      Ci_Name(i) = Trim(Left(Ci_Name(i), finds - 1))
                   End If
                   If Ci_Name(i) = "" Then
                      Ci_Name(i) = "    "
                   End If
               Next
               mapinfo.do my_msg
               my_msg = ""
               For i = 1 To j - 1
                   my_msg = my_msg + "," + Chr(34) + SortBcch(i) + " [" + Ci_Name(i) + "]" + Chr(34) + "display on"
               Next
               If Option4.Value = True Then
                  my_msg = " Title " + Chr(34) + "Ncell 对 TCH 同频碰撞分布" + Chr(34) + "Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "被碰撞小区（C/I<" & Trim(RxLevValue(0)) & "）标注：载干比" + Chr(34) + "Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off " + my_msg
               Else
                  my_msg = " Title " + Chr(34) + "Ncell 对 TCH 同频碰撞分布" + Chr(34) + "Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "被碰撞小区（C/I<" & Trim(RxLevValue(0)) & "）标注：载干比" + Chr(34) + "Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off " + my_msg
               End If
               mapinfo.do "set legend window FrontWindow() Layer prev " & my_msg
               mapinfo.do "set map redraw off"
               'mapinfo.do "Set Map Layer sARF_Jam Label Visibility Font (""Arial"",257,8,16711680,16777215) With c_i Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
               mapinfo.do "Set Map Layer sARF_Jam Label Visibility Font (""Arial"",257,8,16711680,16777215) With c_i Auto On Overlap Off Duplicates On Position Below Auto On Offset 2"
               mapinfo.do "set map redraw on"
            
            End If
         End If
Non_1:
         If Check2.Value = 1 Then
            On Error Resume Next
       OpenTableNum = mapinfo.eval("NumTables()")
       For i = 1 To OpenTableNum
           If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "NARF_JAM" Then
              mapinfo.do "close table NARF_Jam"
              Exit For
           End If
       Next
            
            If Option4.Value = True Then
               If Menu_Flag = 441 Then
                  'msg = "select * From " + tblname + " where (ABS(BCCH_SERV-Bcch_N1)=1 AND abs(RXLEV_f-Rxlev_n1) < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N2)=1 AND abs(RXLEV_f-Rxlev_n2) < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N3)=1 AND abs(RXLEV_f-Rxlev_n3) < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N4)=1 AND abs(RXLEV_f-Rxlev_n4) < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N5)=1 AND abs(RXLEV_f-Rxlev_n5) < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N6)=1 AND abs(RXLEV_f-Rxlev_n6) < " & Format(Val(RxLevValue(1))) & " )  Into NARF_Jam "
                  msg = "select * From " + tblname + " where (ABS(BCCH_SERV-Bcch_N1)=1 AND RXLEV_f-Rxlev_n1 < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N2)=1 AND RXLEV_f-Rxlev_n2 < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N3)=1 AND RXLEV_f-Rxlev_n3 < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N4)=1 AND RXLEV_f-Rxlev_n4 < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N5)=1 AND RXLEV_f-Rxlev_n5 < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N6)=1 AND RXLEV_f-Rxlev_n6 < " & Format(Val(RxLevValue(1))) & " )  Into NARF_Jam "
               Else
                  'msg = "select * From " + tblname + " where (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N1)=1 AND abs(RXLEV_F-Rxlev_n1) < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N2)=1 AND abs(RXLEV_F-Rxlev_n2) < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N3)=1 AND abs(RXLEV_F-Rxlev_n3) < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N4)=1 AND abs(RXLEV_F-Rxlev_n4) < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N5)=1 AND abs(RXLEV_F-Rxlev_n5) < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N6)=1 AND abs(RXLEV_F-Rxlev_n6) < " & Format(Val(RxLevValue(1))) & " )  Into NARF_Jam "
                  msg = "select * From " + tblname + " where (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N1)=1 AND RXLEV_F-Rxlev_n1 < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N2)=1 AND RXLEV_F-Rxlev_n2 < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N3)=1 AND RXLEV_F-Rxlev_n3 < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N4)=1 AND RXLEV_F-Rxlev_n4 < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N5)=1 AND RXLEV_F-Rxlev_n5 < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N6)=1 AND RXLEV_F-Rxlev_n6 < " & Format(Val(RxLevValue(1))) & " )  Into NARF_Jam "
               End If
            Else
               If Menu_Flag = 441 Then
                  'msg = "select * From " + tblname + " where (ABS(BCCH_SERV-Bcch_N1)=1 AND abs(RXLEV_s-Rxlev_n1) < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N2)=1 AND abs(RXLEV_s-Rxlev_n2) < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N3)=1 AND abs(RXLEV_s-Rxlev_n3) < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N4)=1 AND abs(RXLEV_s-Rxlev_n4) < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N5)=1 AND abs(RXLEV_s-Rxlev_n5) < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N6)=1 AND abs(RXLEV_s-Rxlev_n6) < " & Format(Val(RxLevValue(1))) & " )  Into NARF_Jam "
                  msg = "select * From " + tblname + " where (ABS(BCCH_SERV-Bcch_N1)=1 AND RXLEV_s-Rxlev_n1 < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N2)=1 AND RXLEV_s-Rxlev_n2 < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N3)=1 AND RXLEV_s-Rxlev_n3 < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N4)=1 AND RXLEV_s-Rxlev_n4 < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N5)=1 AND RXLEV_s-Rxlev_n5 < " & Format(Val(RxLevValue(1))) & " ) OR (ABS(BCCH_SERV-Bcch_N6)=1 AND RXLEV_s-Rxlev_n6 < " & Format(Val(RxLevValue(1))) & " )  Into NARF_Jam "
               Else
                  'msg = "select * From " + tblname + " where (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N1)=1 AND abs(RXLEV_s-Rxlev_n1) < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N2)=1 AND abs(RXLEV_s-Rxlev_n2) < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N3)=1 AND abs(RXLEV_s-Rxlev_n3) < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N4)=1 AND abs(RXLEV_s-Rxlev_n4) < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N5)=1 AND abs(RXLEV_s-Rxlev_n5) < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N6)=1 AND abs(RXLEV_s-Rxlev_n6) < " & Format(Val(RxLevValue(1))) & " )  Into NARF_Jam "
                  msg = "select * From " + tblname + " where (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N1)=1 AND RXLEV_s-Rxlev_n1 < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N2)=1 AND RXLEV_s-Rxlev_n2 < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N3)=1 AND RXLEV_s-Rxlev_n3 < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N4)=1 AND RXLEV_s-Rxlev_n4 < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N5)=1 AND RXLEV_s-Rxlev_n5 < " & Format(Val(RxLevValue(1))) & " ) OR (val(num_dch)>0 and ABS(val(num_dch)-Bcch_N6)=1 AND RXLEV_s-Rxlev_n6 < " & Format(Val(RxLevValue(1))) & " )  Into NARF_Jam "
               End If
            End If
            mapinfo.do msg
            MyRow = Val(mapinfo.eval("tableinfo(NARF_Jam,8)"))
            If MyRow = 0 Then
               No_PZ2 = True
               GoTo Non_2
            End If
            Gsm_FileName = Gsm_Path + "\Narf_Jam.tab"
            mapinfo.do "commit table narf_jam as " + Chr(34) + Gsm_FileName + Chr(34)
            mapinfo.do "close table narf_jam"
            mapinfo.do "open table " + Chr(34) + Gsm_FileName + Chr(34)
            'mapinfo.do "Alter Table ""narf_jam"" ( add C_I Decimal(3,0) ) Interactive"
            mapinfo.do "Alter Table ""narf_jam"" ( add C_I Decimal(4,0) ) Interactive"
            mapinfo.do "fetch first from NARF_Jam"
            If Menu_Flag = 441 Then
               j = 1
               For i = 1 To MyRow
                   If Val(mapinfo.eval("NARF_Jam.bcch_serv")) > 0 Then
                      If j = 1 Then
                         Bcch_Serv(j) = mapinfo.eval("NARF_Jam.bcch_serv")
                         Ci_Serv(j) = mapinfo.eval("NARF_Jam.ci_serv")
                         For k = 1 To 6
                             If Abs(Val(mapinfo.eval("NARF_Jam.bcch_n" & Format(k))) - Val(Bcch_Serv(j))) = 1 Then
                                SourceBcch(j) = mapinfo.eval("NARF_Jam.bcch_n" & Format(k))
                                Exit For
                             End If
                         Next
                         j = j + 1
                      Else
                         If Bcch_Serv(j - 1) <> mapinfo.eval("NARF_Jam.bcch_serv") Then
                            For k = 1 To j - 1
                                If Bcch_Serv(k) = mapinfo.eval("NARF_Jam.bcch_serv") Then
                                   GoTo Next_Point3
                                End If
                            Next
                            Bcch_Serv(j) = mapinfo.eval("NARF_Jam.bcch_serv")
                            Ci_Serv(j) = mapinfo.eval("NARF_Jam.ci_serv")
                            For k = 1 To 6
                                If Abs(Val(mapinfo.eval("NARF_Jam.bcch_n" & Format(k))) - Val(Bcch_Serv(j))) = 1 Then
                                   SourceBcch(j) = mapinfo.eval("NARF_Jam.bcch_n" & Format(k))
                                   Exit For
                                End If
                            Next
                            
                            j = j + 1
                         End If
                      End If
                   End If
Next_Point3:
                            For k = 1 To 6
                                If Option4.Value Then
                                   If Abs(Val(mapinfo.eval("NARF_Jam.bcch_n" & Format(k))) - Val(mapinfo.eval("NARF_Jam.bcch_serv"))) = 1 And Val(mapinfo.eval("NARF_Jam.rxlev_f")) - Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))) < Val(RxLevValue(1)) Then
                                      'MyRxlev = Format(Abs(Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))) - Val(mapinfo.eval("NARF_Jam.rxlev_f"))))
                                      MyRxlev = Format(Val(mapinfo.eval("NARF_Jam.rxlev_f")) - Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))))
                                      mapinfo.do "UPDATE NARF_Jam set c_i = " & MyRxlev & " where rowid = " & Format(i)
                                      Exit For
                                   End If
                                Else
                                   If Abs(Val(mapinfo.eval("NARF_Jam.bcch_n" & Format(k))) - Val(mapinfo.eval("NARF_Jam.bcch_serv"))) = 1 And Val(mapinfo.eval("NARF_Jam.rxlev_s")) - Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))) < Val(RxLevValue(1)) Then
                                      'MyRxlev = Format(Abs(Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))) - Val(mapinfo.eval("NARF_Jam.rxlev_s"))))
                                      MyRxlev = Format(Val(mapinfo.eval("NARF_Jam.rxlev_s")) - Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))))
                                      mapinfo.do "UPDATE NARF_Jam set c_i = " & MyRxlev & " where rowid = " & Format(i)
                                      Exit For
                                   End If
                                End If
                            Next

                   mapinfo.do "fetch next from NARF_Jam"
               Next
               mapinfo.do "commit table narf_jam"
               For i = 1 To j - 1
                   SortTemp = Bcch_Serv(1)
                   TempNum = 1
                   For k = 1 To j - 1
                       If SortTemp < Bcch_Serv(k) Then
                          SortTemp = Bcch_Serv(k)
                          TempNum = k
                       End If
                   Next
                   SortBcch(j - i) = SortTemp
                   Bcch_Serv(TempNum) = ""
                   SortCi(j - i) = Ci_Serv(TempNum)
                   SortSource(j - i) = SourceBcch(TempNum)
               Next
               
            mapinfo.do "Add Map window Frontwindow() Layer NARF_Jam"
            
               my_msg = "shade window FrontWindow() NARF_Jam With bcch_serv values "
               My_Color = "10535167"
               For i = 1 To j - 1
                   If i = j - 1 Then
                      my_msg = my_msg + SortBcch(i) + " Symbol (39," + My_Color + ",10,""MapInfo Cartographic"",0,0) "
                      my_msg = my_msg + "default Symbol(39,0,10,""MapInfo Cartographic"",0,0)"
                   Else
                      my_msg = my_msg + SortBcch(i) + " Symbol (39," + My_Color + ",10,""MapInfo Cartographic"",0,0),"
                   End If
                   My_Color = Format(Val(My_Color) + 6000)
                   Ci_Name(i) = Findcell(SortCi(i))
                   finds = InStr(Ci_Name(i), Chr(0))
                   If finds > 0 Then
                      Ci_Name(i) = Trim(Left(Ci_Name(i), finds - 1))
                   End If
                   If Ci_Name(i) = "" Then
                      Ci_Name(i) = "    "
                   End If
               Next
               mapinfo.do my_msg
               my_msg = ""
               For i = 1 To j - 1
                   my_msg = my_msg + "," + Chr(34) + SortBcch(i) + " [" + Ci_Name(i) + "]" + space(11 - (Len(Ci_Name(i)) - 1) * 2) + SortSource(i) + Chr(34) + "display on"
               Next
               'my_msg = " Title " + Chr(34) + "上行同频干扰观测 OF " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "Ncell 对 BCCH 干扰" + Chr(34) + "Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""无服务小区"" display on" + my_msg
               
               If Option4.Value = True Then
                  my_msg = " Title " + Chr(34) + "Ncell 对 BCCH 邻频碰撞分布" + Chr(34) + "Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "[被碰撞小区] 碰撞来源（C/A<" & Trim(RxLevValue(1)) & "）标注：载干比" + Chr(34) + "Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off " + my_msg
               Else
                  my_msg = " Title " + Chr(34) + "Ncell 对 BCCH 邻频碰撞分布" + Chr(34) + "Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "[被碰撞小区] 碰撞来源（C/A<" & Trim(RxLevValue(1)) & "）标注：载干比" + Chr(34) + "Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off " + my_msg
               End If
               mapinfo.do "set legend window FrontWindow() Layer prev " & my_msg
               mapinfo.do "set map redraw off"
               'mapinfo.do "Set Map Layer NARF_Jam Label Visibility Font (""Arial"",257,8,16711680,16777215) With c_i Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
               mapinfo.do "Set Map Layer NARF_Jam Label Visibility Font (""Arial"",257,8,16711680,16777215) With c_i Auto On Overlap Off Duplicates On Position Below Auto On Offset 2"
               mapinfo.do "set map redraw on"
               
            Else
               j = 1
               For i = 1 To MyRow
                   If Val(mapinfo.eval("NARF_Jam.num_dch")) > 0 Then
                      If j = 1 Then
                         Bcch_Serv(j) = mapinfo.eval("NARF_Jam.num_dch")
                         Ci_Serv(j) = mapinfo.eval("NARF_Jam.ci_serv")
                         For k = 1 To 6
                             If Abs(Val(mapinfo.eval("NARF_Jam.bcch_n" & Format(k))) - Val(Bcch_Serv(j))) = 1 Then
                                SourceBcch(j) = mapinfo.eval("NARF_Jam.bcch_n" & Format(k))
                                Exit For
                             End If
                         Next
                         j = j + 1
                      Else
                         If Bcch_Serv(j - 1) <> mapinfo.eval("NARF_Jam.num_dch") Then
                            For k = 1 To j - 1
                                If Bcch_Serv(k) = mapinfo.eval("NARF_Jam.num_dch") Then
                                   GoTo Next_Point4
                                End If
                            Next
                            Bcch_Serv(j) = mapinfo.eval("NARF_Jam.num_dch")
                            Ci_Serv(j) = mapinfo.eval("NARF_Jam.ci_serv")
                            For k = 1 To 6
                                If Abs(Val(mapinfo.eval("NARF_Jam.bcch_n" & Format(k))) - Val(Bcch_Serv(j))) = 1 Then
                                   SourceBcch(j) = mapinfo.eval("NARF_Jam.bcch_n" & Format(k))
                                   Exit For
                                End If
                            Next
                            j = j + 1
                         End If
                      End If
                   End If
Next_Point4:
                            For k = 1 To 6
                                If Option4.Value Then
                                   If Abs(Val(mapinfo.eval("NARF_Jam.bcch_n" & Format(k))) - Val(mapinfo.eval("NARF_Jam.num_dch"))) = 1 And Val(mapinfo.eval("NARF_Jam.rxlev_f")) - Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))) < Val(RxLevValue(1)) Then
                                      'MyRxlev = Format(Abs(Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))) - Val(mapinfo.eval("NARF_Jam.rxlev_f"))))
                                      MyRxlev = Format(Val(mapinfo.eval("NARF_Jam.rxlev_f")) - Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))))
                                      mapinfo.do "UPDATE NARF_Jam set c_i = " & MyRxlev & " where rowid = " & Format(i)
                                      Exit For
                                   End If
                                Else
                                   If Abs(Val(mapinfo.eval("NARF_Jam.bcch_n" & Format(k))) - Val(mapinfo.eval("NARF_Jam.num_dch"))) = 1 And Val(mapinfo.eval("NARF_Jam.rxlev_s")) - Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))) < Val(RxLevValue(1)) Then
                                      'MyRxlev = Format(Abs(Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))) - Val(mapinfo.eval("NARF_Jam.rxlev_s"))))
                                      MyRxlev = Format(Val(mapinfo.eval("NARF_Jam.rxlev_s")) - Val(mapinfo.eval("NARF_Jam.rxlev_n" & Format(k))))
                                      mapinfo.do "UPDATE NARF_Jam set c_i = " & MyRxlev & " where rowid = " & Format(i)
                                      Exit For
                                   End If
                                End If
                            Next

                   mapinfo.do "fetch next from NARF_Jam"
               Next
            
            mapinfo.do "Add Map window Frontwindow() Layer NARF_Jam"
               mapinfo.do "commit table narf_jam"
               For i = 1 To j - 1
                   SortTemp = Bcch_Serv(1)
                   TempNum = 1
                   For k = 1 To j - 1
                       If SortTemp < Bcch_Serv(k) Then
                          SortTemp = Bcch_Serv(k)
                          TempNum = k
                       End If
                   Next
                   SortBcch(j - i) = SortTemp
                   Bcch_Serv(TempNum) = ""
                   SortCi(j - i) = Ci_Serv(TempNum)
                   SortSource(j - i) = SourceBcch(TempNum)
               Next
            
               my_msg = "shade window FrontWindow() NARF_Jam With val(num_dch) values "
               My_Color = "10535167"
               For i = 1 To j - 1
                   If i = j - 1 Then
                      my_msg = my_msg + SortBcch(i) + " Symbol (39," + My_Color + ",10,""MapInfo Cartographic"",0,0) "
                      my_msg = my_msg + "default Symbol(39,0,10,""MapInfo Cartographic"",0,0)"
                   Else
                      my_msg = my_msg + SortBcch(i) + " Symbol (39," + My_Color + ",10,""MapInfo Cartographic"",0,0),"
                   End If
                   My_Color = Format(Val(My_Color) + 6000)
                   Ci_Name(i) = Findcell(SortCi(i))
                   finds = InStr(Ci_Name(i), Chr(0))
                   If finds > 0 Then
                      Ci_Name(i) = Trim(Left(Ci_Name(i), finds - 1))
                   End If
                   If Ci_Name(i) = "" Then
                      Ci_Name(i) = "    "
                   End If
               Next
               mapinfo.do my_msg
               my_msg = ""
               For i = 1 To j - 1
                   my_msg = my_msg + "," + Chr(34) + SortBcch(i) + " [" + Ci_Name(i) + "]" + space(11 - (Len(Ci_Name(i)) - 1) * 2) + SortSource(i) + Chr(34) + "display on"
               Next
               If Option4.Value = True Then
                  my_msg = " Title " + Chr(34) + "Ncell 对 TCH 邻频碰撞分布" + Chr(34) + "Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "[被碰撞小区] 碰撞来源（C/A<" & Trim(RxLevValue(1)) & "）标注：载干比" + Chr(34) + "Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off " + my_msg
               Else
                  my_msg = " Title " + Chr(34) + "Ncell 对 TCH 邻频碰撞分布" + Chr(34) + "Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "[被碰撞小区] 碰撞来源（C/A<" & Trim(RxLevValue(1)) & "）标注：载干比" + Chr(34) + "Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off " + my_msg
               End If
               mapinfo.do "set legend window FrontWindow() Layer prev " & my_msg
               mapinfo.do "set map redraw off"
               'mapinfo.do "Set Map Layer NARF_Jam Label Visibility Font (""Arial"",257,8,16711680,16777215) With c_i Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
               mapinfo.do "Set Map Layer NARF_Jam Label Visibility Font (""Arial"",257,8,16711680,16777215) With c_i Auto On Overlap Off Duplicates On Position Below Auto On Offset 2"
               mapinfo.do "set map redraw on"
            
            End If
         End If
Non_2:
         If No_PZ1 And No_PZ2 Then
            If Menu_Flag = 442 Then
               MsgBox "该路段不存在Ncell对TCH的同、邻频碰撞", 64, "提示"
            Else
               MsgBox "该路段不存在Ncell对BCCH的同、邻频碰撞", 64, "提示"
            End If
         ElseIf No_PZ1 Then
            If Menu_Flag = 442 Then
               MsgBox "该路段不存在Ncell对TCH的同频碰撞", 64, "提示"
            Else
               MsgBox "该路段不存在Ncell对BCCH的同频碰撞", 64, "提示"
            End If
         ElseIf No_PZ2 Then
            If Menu_Flag = 442 Then
               MsgBox "该路段不存在Ncell对TCH的邻频碰撞", 64, "提示"
            Else
               MsgBox "该路段不存在Ncell对BCCH的邻频碰撞", 64, "提示"
            End If
         End If
         Unload Me
End Sub
