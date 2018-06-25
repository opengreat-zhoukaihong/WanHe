VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form DocManager 
   BackColor       =   &H00C0C0C0&
   Caption         =   "文档管理"
   ClientHeight    =   4935
   ClientLeft      =   3270
   ClientTop       =   1695
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Doc_mana.frx":0000
   LinkTopic       =   "DOC_MANA"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4935
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2490
      Left            =   240
      TabIndex        =   17
      Top             =   150
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   4392
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "测试目的"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "测试背景"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "测试对象"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "测试天气"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   1980
      Left            =   330
      TabIndex        =   16
      Top             =   540
      Width           =   3975
      Begin VB.OptionButton TWETHER 
         Caption         =   "夜晚"
         Height          =   240
         Index           =   2
         Left            =   390
         TabIndex        =   39
         Top             =   1275
         Width           =   795
      End
      Begin VB.OptionButton TWETHER 
         Caption         =   "阴雨"
         Height          =   240
         Index           =   1
         Left            =   390
         TabIndex        =   38
         Top             =   900
         Width           =   795
      End
      Begin VB.OptionButton TWETHER 
         Caption         =   "晴好"
         Height          =   240
         Index           =   0
         Left            =   390
         TabIndex        =   37
         Top             =   510
         Width           =   795
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1980
      Left            =   330
      TabIndex        =   15
      Top             =   540
      Width           =   3975
      Begin VB.OptionButton TOBJ 
         Caption         =   "室内"
         Height          =   240
         Index           =   3
         Left            =   390
         TabIndex        =   36
         Top             =   1560
         Width           =   765
      End
      Begin VB.OptionButton TOBJ 
         Caption         =   "路段"
         Height          =   240
         Index           =   2
         Left            =   390
         TabIndex        =   35
         Top             =   1185
         Width           =   765
      End
      Begin VB.OptionButton TOBJ 
         Caption         =   "基站"
         Height          =   240
         Index           =   1
         Left            =   390
         TabIndex        =   34
         Top             =   795
         Width           =   765
      End
      Begin VB.OptionButton TOBJ 
         Caption         =   "区域"
         Height          =   240
         Index           =   0
         Left            =   390
         TabIndex        =   33
         Top             =   405
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1980
      Left            =   330
      TabIndex        =   18
      Top             =   540
      Width           =   3975
      Begin VB.OptionButton TDIST 
         Caption         =   "客户投诉"
         Height          =   240
         Index           =   7
         Left            =   2280
         TabIndex        =   32
         Top             =   1545
         Width           =   1050
      End
      Begin VB.OptionButton TDIST 
         Caption         =   "异网评估"
         Height          =   240
         Index           =   6
         Left            =   2280
         TabIndex        =   31
         Top             =   1185
         Width           =   1080
      End
      Begin VB.OptionButton TDIST 
         Caption         =   "网络评估"
         Height          =   240
         Index           =   5
         Left            =   2280
         TabIndex        =   30
         Top             =   810
         Width           =   1110
      End
      Begin VB.OptionButton TDIST 
         Caption         =   "干扰调查"
         Height          =   240
         Index           =   4
         Left            =   2280
         TabIndex        =   29
         Top             =   420
         Width           =   1125
      End
      Begin VB.OptionButton TDIST 
         Caption         =   "话务量调整"
         Height          =   240
         Index           =   3
         Left            =   390
         TabIndex        =   28
         Top             =   1545
         Width           =   1200
      End
      Begin VB.OptionButton TDIST 
         Caption         =   "切换带调整"
         Height          =   240
         Index           =   2
         Left            =   390
         TabIndex        =   27
         Top             =   1170
         Width           =   1200
      End
      Begin VB.OptionButton TDIST 
         Caption         =   "覆盖调查"
         Height          =   240
         Index           =   1
         Left            =   390
         TabIndex        =   26
         Top             =   780
         Width           =   1050
      End
      Begin VB.OptionButton TDIST 
         Caption         =   "一般调查"
         Height          =   240
         Index           =   0
         Left            =   390
         TabIndex        =   25
         Top             =   405
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1980
      Left            =   330
      TabIndex        =   14
      Top             =   540
      Width           =   3975
      Begin VB.OptionButton TBACK 
         Caption         =   "气侯影响"
         Height          =   240
         Index           =   5
         Left            =   2280
         TabIndex        =   24
         Top             =   1275
         Width           =   1050
      End
      Begin VB.OptionButton TBACK 
         Caption         =   "上级普查"
         Height          =   240
         Index           =   4
         Left            =   2280
         TabIndex        =   23
         Top             =   870
         Width           =   1065
      End
      Begin VB.OptionButton TBACK 
         Caption         =   "扩容工程"
         Height          =   240
         Index           =   3
         Left            =   2280
         TabIndex        =   22
         Top             =   465
         Width           =   1065
      End
      Begin VB.OptionButton TBACK 
         Caption         =   "小区参数验证"
         Height          =   240
         Index           =   2
         Left            =   375
         TabIndex        =   21
         Top             =   1260
         Width           =   1395
      End
      Begin VB.OptionButton TBACK 
         Caption         =   "天线馈线调整"
         Height          =   240
         Index           =   1
         Left            =   375
         TabIndex        =   20
         Top             =   870
         Width           =   1395
      End
      Begin VB.OptionButton TBACK 
         Caption         =   "射频功率调整"
         Height          =   240
         Index           =   0
         Left            =   375
         TabIndex        =   19
         Top             =   465
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2190
      Left            =   240
      TabIndex        =   2
      Top             =   2655
      Width           =   4170
      Begin VB.ComboBox TESTDOC 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1005
         TabIndex        =   13
         Text            =   "DocName"
         Top             =   1740
         Width           =   3000
      End
      Begin VB.TextBox GPS 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1020
         TabIndex        =   11
         Text            =   "3"
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox TESTIMG 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   270
         Left            =   1035
         TabIndex        =   10
         Text            =   "无"
         Top             =   1020
         Width           =   1575
      End
      Begin VB.TextBox TESTDATE 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1035
         TabIndex        =   6
         Text            =   "1997.01.23"
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox Partner 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1035
         TabIndex        =   4
         Text            =   "无网不通"
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "文件名:"
         Height          =   180
         Left            =   315
         TabIndex        =   12
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "位"
         Height          =   180
         Left            =   1470
         TabIndex        =   9
         Top             =   1425
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GPS精度:"
         Height          =   180
         Left            =   225
         TabIndex        =   8
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "现场图象:"
         Height          =   180
         Left            =   150
         TabIndex        =   7
         Top             =   1065
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "测试日期:"
         Height          =   180
         Left            =   165
         TabIndex        =   5
         Top             =   705
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "测试人员:"
         Height          =   180
         Left            =   165
         TabIndex        =   3
         Top             =   315
         Width           =   810
      End
   End
   Begin VB.CommandButton SSCommand2 
      Caption         =   "取消"
      Height          =   320
      Left            =   4650
      TabIndex        =   1
      Top             =   900
      Width           =   1080
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "确定"
      Height          =   320
      Left            =   4650
      TabIndex        =   0
      Top             =   465
      Width           =   1080
   End
End
Attribute VB_Name = "DocManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doc_rec  As doc
Dim Current_Page As Integer

Private Sub TabStrip1_Click()
    On Error Resume Next
    If TabStrip1.SelectedItem.Index = Current_Page Then
       Exit Sub
    End If
    Select Case TabStrip1.SelectedItem.Index
       Case 1
            Frame2.Visible = False
            Frame4.Visible = False
            Frame5.Visible = False
            Frame3.Visible = True
            Frame3.ZOrder 0
            Current_Page = 1
       Case 2
            Frame3.Visible = False
            Frame4.Visible = False
            Frame5.Visible = False
            Frame2.Visible = True
            Frame2.ZOrder 0
            Current_Page = 2
       Case 3
            Frame2.Visible = False
            Frame3.Visible = False
            Frame5.Visible = False
            Frame4.Visible = True
            Frame4.ZOrder 0
            Current_Page = 3
       Case 4
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = False
            Frame5.Visible = True
            Frame5.ZOrder 0
            Current_Page = 4
    End Select
End Sub

Private Sub TESTDOC_Click()
        On Error Resume Next
        Seek 1#, 1
        p = 1
        Get #1, p, doc_rec
        msg = Trim(TESTDOC.Text)
        Do While Not EOF(1) And UCase(msg) <> UCase(Trim(doc_rec.DOCNAME))
           Get #1, p, doc_rec
           If EOF(1) Then
              Exit Do
           End If
           p = p + 74
           Seek #1, p
        Loop

        If Not EOF(1) Then
           TESTDOC.Text = doc_rec.DOCNAME
           GPS.Text = doc_rec.GPS
           TESTDATE.Text = doc_rec.DATE
           Partner.Text = doc_rec.Partner
           TESTIMG.Text = doc_rec.IMG

           For i = 0 To 3
               TOBJ(i) = False
           Next i
           TOBJ(Val(doc_rec.TESTOBJECT)) = True
           For i = 0 To 7
               TDIST(i) = False
           Next i
           TDIST(Val(doc_rec.TESTDIST)) = True
           For i = 0 To 5
               TBACK(i) = False
           Next i
           TBACK(Val(doc_rec.TESTBACK)) = True
           For i = 0 To 2
               TWETHER(i) = False
           Next i
           TWETHER(Val(doc_rec.WEATHER)) = True

        End If
End Sub

Private Sub Form_Load()
    On Error GoTo Go_OUT
    Gsm_FileName = Gsm_Path + "\doc_man.dat"
    Open Gsm_FileName For Binary Shared As #1
    On Error Resume Next
    Frame2.Visible = False
    Frame4.Visible = False
    Frame5.Visible = False
    Frame3.Visible = True
    Frame3.ZOrder 0
    Current_Page = 1
    p = 1
    
  Select Case Menu_Flag
   Case 121, 123, 1244, 2222, 4444, 4449, 128
        TESTDOC.Text = sinput
        Get #1, p, doc_rec
        Do While Not EOF(1)
           Get #1, p, doc_rec
           If UCase(Trim(doc_rec.DOCNAME)) = UCase(sinput) Then
              TESTDOC.Text = doc_rec.DOCNAME
              GPS.Text = doc_rec.GPS
              TESTDATE.Text = doc_rec.DATE
              Partner.Text = doc_rec.Partner
              TESTIMG.Text = doc_rec.IMG

              TOBJ(Val(doc_rec.TESTOBJECT)) = True
              TDIST(Val(doc_rec.TESTDIST)) = True
              TBACK(Val(doc_rec.TESTBACK)) = True
              TWETHER(Val(doc_rec.WEATHER)) = True
           End If
           TESTDOC.AddItem doc_rec.DOCNAME
           p = p + 74
           If EOF(1) Then
              Exit Do
           End If
           Seek #1, p
        Loop
        TESTDATE.Text = Format(year(Now)) & "." & Format(month(Now), "00") & "." & Format(day(Now), "00")
   Case 122
           Get #1, p, doc_rec
           TESTDOC.Text = doc_rec.DOCNAME
           GPS.Text = doc_rec.GPS
           TESTDATE.Text = doc_rec.DATE
           Partner.Text = doc_rec.Partner
           TESTIMG.Text = doc_rec.IMG

           TOBJ(Val(doc_rec.TESTOBJECT)) = True
           TDIST(Val(doc_rec.TESTDIST)) = True
           TBACK(Val(doc_rec.TESTBACK)) = True
           TWETHER(Val(doc_rec.WEATHER)) = True
           Get #1, p, doc_rec
           Do While Not EOF(1)
              Get #1, p, doc_rec
              TESTDOC.AddItem doc_rec.DOCNAME
              p = p + 74
              If EOF(1) Then
                 Exit Do
              End If
              Seek #1, p
           Loop

  End Select
Go_OUT:
End Sub

Private Sub Form_Unload(Cancel As Integer)
     On Error Resume Next
     Close #1
End Sub

Private Sub SSCommand1_Click()
    Dim Is_Exist As Boolean
    Dim pot, pot1, i As Integer, j As Integer, k As Integer
    Dim str, stab, ttab, my_tab As String
    Dim Convert_Flag As Boolean
    Dim finds As Integer
    Dim Tab_Rows As Variant
    
    On Error Resume Next
    GPS_NO = Val(GPS.Text)
    DocManager.Hide
    Select Case Menu_Flag
        Case 121, 123, 2222, 4444, 4449, 128
             Close
             If (Menu_Flag = 128 And Data_Report = True) Or (Menu_Flag = 121 And Data_Report = True) Or Menu_Flag = 4444 Then
                Stre_Sel.Show 1
             End If
             Screen.MousePointer = 11
            'If sys = 1 Then
            '   tran_fn = 1
            '   tran_f(1) = sinput
            'End If
            For j = 1 To tran_fn
                sinput = tran_f(j)
                msg = sinput
                Is_Exist = False
                Gsm_FileName = Gsm_Path + "\doc_man.dat"
                Open Gsm_FileName For Binary Shared As #1
                p = 1
                Get #1, p, doc_rec
                Do While Not EOF(1)
                   If UCase(msg) = UCase(Trim(doc_rec.DOCNAME)) Then
                      Is_Exist = True
                      Exit Do
                   End If
                   Get #1, p, doc_rec
                   If EOF(1) Then
                      Exit Do
                   End If
                   p = p + 74
                   Seek #1, p
                Loop
                doc_rec.DOCNAME = sinput
                doc_rec.GPS = GPS.Text
                doc_rec.DATE = TESTDATE.Text
                doc_rec.Partner = Partner.Text
                doc_rec.IMG = TESTIMG.Text
                For i = 0 To 3
                    If TOBJ(i) = True Then doc_rec.TESTOBJECT = i
                Next i
                For i = 0 To 7
                    If TDIST(i) = True Then doc_rec.TESTDIST = i
                Next i
                For i = 0 To 5
                    If TBACK(i) = True Then doc_rec.TESTBACK = i
                Next i
                For i = 0 To 2
                    If TWETHER(i) = True Then doc_rec.WEATHER = i
                Next i
                Gsm_FileName = Gsm_Path + "\doc_man.dat"
                If Is_Exist = False Then
                   p = FileLen(Gsm_FileName)
                Else
                   p = p - 1
                End If
                Put #1, p + 1, doc_rec
                Close
                pot = Len(sinput) - 4
                soutput = Left$(sinput, pot)
                soutput = soutput + ".DBF"
                stab = Left$(sinput, pot)
                If tran_del = 2 And sys = 0 Then
                   my_tab = my_tab + "f"
                   stab = stab + "f"
                Else
                   If tran_del = 3 And sys = 0 Then
                      my_tab = my_tab + "e"
                      stab = stab + "e"
                   End If
                End If
                stab = stab + ".tab"
                ttab = stab
                pot = 1
                While (pot <> 0)
                      ttab = Mid$(ttab, pot + 1, Len(ttab) - pot + 1)
                      pot = InStr(1, ttab, "\")
                Wend
                If Mid$(ttab, 1, 1) < "9" Then
                   my_tab = "_" + Mid$(ttab, 1, Len(ttab) - 4)
                Else
                   my_tab = Mid$(ttab, 1, Len(ttab) - 4)
                End If
             If Menu_Flag <> 4449 Then
                Screen.MousePointer = 11
                Data_Tran_Flag = 1
                Per_Show.Show 1
                If tran_del = 1 And sys = 0 Then
                   Data_Tran_Flag = 2
                   Per_Show.Show 1
                End If
             Else
                FileCopy sinput, soutput
             End If
                On Error Resume Next
                msg = "Register Table  " + " " + Chr(34) + soutput + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + stab + Chr(34)
                mapinfo.Do msg
                msg = "Open Table " + Chr(34) + stab + Chr(34)
                mapinfo.Do msg
                mapinfo.Do "Create Map For " & my_tab & " CoordSys Earth Projection 1, 0"
                mapinfo.Do "Set Style Symbol MakeSymbol(33,0,2)" '
                On Error GoTo MAP_OUT
                If Menu_Flag = 121 Or Menu_Flag = 124 Then
                   cc = Chr(34) + "&H" + Chr(34)
                   msg = "update " + my_tab + " set Obj= CreatePoint(Lon, Lat), CI_SERV = str$(val((" & cc & "+CI_SERV))),LAC_SERV = str$(val((" & cc & "+LAC_SERV)))"
                    ' ,RXLEV_F=110-RXLEV_F,RXLEV_S=110-RXLEV_S,RXLEV_N1=110-rxlev_n1,rxlev_n2=110-rxlev_n2,rxlev_n3=110-rxlev_n3,rxlev_n4=110-rxlev_n4,rxlev_n5=110-rxlev_n5,rxlev_n6=110-rxlev_n6 "
                Else
                   msg = "update " + my_tab + " set Obj= CreatePoint(Lon, Lat)"
                End If
                mapinfo.Do msg
                On Error Resume Next
                mapinfo.Do "commit table " & my_tab
                mapinfo.Do "close table " & my_tab
            Next
            Close #1
            Unload Me
            Screen.MousePointer = 0
            If (Menu_Flag = 128 And Data_Report = True) Or (Menu_Flag = 121 And Data_Report = True) Or Menu_Flag = 4444 Then
               i = 1
               j = 1
               For k = 1 To tran_fn
                   stre_tab(j) = tran_f(i)
                   finds = InStr(stre_tab(j), ".")
                   If finds > 0 Then
                      stre_tab(j) = Left(stre_tab(j), finds - 1)
                   End If
                   If tran_del = 2 Then
                      stre_tab(j) = stre_tab(j) + "f"
                   Else
                      If tran_del = 3 Then
                         stre_tab(j) = stre_tab(j) + "e"
                      End If
                   End If
                   stre_tab(j) = stre_tab(j) + ".tab"
                   If dir(stre_tab(j)) <> "" Then
                      mapinfo.Do "open table " + Chr(34) + stre_tab(j) + Chr(34)
                      finds = InStr(stre_tab(j), ".")
                      If finds > 0 Then
                         stre_tab(j) = Left(stre_tab(j), finds - 1)
                      End If
                      finds = InStr(stre_tab(j), "\")
                      Do While finds > 0
                         stre_tab(j) = Right(stre_tab(j), Len(stre_tab(j)) - finds)
                         finds = InStr(stre_tab(j), "\")
                      Loop
                      If Asc(Left(stre_tab(j), 1)) > 47 And Asc(Left(stre_tab(j), 1)) < 58 Then
                         stre_tab(j) = "_" + stre_tab(j)
                      End If
                      mapinfo.Do "fetch first from " & stre_tab(j)
                      Tab_Rows = mapinfo.eval("tableinfo(" + stre_tab(j) + ",8)")
                      If Tab_Rows > 0 Then
                         j = j + 1
                      Else
                         mapinfo.Do "close table " & stre_tab(j)
                      End If
                   End If
                   i = i + 1
               Next
               If j > 1 Then
                  stre_num = j - 1
                  mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\map\cell" + Chr(34)
                  'My_Report
               End If
            Else
               MsgBox "数据已生成，您可开始分析了", 64, "提示"
            End If
            Exit Sub

ERR_OUT:
            On Error Resume Next
            Screen.MousePointer = 0
            MsgBox "文件不在处理范畴，请查阅帮助内容", 64, "提示"
            Close #1
            Unload Me
            Exit Sub
MAP_OUT:
            On Error Resume Next
            Close #1
            Unload Me
'       mapinfo.do "close table """ & my_tab
            Screen.MousePointer = 0
            MsgBox "数据生成不完整，可能无磁盘空间或地图系统出错！请退出ANT，排除错误后重新转换本数据 。", 64, "提示"
            Unload Me
            Exit Sub
        Case 122
            msg = Trim(TESTDOC.Text)
            p = 1
            Get #1, p, doc_rec
            Do While Not EOF(1) And UCase(msg) <> UCase(Trim(doc_rec.DOCNAME))
                  p = p + 74
                  Seek #1, p
                  Get #1, p, doc_rec
                  If EOF(1) Then
                     Exit Do
                  End If
            Loop
            If Not EOF(1) Then
               doc_rec.GPS = GPS.Text
               doc_rec.DATE = TESTDATE.Text
               doc_rec.Partner = Partner.Text
               doc_rec.IMG = TESTIMG.Text
               For i = 0 To 3
                   If TOBJ(i) = True Then doc_rec.TESTOBJECT = i
               Next i
               For i = 0 To 7
                   If TDIST(i) = True Then doc_rec.TESTDIST = i
               Next i
               For i = 0 To 5
                   If TBACK(i) = True Then doc_rec.TESTBACK = i
               Next i
               For i = 0 To 2
                   If TWETHER(i) = True Then doc_rec.WEATHER = i
               Next i
               Put #1, p, doc_rec
            End If
            Close
            Unload Me
        Case 1244
            Dim txt_temp As String, tab_temp As String, othername As String, tab_name As String, dbf_name As String, use_name As String
            Dim findme As Integer
            Dim msg1 As String, msg2 As String
            Dim col_no As Integer
            On Error Resume Next
            Screen.MousePointer = 11
            Close
            Convert_Flag = False
            For j = 1 To tran_fn
                othername = tran_f(j)
                txt_temp = ""
                Gsm_FileName = Gsm_Path + "\doc_man.dat     "
                Open Gsm_FileName For Binary Shared As #1
                p = 1
                Get #1, p, doc_rec
                Do While Not EOF(1) And UCase(Trim(msg)) <> UCase(Trim(doc_rec.DOCNAME))
                      Get #1, p, doc_rec
                      If EOF(1) Then
                         Exit Do
                      End If
                      p = p + 74
                      Seek #1, p
                Loop
                If EOF(1) Then
                   doc_rec.DOCNAME = sinput
                   Gsm_FileName = Gsm_Path + "\doc_man.dat"
                   p = FileLen(Gsm_FileName)
                   doc_rec.GPS = GPS.Text
                   doc_rec.DATE = TESTDATE.Text
                   doc_rec.Partner = Partner.Text
                   doc_rec.IMG = TESTIMG.Text
                   For i = 0 To 3
                       If TOBJ(i) = True Then doc_rec.TESTOBJECT = i
                   Next i
                   For i = 0 To 7
                       If TDIST(i) = True Then doc_rec.TESTDIST = i
                   Next i
                   For i = 0 To 5
                       If TBACK(i) = True Then doc_rec.TESTBACK = i
                   Next i
                   For i = 0 To 2
                       If TWETHER(i) = True Then doc_rec.WEATHER = i
                   Next i
                   Put #1, p + 1, doc_rec
                End If
                Close
                Do
                    findme = InStr(othername, "\")
                    If findme > 0 Then
                       txt_temp = txt_temp + Left(othername, findme)
                       othername = Right(othername, Len(othername) - findme)
                    Else
                       tab_temp = txt_temp + "temp.tab"
                       txt_temp = txt_temp + "temp.txt"
                       Exit Do
                    End If
                Loop
                findme = InStr(tran_f(j), ".")
                If findme = 0 Then
                   tab_name = obt_name + ".tab"
                   dbf_name = obt_name + ".dbf"
                Else
                   dbf_name = Left(tran_f(j), findme - 1)
                   tab_name = dbf_name + ".tab"
                   dbf_name = dbf_name + ".dbf"
                End If
                If dir(txt_temp) <> "" Then
                   Kill txt_temp
                End If
                If dir(dbf_name) <> "" Then
                   Kill dbf_name
                End If
                If dir(tab_name) <> "" Then
                   Kill tab_name
                End If
                FileCopy tran_f(j), txt_temp
                Gsm_FileName = Gsm_Path + "\data_tem.dbf"
                Gsm_File2 = Gsm_Path + "\data_tem.tab"
                FileCopy Gsm_FileName, dbf_name
                FileCopy Gsm_File2, tab_name
                use_name = tran_f(j)
                Do
                    findme = InStr(use_name, "\")
                    If findme > 0 Then
                       use_name = Right(use_name, Len(use_name) - findme)
                    Else
                       findme = InStr(use_name, ".")
                       If findme > 0 Then
                          use_name = Left(use_name, findme - 1)
                       End If
                       Exit Do
                    End If
                Loop
                mapinfo.Do "Register Table " + Chr(34) + txt_temp + Chr(34) + " TYPE ASCII Delimiter 9 Titles Charset " + Chr(34) + "CodePage437" + Chr(34) + " Into " + Chr(34) + tab_temp + Chr(34)
                mapinfo.Do "open table " + Chr(34) + tab_temp + Chr(34)
                mapinfo.Do "fetch first from temp"
                col_no = mapinfo.eval("tableinfo(temp,4)")
                If col_no <> 56 Then
                   MsgBox "Export 格式选择错误！" + Chr(10) + "请采用 Actix layer message ASCII 格式", 64, "提示"
                   mapinfo.Do "close table temp"
                   GoTo next_file
                End If
                Convert_Flag = True
                mapinfo.Do "open table " + Chr(34) + tab_name + Chr(34)
                mapinfo.Do "fetch first from " + use_name
                msg1 = " (col1,col3,col4,col5,col13,col14,col15,col16,col17,col18,col19,col20,col21,col22,col23,col24,col25,col26,col27,col28,col29,col40,col41,col42,col43,col44,col45,col46,col47,col48,col49,col50,col51,col52,col53,col54,col55,col56,col57,col58)"
                msg2 = " col2,col3,col4,col5,col49,col46,col47,col45,col23,col16,col42,col43,col44,col6,col14,col13,col15,col50,col31,col34,col33,col32,col17,col7,col24,col18,col8,col25,col19,col9,col26,col20,col10,col27,col21,col11,col28,col22,col12,col29"
                mapinfo.Do "insert into " + use_name + msg1 + " select " + msg2 + " from temp"
                mapinfo.Do "close table temp"
                mapinfo.Do "fetch first from " + use_name
                cc = Chr(34) + "&H" + Chr(34)
                mapinfo.Do "Create Map For  " & use_name & "  CoordSys Earth Projection 1, 0"
                mapinfo.Do "Set Style Symbol MakeSymbol(33,0,2)" '
                msg = "update " + use_name + " set Obj= CreatePoint(Lon, Lat), CI_SERV = str$(val((" & cc & "+CI_SERV)))"
                msg = msg + ", rxlev_f=110+rxlev_f,RXLEV_S=110+RXLEV_S,RXLEV_N1=110+rxlev_n1,rxlev_n2=110+rxlev_n2,rxlev_n3=110+rxlev_n3,rxlev_n4=110+rxlev_n4,rxlev_n5=110+rxlev_n5,rxlev_n6=110+rxlev_n6"
                mapinfo.Do msg
                row = Val(mapinfo.eval("tableinfo(" & use_name & ",8)"))
                mapinfo.Do "fetch first from " & use_name
                i = 1
                str = use_name + ".message"
                While i <= row
                      msg = mapinfo.eval(str)
                      If Left(msg, 2) = "L3" Then
                         mapinfo.Do "update  " + use_name + "  set message=mid$(message,10,len(message)-9) where ROWID=" & i
                      End If
                      mapinfo.Do "fetch next from " & use_name
                      i = i + 1
                Wend
                mapinfo.Do "commit table " + use_name
                mapinfo.Do "close table " + use_name
                Kill tab_temp
                Kill txt_temp
next_file:
            Next j
            Screen.MousePointer = 0
            If Convert_Flag = True Then
               MsgBox "数据已生成，您可开始分析了", 64, "提示"
            End If
    End Select
OUT_1:
    Screen.MousePointer = 0
    Close
    Unload Me

End Sub

Private Sub SSCommand2_Click()
    On Error Resume Next
    Unload DocManager
End Sub
