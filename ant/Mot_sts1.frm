VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Mot_Sts1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数据转换"
   ClientHeight    =   1380
   ClientLeft      =   795
   ClientTop       =   6270
   ClientWidth     =   4770
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mot_sts1.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1380
   ScaleWidth      =   4770
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   320
      Left            =   1770
      TabIndex        =   2
      Top             =   990
      Width           =   1080
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   285
      TabIndex        =   1
      Top             =   600
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   3660
      Top             =   30
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "正在转换 "
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   285
      Width           =   810
   End
End
Attribute VB_Name = "Mot_Sts1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim convert_num As Integer
Dim hcol(1 To 33) As String * 6
Dim n As Integer
Dim a(1 To 10) As String
Dim b(1 To 10) As String
Dim c(1 To 10) As String
Dim pn(1 To 10) As Integer
Dim colnum(1 To 10) As Integer
Dim MoreFile As Boolean

Sub fa_line(ltx As String)       '读原始文件抬头
    Dim tmp As String
    Dim i As Integer
    Dim finds As Integer
    
    On Error Resume Next
    i = 1
    Do
       finds = InStr(ltx, " ")
       If finds = 0 Then
          a(i) = ltx
          a(i) = UCase(a(i))
          If a(i) = "CELL" Or a(i) = "ASSIGN." Or a(i) = "TIME" Or a(i) = "BLOCKED" Then
             a(i) = ""
          End If
          Exit Do
       Else
          a(i) = Left(ltx, finds - 1)
          a(i) = UCase(a(i))
          ltx = Right(ltx, Len(ltx) - finds)
          If a(i) = "SDCCH" Or a(i) = "TCH" Then
             If Len(ltx) >= 6 Then
                tmp = Mid(ltx, 1, 6)
                If UCase(tmp) = "RFLOSS" Then
                   finds = InStr(ltx, " ")
                   If finds = 0 Then
                      Exit Do
                   Else
                      ltx = Right(ltx, Len(ltx) - finds)
                   End If
                End If
             End If
          Else
             If a(i) = "SUCCESS" Or a(i) = "FAILURE" Then
                If Len(ltx) >= 4 Then
                   tmp = Mid(ltx, 1, 4)
                   If UCase(tmp) = "RATE" Then
                      finds = InStr(ltx, " ")
                      If finds = 0 Then
                         Exit Do
                      Else
                         ltx = Right(ltx, Len(ltx) - finds)
                      End If
                   End If
                End If
             Else
                If a(i) = "CELL" Or a(i) = "ASSIGN." Or a(i) = "TIME" Or a(i) = "BLOCKED" Then
                   a(i) = ""
                   i = i - 1
                End If
             End If
          End If
          If Len(ltx) > 0 Then
             ltx = LTrim$(ltx)
             If Len(ltx) = 0 Then
                Exit Do
             End If
          Else
             Exit Do
          End If
       End If
       i = i + 1
       If i > 10 Then
          Exit Do
       End If
    Loop
End Sub

Sub fb_line(ltx As String)       '读原始文件抬头
    Dim tmp As String
    Dim i As Integer
    Dim finds As Integer
    
    On Error Resume Next
    i = 1
    Do
       finds = InStr(ltx, " ")
       If finds = 0 Then
          b(i) = ltx
          b(i) = UCase(b(i))
          If b(i) = "CELL" Or b(i) = "ASSIGN." Or b(i) = "TIME" Or b(i) = "BLOCKED" Then
             b(i) = ""
          End If
          Exit Do
       Else
          b(i) = Left(ltx, finds - 1)
          b(i) = UCase(b(i))
          ltx = Right(ltx, Len(ltx) - finds)
          If b(i) = "SDCCH" Or b(i) = "TCH" Then
             If Len(ltx) >= 6 Then
                tmp = Mid(ltx, 1, 6)
                If UCase(tmp) = "RFLOSS" Then
                   finds = InStr(ltx, " ")
                   If finds = 0 Then
                      Exit Do
                   Else
                      ltx = Right(ltx, Len(ltx) - finds)
                   End If
                End If
             End If
          Else
             If b(i) = "SUCCESS" Or b(i) = "FAILURE" Then
                If Len(ltx) >= 4 Then
                   tmp = Mid(ltx, 1, 4)
                   If UCase(tmp) = "RATE" Then
                      finds = InStr(ltx, " ")
                      If finds = 0 Then
                         Exit Do
                      Else
                         ltx = Right(ltx, Len(ltx) - finds)
                      End If
                   End If
                End If
             Else
                If b(i) = "CELL" Or b(i) = "ASSIGN." Or b(i) = "TIME" Or b(i) = "BLOCKED" Then
                   b(i) = ""
                   i = i - 1
                End If
             End If
          End If
          If Len(ltx) > 0 Then
             ltx = LTrim$(ltx)
             If Len(ltx) = 0 Then
                Exit Do
             End If
          Else
             Exit Do
          End If
       End If
       i = i + 1
       If i > 10 Then
          Exit Do
       End If
    Loop
End Sub

Sub fc_line(ltx As String)       '读原始文件抬头
    Dim tmp As String
    Dim i As Integer
    Dim finds As Integer
    
    On Error Resume Next
    i = 1
    Do
       finds = InStr(ltx, " ")
       If finds = 0 Then
          c(i) = ltx
          c(i) = UCase(c(i))
          If c(i) = "CELL" Or c(i) = "ASSIGN." Or c(i) = "TIME" Or c(i) = "BLOCKED" Then
             c(i) = ""
          End If
          Exit Do
       Else
          c(i) = Left(ltx, finds - 1)
          c(i) = UCase(c(i))
          ltx = Right(ltx, Len(ltx) - finds)
          If c(i) = "SDCCH" Or c(i) = "TCH" Then
             If Len(ltx) >= 6 Then
                tmp = Mid(ltx, 1, 6)
                If UCase(tmp) = "RFLOSS" Then
                   finds = InStr(ltx, " ")
                   If finds = 0 Then
                      Exit Do
                   Else
                      ltx = Right(ltx, Len(ltx) - finds)
                   End If
                End If
             End If
          Else
             If c(i) = "SUCCESS" Or c(i) = "FAILURE" Then
                If Len(ltx) >= 4 Then
                   tmp = Mid(ltx, 1, 4)
                   If UCase(tmp) = "RATE" Then
                      finds = InStr(ltx, " ")
                      If finds = 0 Then
                         Exit Do
                      Else
                         ltx = Right(ltx, Len(ltx) - finds)
                      End If
                   End If
                End If
             Else
                If c(i) = "CELL" Or c(i) = "ASSIGN." Or c(i) = "TIME" Or c(i) = "BLOCKED" Then
                   c(i) = ""
                   i = i - 1
                End If
             End If
          End If
          If Len(ltx) > 0 Then
             ltx = LTrim$(ltx)
             If Len(ltx) = 0 Then
                Exit Do
             End If
          Else
             Exit Do
          End If
       End If
       i = i + 1
       If i > 10 Then
          Exit Do
       End If
    Loop
End Sub

Sub ch_line(ltx As String, endfra As Integer, c_f As String, ByVal blog As Boolean)   '读一行数据
    Dim buffer As String * 1
    Static log As Long
    
    On Error Resume Next
    If blog = True Then
       log = 0
       Exit Sub
    End If
    
    ltx = ""
    Do While Not EOF(1)
       endfra = 0
       Get #1, , buffer
       log = log + 1
       If buffer = Chr(10) Then
          If log >= FileLen(c_f) Then
              endfra = 1
          End If
          ltx = Trim(ltx)
          If ltx <> "" Then
             Exit Do
          End If
       Else
          If buffer <> Chr(13) Then
             ltx = ltx + buffer
          End If
       End If
       If log >= FileLen(c_f) Then
          endfra = 1
          Exit Do
       End If
    Loop
    If EOF(1) Then           '1/4
       endfra = 1
    End If
End Sub

Sub make_pos()   '确定抬头位置
    Dim i As Integer, j As Integer

    On Error Resume Next
    For i = 1 To n
        Select Case b(i)
            Case "RATE"
                 Select Case a(i)
                      Case "SDCCH"
                           pn(i) = 1
                           colnum(i) = 12
                      Case "TCH"
                           pn(i) = 2
                           colnum(i) = 29
                      Case "RFLOSS"
                           pn(i) = 3
                 End Select
            Case "ASSIGNMENTS"
                 pn(i) = 4
            Case "HOLDING"
                 If a(i) = "SDCCH" Then
                    pn(i) = 5
                    colnum(i) = 34
                 Else
                    pn(i) = 9
                    colnum(i) = 19
                 End If
            Case "ARRIVAL"
                 If a(i) = "SDCCH" Then
                    pn(i) = 6
                    colnum(i) = 8
                 Else
                    pn(i) = 10
                    colnum(i) = 18
                 End If
            Case "TRAFFIC"
                 If a(i) = "SDCCH" Then
                    pn(i) = 7
                    colnum(i) = 35
                 Else
                    pn(i) = 11
                    colnum(i) = 16
                 End If
            Case "CALLS", "CONGESTION"
                 Select Case a(i)
                      Case "SDCCH"
                          pn(i) = 8
                          colnum(i) = 9
                      Case "TCH"
                          pn(i) = 12
                          colnum(i) = 26
                      Case "TOTAL", "TOTAL_"
                          pn(i) = 25
                          colnum(i) = 25
                 End Select
            Case "_PER_R"
                 pn(i) = 13
                 colnum(i) = 3
            Case "_PROC_"
                 pn(i) = 50
            Case "SDCCH"
                 pn(i) = 20
                 colnum(i) = 6
            Case "SDCCH_"
                 pn(i) = 21
                 colnum(i) = 7
            Case "_FROM_"
                 pn(i) = 22
                 colnum(i) = 23
            Case "CELL_H"
                 pn(i) = 50
            Case "_TO_MS"
                 pn(i) = 24
                 colnum(i) = 24
            Case "SUCCESS"
                 If a(i) = "HANDOVER" Then
                    pn(i) = 27
                    colnum(i) = 33
                 Else
                    pn(i) = 29
                    colnum(i) = 27
                 End If
            Case "FAILURE"
                 pn(i) = 28
                 colnum(i) = 32
            Case "SES_SD"
                 pn(i) = 31
                 colnum(i) = 11
            Case "SES_TC"
                 pn(i) = 32
                 colnum(i) = 28
            Case "Q_TO_M"
                 pn(i) = 33
            Case "EFUSED"
                 pn(i) = 34
            Case "BSS_HO"
                 If a(i) = "INTER_" Then
                    pn(i) = 35
                    colnum(i) = 30
                 Else
                    pn(i) = 36
                    colnum(i) = 31
                 End If
        End Select
    Next
    j = 1
    For i = 1 To n
        Select Case c(i)
            Case "SUC_RA"
                Call get_50(j, 14)
                colnum(j - 1) = 4
            Case "CM_SER"
                Call get_50(j, 15)
                colnum(j - 1) = 20
            Case "CM_REE"
                Call get_50(j, 16)
                colnum(j - 1) = 22
            Case "PAGE_R"
                Call get_50(j, 17)
                colnum(j - 1) = 21
            Case "LOC_UP"
                Call get_50(j, 18)
            Case "IMSI_D"
                Call get_50(j, 19)
            Case "O_REQ"
                Call get_50(j, 23)
                colnum(j - 1) = 5
            Case "O"
                Call get_50(j, 26)
                colnum(j - 1) = 10
            Case "O_ATMP"
                Call get_50(j, 37)
                colnum(j - 1) = 17
        End Select
    Next
        
End Sub

Sub get_50(j As Integer, ByVal mm As Integer)

    On Error Resume Next
    Do While pn(j) <> 50
       j = j + 1
    Loop
    If j <= n Then
       pn(j) = mm
       j = j + 1
    End If
End Sub

Sub hua_ci(ltx As String, endfra As Integer, cc As String)   '取数据
    Dim ee As String
    Dim finds As Integer
    endfra = 0
    
    Do
       Call ch_line(ltx, endfra, cc, False)
       If InStr(UCase(ltx), "TIME  MODE:") > 0 Then
          endfra = 3
          MoreFile = True
          Exit Do
       End If
       If Len(ltx) > 5 Then
          ltx = LTrim$(ltx)
          If Len(ltx) > 5 Then
             ee = Mid(ltx, 1, 5)
             If ee = "460-0" Or ee = "460_0" Or ee = "460-1" Or ee = "460_1" Then
                If endfra = 1 Then
                   endfra = 2
                End If
                Exit Do
             End If
          End If
       End If
       If endfra = 1 Then
          Exit Do
       End If
    Loop
End Sub

Sub hua_time(ltx As String, endfra As Integer, cc As String)   '取数据
    Dim ee As String
    Dim finds As Integer
    Dim m1 As Integer
    Dim m2 As Integer

    On Error Resume Next
    endfra = 0
    Do
       Call ch_line(ltx, endfra, cc, False)
       If Len(ltx) > 5 Then
          ltx = LTrim$(ltx)
          If Len(ltx) > 5 Then
             finds = InStr(ltx, ":")
             If finds > 0 Then
                If Len(ltx) - finds > 2 Then
                   m1 = Asc(Mid(ltx, finds + 1, 1))
                   m2 = Asc(Mid(ltx, finds + 2, 1))
                   If m1 < 58 And m1 > 47 And m2 < 58 And m2 > 47 Then
                      If endfra = 1 Then
                         endfra = 2
                      End If
                      Exit Do
                   End If
                End If
             End If
          End If
       End If
       If endfra = 1 Then
          Exit Do
       End If
    Loop
End Sub

Sub get_da(ltx As String)       '赋记录值
    Dim finds As Integer
    Dim t As Integer
    
    On Error Resume Next
    'finds = InStr(ltx, " ")
    'ltx = Right(ltx, Len(ltx) - finds)
    'ltx = LTrim$(ltx)
    For t = 1 To n
        finds = InStr(ltx, " ")
        Select Case pn(t)
            Case 1, 2, 5 To 17, 20 To 29, 31, 32, 35 To 37
                 If finds = 0 Then
                    hcol(colnum(t) - 2) = ltx
                 Else
                    hcol(colnum(t) - 2) = Left(ltx, finds - 1)
                 End If
        End Select
        ltx = LTrim$(Right(ltx, Len(ltx) - finds))
    Next
End Sub

Sub fini(tmp As String, tr1 As String, tr2 As String, p As Boolean)  '某些字段值计算
    Dim t1 As String * 6, t2 As String * 6
    Dim tmp1 As String
    Dim tmp2 As String
    Dim finds As Integer

    On Error Resume Next
    t1 = tr1
    t2 = tr2
    tmp1 = RTrim$(t1)
    tmp2 = RTrim$(t2)
    If tmp1 <> "" And tmp2 <> "" Then
       p = True
       If Val(tmp1) <> 0 Then
          tmp = Format(Val(tmp2) / Val(tmp1) * 100, "fixed")
       Else
          tmp = "0.00"
       End If
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next

    If (MsgBox("确实要中止转换吗？", 33, "提示")) = 1 Then
       Convert_Stop = True
    End If
End Sub

Private Sub Timer1_Timer()
    Dim data_cell As cell
    Dim data_tch As tch
    Dim old_tch As tch
    Dim data_sdc As sdcch
    Dim old_sdc As sdcch
    Dim ltxt As String
    Dim buffer As String * 1
    Dim s As String
    Dim p As Boolean
    Dim ci As String * 5
    Dim recordno As Long
    Dim tmp As String, tm1 As String, tm2 As String
    Dim i As Long, j As Long, l As Long, h As Long
    Dim k  As Variant, t As Integer, q As Integer, Y As Integer
    Dim e As Integer
    Dim finds As Integer, endfrag As Integer, ot As Integer
    Dim lenth As Long
    Dim nline As Long
    Dim bline As Integer, scnline As Integer, scnline2 As Integer
    Dim percent_step  As Integer, bs As Integer, bs2 As Integer
    Dim End_Char As String * 1
    Dim MyPosVar As Variant
    Dim filetemp As String
    Dim FileCols As Long
    Dim StartFlag As Boolean
    Dim Citemp As String
    
    On Error Resume Next
    MoreFile = False
    For i = 1 To 50
        If convert_filename(i) = "" Then
           convert_num = i - 1
           Exit For
        End If
    Next
    End_Char = Chr(26)
    recordno = 0
    data_tch.bb = " "
    data_tch.Name = space$(10)
    data_tch.tcol1 = space$(6)
    data_tch.tcol2 = space$(6)
    data_tch.tcol3 = space$(6)
    data_sdc.bb = " "
    data_sdc.Name = space$(10)
    data_sdc.scol1 = space$(6)
    data_sdc.scol2 = space$(6)
    data_sdc.scol3 = space$(6)
    data_sdc.scol4 = space$(6)        'add
    data_sdc.scol5 = space$(6)
    data_sdc.scol6 = space$(6)        'add
    
    On Error GoTo 0
    Gsm_FileName = Gsm_Path + "\tch.dbf"
    Gsm_File2 = Gsm_Path + "\sts\tch_sts.dbf"
    FileCopy Gsm_FileName, Gsm_File2
    Gsm_FileName = Gsm_Path + "\cch.dbf"
    Gsm_File2 = Gsm_Path + "\sts\cch_sts.dbf"
    FileCopy Gsm_FileName, Gsm_File2
    
    filetemp = Gsm_Path + "\ststemp.tab"
    mapinfo.do "Register Table " + Chr(34) + convert_filename(1) + Chr(34) + " TYPE ASCII Delimiter 9 Titles Charset " + Chr(34) + "CodePage437" + Chr(34) + " Into " + Chr(34) + filetemp + Chr(34)
    mapinfo.do "open table " + Chr(34) + filetemp + Chr(34)
    FileCols = mapinfo.eval("tableinfo(ststemp,4)")
    If FileCols = 24 Then
       ProgressBar1.Value = 1
       Label1.Caption = "正在转换 " + convert_filename(1)
       FileCols = mapinfo.eval("tableinfo(ststemp,8)")
       percent_step = Int((FileCols / 100) + 0.5)
       bs = 0
       recordno = 0
       StartFlag = False
       Gsm_FileName = Gsm_Path + "\sts\tch_sts.dbf"
       Open Gsm_FileName For Binary As #2
       Open Gsm_File2 For Binary As #3
       Seek #2, 866
       Seek #3, 706
       For i = 1 To FileCols
           If bs = percent_step And ProgressBar1.Value < 90 Then
              ProgressBar1.Value = ProgressBar1.Value + 1
              bs = 0
           End If
           bs = bs + 1
           If StartFlag Then
              Citemp = mapinfo.eval("ststemp.col3")
              finds = InStr(Citemp, "-")
              If finds > 0 Then
                 data_sdc.ci = Right(Citemp, Len(Citemp) - finds)
                 data_tch.ci = data_sdc.ci
              Else
                 data_sdc.ci = Citemp
                 data_tch.ci = data_sdc.ci
              End If
              data_sdc.scol2 = mapinfo.eval("ststemp.col7")
              If InStr(mapinfo.eval("ststemp.col8"), "%") > 0 Then
                 data_sdc.scol(7) = Left(mapinfo.eval("ststemp.col8"), Len(mapinfo.eval("ststemp.col8")) - 1)
              Else
                 data_sdc.scol(7) = mapinfo.eval("ststemp.col8")
              End If
              data_sdc.scol(6) = mapinfo.eval("ststemp.col9")
              data_sdc.scol6 = mapinfo.eval("ststemp.col10")
              data_sdc.scol4 = mapinfo.eval("ststemp.col11")
              If InStr(mapinfo.eval("ststemp.col20"), "%") > 0 Then
                 data_sdc.scol(10) = Left(mapinfo.eval("ststemp.col20"), Len(mapinfo.eval("ststemp.col20")) - 1)
              Else
                 data_sdc.scol(10) = mapinfo.eval("ststemp.col20")
              End If
              data_sdc.scol(9) = mapinfo.eval("ststemp.col24")
              If InStr(mapinfo.eval("ststemp.col12"), "%") > 0 Then
                 data_tch.tcol(11) = Left(mapinfo.eval("ststemp.col12"), Len(mapinfo.eval("ststemp.col12")) - 1)
              Else
                 data_tch.tcol(11) = mapinfo.eval("ststemp.col12")
              End If
              data_tch.tcol(10) = mapinfo.eval("ststemp.col13")
              data_tch.tcol(4) = mapinfo.eval("ststemp.col14")
              data_tch.tcol(1) = mapinfo.eval("ststemp.col15")
              If InStr(mapinfo.eval("ststemp.col16"), "%") > 0 Then
                 data_tch.tcol(18) = Left(mapinfo.eval("ststemp.col16"), Len(mapinfo.eval("ststemp.col16")) - 1)
              Else
                 data_tch.tcol(18) = mapinfo.eval("ststemp.col16")
              End If
              If InStr(mapinfo.eval("ststemp.col17"), "%") > 0 Then
                 data_tch.tcol(17) = Left(mapinfo.eval("ststemp.col17"), Len(mapinfo.eval("ststemp.col17")) - 1)
              Else
                 data_tch.tcol(17) = mapinfo.eval("ststemp.col17")
              End If
              If InStr(mapinfo.eval("ststemp.col21"), "%") > 0 Then
                 data_tch.tcol(14) = Left(mapinfo.eval("ststemp.col21"), Len(mapinfo.eval("ststemp.col21")) - 1)
              Else
                 data_tch.tcol(14) = mapinfo.eval("ststemp.col21")
              End If
              data_tch.tcol(2) = mapinfo.eval("ststemp.col22")
              data_tch.tcol(13) = mapinfo.eval("ststemp.col23")
              recordno = recordno + 1
              Put #2, , data_tch
              Put #3, , data_sdc
              mapinfo.do "fetch next from ststemp"
           Else
              If InStr(mapinfo.eval("ststemp.col8"), "%") > 0 Then
                 StartFlag = True
                 i = i - 1
              Else
                 mapinfo.do "fetch next from ststemp"
              End If
           End If
       Next
       mapinfo.do "close table ststemp"
       filetemp = Gsm_Path + "\ststemp.*"
       Kill filetemp
       Seek #2, 5
       Put #2, , recordno
       Seek #3, 5
       Put #3, , recordno
       Close
       Unload Me
       Exit Sub
    Else
       If FileCols > 20 Then
          MsgBox "转换文件格式不对！"
          mapinfo.do "close table ststemp"
          filetemp = Gsm_Path + "\ststemp.*"
          Kill filetemp
          Unload Me
          Exit Sub
       Else
          mapinfo.do "close table ststemp"
          filetemp = Gsm_Path + "\ststemp.*"
          Kill filetemp
       End If
    End If
    
    Open convert_filename(1) For Binary As #1
    Gsm_FileName = Gsm_Path + "\sts\tch_sts.dbf"
    Open Gsm_FileName For Binary As #2
    Open Gsm_File2 For Binary As #3
    
    lenth = FileLen(convert_filename(1))
    nline = lenth / 80
    bline = Fix(nline / 100)
    percent_step = 1
    If bline = 0 Then
       bline = 1
       percent_step = 100 / nline
    End If
    bs = 1
    scnline = 0
    ProgressBar1.Value = 1
    Label1.Caption = "正在转换 " + convert_filename(1)
    
    Call ch_line(ltxt, endfrag, convert_filename(1), True)
    
    Do
       Call ch_line(ltxt, endfrag, convert_filename(1), False)
       If endfrag = 1 Then
          ltxt = ""
          Exit Do
       End If
       If Len(ltxt) > 5 Then
          ltxt = LTrim$(ltxt)
          If Len(ltxt) > 5 Then
             If InStr(UCase(ltxt), "TIME  MODE:") > 0 Then
                Call ch_line(ltxt, endfrag, convert_filename(1), False)
                If endfrag = 1 Then
                   ltxt = ""
                End If
                Exit Do
             End If
          End If
       End If
    Loop
    If Len(ltxt) = 0 Then
       Close
       MsgBox "转换文件格式不对！"
       Unload Me
       Exit Sub
    End If
    
    If Len(ltxt) > 4 Then
       s = Left(ltxt, 4)
       If UCase(s) = "DATE" Then
          ltxt = LTrim$(Right(ltxt, Len(ltxt) - 4))
          If Len(ltxt) = 0 Then
             Close
             MsgBox "转换文件格式不对！", 64, "提示"
             Unload Me
             Exit Sub
          End If
       End If
       s = Left(ltxt, 4)
       If UCase(s) = "TIME" Then
          ltxt = LTrim$(Right(ltxt, Len(ltxt) - 4))
          If Len(ltxt) = 0 Then
             Close
             MsgBox "转换文件格式不对！", 64, "提示"
             Unload Me
             Exit Sub
          End If
       End If
    End If
    Call fa_line(ltxt)
    Call ch_line(ltxt, endfrag, convert_filename(1), False)
    ltxt = LTrim$(ltxt)
    Call fb_line(ltxt)
    Call ch_line(ltxt, endfrag, convert_filename(1), False)
    If endfrag = 1 Then
       ltxt = ""
    End If
    If Len(ltxt) > 0 Then
       ltxt = LTrim$(ltxt)
       Call fc_line(ltxt)
    End If
    For i = 1 To 10
        If a(i) = "" Then
           n = i - 1
           Exit For
        End If
    Next
    make_pos
'*************************************************
'*************************************************
    bs = 1
    scnline = 0
    bs2 = 1
    scnline2 = 0

    Seek #2, 866
    Seek #3, 706
    Do
       If ProgressBar1.Value < 99 Then
          scnline2 = scnline2 + 1
       End If
       If scnline2 = bs2 * bline And ProgressBar1.Value + percent_step < 99 Then
          If ProgressBar1.Value + percent_step < 99 Then
             ProgressBar1.Value = ProgressBar1.Value + percent_step
          End If
          bs2 = bs2 + 1
       End If

       For i = 1 To 33
           hcol(i) = space$(6)
       Next
       Call hua_ci(ltxt, endfrag, convert_filename(1))
       If endfrag = 1 Then
          Exit Do
       Else
          If endfrag = 3 Then
             Put #2, , End_Char
             Seek #2, 5
             Put #2, , recordno
             Put #3, , End_Char
             Seek #3, 5
             Put #3, , recordno
             Call ch_line(ltxt, endfrag, convert_filename(1), False)
             i = 1
             GoTo MoreFileProcess
             
          End If
       End If
       For i = 1 To 3
           finds = InStr(ltxt, "-")
           ltxt = Right(ltxt, Len(ltxt) - finds)
       Next
       data_tch.ci = ltxt
       data_sdc.ci = ltxt
       'Call hua_time(ltxt, endfrag, convert_filename(1))
       Call ch_line(ltxt, endfrag, convert_filename(1), False)
       If endfrag = 1 Then
          Exit Do
       End If
       If InStr(ltxt, ":") > 0 Then
          ltxt = Trim(Right(ltxt, Len(ltxt) - InStr(ltxt, " ")))
       End If
       Call get_da(ltxt)
       For i = 1 To 10
           data_sdc.scol(i) = hcol(i)
           data_tch.tcol(i) = hcol(13 + i)
       Next
       For i = 11 To 18
           data_tch.tcol(i) = hcol(13 + i)
       Next
       data_sdc.scol4 = hcol(33)
       data_sdc.scol6 = hcol(32)
       recordno = recordno + 1
       Put #2, , data_tch
       Put #3, , data_sdc
       If endfrag = 2 Then
          Exit Do
       End If
       DoEvents
       If Convert_Stop = True Then
          Put #2, , End_Char
          Seek #2, 5
          Put #2, , recordno
          Put #3, , End_Char
          Seek #3, 5
          Put #3, , recordno
          Close
          Unload Me
          Exit Sub
       End If
    Loop
    Put #2, , End_Char
    Seek #2, 5
    Put #2, , recordno
    Put #3, , End_Char
    Seek #3, 5
    Put #3, , recordno
    Close #1
'*************************************************
'*************************************************
    For i = 2 To convert_num
        If ProgressBar1.Value < 100 Then
           ProgressBar1.Value = 100
        End If
    
        For k = 1 To 10
            a(k) = ""
            b(k) = ""
            c(k) = ""
            colnum(k) = 0
            pn(k) = 0
        Next
        
        Open convert_filename(i) For Binary As #1
        Call ch_line(ltxt, endfrag, convert_filename(i), True)
        
        lenth = FileLen(convert_filename(i))
        nline = lenth / 80
        bline = Fix(nline / 100)
        percent_step = 1
        If bline = 0 Then
           bline = 1
           percent_step = 100 / nline
        End If
        bs = 1
        scnline = 0
        ProgressBar1.Value = 1
        Label1.Caption = "正在转换 " + convert_filename(i)
        Do
           Call ch_line(ltxt, endfrag, convert_filename(i), False)
           If endfrag = 1 Then
              ltxt = ""
              Exit Do
           End If
           If Len(ltxt) > 5 Then
              ltxt = LTrim$(ltxt)
              If Len(ltxt) > 5 Then
                 s = Mid(ltxt, 1, 4)
                 If UCase(s) = "DATE" Then
                    ltxt = LTrim$(Right(ltxt, Len(ltxt) - 4))
                    If Len(ltxt) > 0 Then
                       Exit Do
                    End If
                 End If
              End If
           End If
        Loop
        If Len(ltxt) = 0 Then
           Close
           MsgBox "转换文件格式不对！"
           Unload Me
           Exit Sub
        End If
MoreFileProcess:
        If Len(ltxt) > 4 Then
           s = Left(ltxt, 4)
           If UCase(s) = "DATE" Then
              ltxt = LTrim$(Right(ltxt, Len(ltxt) - 4))
              If Len(ltxt) = 0 Then
                 Close
                 MsgBox "转换文件格式不对！"
                 Unload Me
                 Exit Sub
              End If
           End If
           s = Left(ltxt, 4)
           If UCase(s) = "TIME" Then
              ltxt = LTrim$(Right(ltxt, Len(ltxt) - 4))
              If Len(ltxt) = 0 Then
                 Close
                 MsgBox "转换文件格式不对！"
                 Unload Me
                 Exit Sub
              End If
           End If
        End If
        Call fa_line(ltxt)
        Call ch_line(ltxt, endfrag, convert_filename(i), False)
        ltxt = LTrim$(ltxt)
        Call fb_line(ltxt)
        Call ch_line(ltxt, endfrag, convert_filename(i), False)
        If endfrag = 1 Then
           ltxt = ""
        End If
        If Len(ltxt) > 0 Then
           ltxt = LTrim$(ltxt)
           Call fc_line(ltxt)
        End If
        For j = 1 To 10
            If a(j) = "" Then
               n = j - 1
               Exit For
            End If
        Next
        make_pos
'**********************************************************
'**********************************************************
        If Not MoreFile Then
           bs = 1
           scnline = 0
           bs2 = 1
           scnline2 = 0
        End If
        Do
            If ProgressBar1.Value < 99 Then
               scnline2 = scnline2 + 1
            End If
            If scnline2 = bs2 * bline And ProgressBar1.Value + percent_step < 99 Then
               If ProgressBar1.Value + percent_step < 99 Then
                  ProgressBar1.Value = ProgressBar1.Value + percent_step
               End If
               bs2 = bs2 + 1
            End If

            For t = 1 To 33
                hcol(t) = space$(6)
            Next
            Call hua_ci(ltxt, endfrag, convert_filename(i))
            If endfrag = 1 Then
               Exit Do
            Else
               If endfrag = 3 Then
                  For k = 1 To 10
                      a(k) = ""
                      b(k) = ""
                      c(k) = ""
                      colnum(k) = 0
                      pn(k) = 0
                  Next
                  Call ch_line(ltxt, endfrag, convert_filename(i), False)
                  GoTo MoreFileProcess
               End If
            End If
            For q = 1 To 3
                finds = InStr(ltxt, "-")
                ltxt = Right(ltxt, Len(ltxt) - finds)
            Next
            data_tch.ci = ltxt
            data_sdc.ci = ltxt
            'Call hua_time(ltxt, endfrag, convert_filename(i))
            Call ch_line(ltxt, endfrag, convert_filename(i), False)
            If endfrag = 1 Then
               Exit Do
            End If
            If InStr(ltxt, ":") > 0 Then
               ltxt = Trim(Right(ltxt, Len(ltxt) - InStr(ltxt, " ")))
            End If
            Call get_da(ltxt)
            
            'Seek #2, 770
            Seek #2, 866
            For Y = 0 To recordno - 1
                'Seek #2, 770 + 142 * y
                Seek #2, 866 + 169 * Y
                Get #2, , old_tch
                If Trim(old_tch.ci) = Trim(data_tch.ci) Then
                   'Seek #3, 610 + 112 * y
                   Seek #3, 706 + 139 * Y
                   Get #3, , old_sdc
                   For e = 1 To 10
                       If hcol(e) <> space$(6) Then
                          old_sdc.scol(e) = hcol(e)
                       End If
                   Next
                   For e = 14 To 31
                       If hcol(e) <> space$(6) Then
                          old_tch.tcol(e - 13) = hcol(e)
                       End If
                   Next
                   If hcol(33) <> space$(6) Then
                      old_sdc.scol4 = hcol(33)
                   End If
                   If hcol(32) <> space$(6) Then
                      old_sdc.scol6 = hcol(32)
                   End If
                   'Seek #2, 770 + 142 * y
                   Seek #2, 866 + 169 * Y
                   Put #2, , old_tch
                   'Seek #3, 610 + 112 * y
                   Seek #3, 706 + 139 * Y
                   Put #3, , old_sdc
                   Exit For
                End If
            Next
            If endfrag = 2 Then
               Exit Do
            End If
            DoEvents
            If Convert_Stop = True Then
               Close
               Unload Me
               Exit Sub
            End If
        Loop
        Close #1
        If MoreFile Then
           Exit For
        End If
    Next
'******************************************************
'******************************************************
GeneralProcess:
    Close #2
    Gsm_FileName = Gsm_Path + "\sts\tch_sts.dbf"
    Open Gsm_FileName For Binary As #2
    Close #3
    Gsm_FileName = Gsm_Path + "\sts\cch_sts.dbf"
    Open Gsm_FileName For Binary As #3
    Gsm_FileName = Gsm_Path + "\map\cell.dbf"
    Open Gsm_FileName For Binary As #4
    For i = 0 To recordno - 1
        'Seek #2, 770 + i * 142
        Seek #2, 866 + i * 169
        Get #2, , data_tch
        'Seek #3, 610 + i * 112
        Seek #3, 706 + i * 139
        Get #3, , data_sdc
        tm1 = RTrim$(data_sdc.scol(4))
        tm2 = RTrim$(data_sdc.scol(5))
        If tm1 <> "" And tm2 <> "" Then
           data_sdc.scol(5) = Format(Val(tm1) - Val(tm2), "#####0")
        Else
           data_sdc.scol(5) = space$(6)
        End If
        tm1 = RTrim$(data_tch.tcol(8))
        tm2 = RTrim$(data_sdc.scol(3))
        If tm1 <> "" And tm2 <> "" Then
           data_tch.tcol(8) = Format(Val(tm1) + Val(tm2), "#####0")
        End If
        p = False
        Call fini(tmp, data_sdc.scol(1), data_sdc.scol(2), p)
        If p = True Then
           data_sdc.scol(3) = tmp
        Else
           data_sdc.scol(3) = space$(6)
        End If
        tm1 = RTrim$(data_tch.tcol(10))
        tm2 = RTrim$(data_sdc.scol(8))
        If tm1 <> "" And tm2 <> "" Then
           data_tch.tcol(10) = Format(Val(tm1) + Val(tm2), "#####0")
        End If
        p = False
        Call fini(tmp, data_sdc.scol(4), data_sdc.scol(6), p)
        If p = True Then
           data_sdc.scol(8) = tmp
        Else
           data_sdc.scol(8) = space$(6)
        End If
        tmp = RTrim$(data_tch.tcol(15))
        tm1 = RTrim$(data_tch.tcol(16))
        tm2 = RTrim$(data_tch.tcol(2))
        If tmp <> "" And tm1 <> "" And tm2 <> "" Then
           data_tch.tcol(15) = Format(Val(tmp) + Val(tm1) + Val(tm2), "#####0")
           tmp = RTrim$(data_tch.tcol(17))
           If tmp <> "" Then
              data_tch.tcol(16) = Format(Val(tmp) * Val(data_tch.tcol(15)), "#####0")
           Else
              data_tch.tcol(16) = space$(6)
           End If
        End If
        data_tch.tcol(2) = space$(6)
        
'0815        k = 0
'0815        MyPosVar = 0
'0815        Do While Not EOF(4)
'0815           Seek #4, 1026 + MyPosVar
'0815           Get #4, , data_cell
'0815           If data_cell.ci = data_tch.ci Then
'0815              data_tch.Name = data_cell.Name
'0815              data_sdc.Name = data_cell.Name
'0815              Exit Do
'0815           End If
'0815           k = k + 1
'0815           MyPosVar = 277 * k
'0815        Loop
'0815        Seek #4, 1
        
        'seek #2, 770 + i * 142
        Seek #2, 866 + i * 169
        Put #2, , data_tch
        'Seek #3, 610 + i * 112
        Seek #3, 706 + i * 139
        Put #3, , data_sdc
        DoEvents
        If Convert_Stop = True Then
           Close
           Unload Me
           Exit Sub
        End If
    Next

    Close
    
    If ProgressBar1.Value < 100 Then
       If ProgressBar1.Value < 80 Then
          Do While ProgressBar1.Value < 100
               ProgressBar1.Value = ProgressBar1.Value + 1
          Loop
       Else
          ProgressBar1.Value = 100
       End If
    End If
    
    Erase a
    Erase b
    Erase c
    Erase pn
    Erase colnum
    Erase hcol
    Erase old_tch.tcol
    Erase data_tch.tcol
    Erase old_sdc.scol
    Erase data_sdc.scol
    Unload Me
End Sub
