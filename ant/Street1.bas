Attribute VB_Name = "report11"
Public Data_Report As Boolean
Public stre_s(0 To 12) As Boolean
Public stre_tab(1 To 12) As String
Dim stre_tab_cell(1 To 12) As String
Public stre_num As Integer
Public Cell_Report As Boolean
Dim all_max As Integer, all_min As Integer
Dim all_avg As Single
Dim table_s(1 To 9, 1 To 6) As Single
Dim table_f(1 To 9, 1 To 6) As String
Dim tibeh(1 To 9, 1 To 6) As Single
Dim cc_all
Dim m_per As String
Dim ta_all
Public Rcellname() As String
Public RCellNo As Integer
Public Rep_Ci() As String * 5
Public select_name As String
Public Report_Qual As Integer, Report_Full As Boolean
Public Report_Rxlev1 As Integer, Report_Rxlev2 As Integer


Sub My_Report()
    Dim putin, putin1, putin2
'    Dim cellname() As String, dtx() As String
    Dim My_Bearing() As String
    Dim dtx() As String
    Dim lon() As String, bearing() As String, down() As String
    Dim No_Ncell() As String, No_Ncell_Percent() As String
    Dim lat() As String
    Dim mapci
    Dim cellci As String * 5, oldci As String * 5
    Dim stcname As String
    Dim setup_n As Integer, tmp1_n As Integer, tmp2_n As Integer, tmp3_n As Integer
    Dim perce As Single
    Dim my_enum, tt
    Dim myb As String
    Dim com_hc As String, com_hs As String, com_hf As String
    Dim com_hmax As String, com_hmin As String, com_havg As String
    Dim com_qmax As String, com_qmin As String, com_qavg As String
    Dim com_xmax As String, com_xmin As String, com_xavg As String
    Dim r_f83 As String, r_f93 As String, r_s83 As String, r_s93 As String
    Dim rq_f As String, rq_s As String, mta As String
    Dim tt_ta As Single, ttnum, ttall
    Dim report_file As String, doc_man_file As String
    Dim stre_save(1 To 10) As String
    Dim Get_Percent As String
    Dim finds As Integer
    Dim Is_Exist As Boolean
    Dim PointNum As Integer
    Dim NTempNum As Integer
    Dim NTempTotal As Integer
    Dim My_Source(1 To 10) As String, CellArfcn(1 To 10) As String, CellBsic(1 To 10) As String
    Dim OldArfcn As String, OldBsic As String
    Dim ArfcnBsicNum As Integer
    Dim k As Integer, q As Integer, p As Integer
    Dim Rselect_Num(1 To 10) As Integer, Rtotal_Num As Long
    
    On Error Resume Next

    Dim doc_rec As doc    'hua_del
    Gsm_FileName = Gsm_Path + "\doc_man.dat"
    Open Gsm_FileName For Binary Shared As #1
    Seek #1, 1

    msg1 = stre_tab(1)
    If Left(msg1, 1) = "_" Then
       msg1 = Mid(msg1, 2)
    End If
    If sys = 0 Then
       msg = Gsm_Path + "\NORMAL\" + msg1 + ".TXT"
    Else
       msg = Gsm_Path + "\SCAN\" + msg1 + ".SCN"
    End If
    report_file = Trim(stre_tab(1))
    If Mid(report_file, 1, 1) = "_" Then
       report_file = Right(report_file, Len(report_file) - 1)
    End If
    If Right(report_file, 1) = "f" Or Right(report_file, 1) = "F" Then
       report_file = Left(report_file, Len(report_file) - 1)
    End If
    p = 1
    Is_Exist = False
    Do While Not EOF(1)
       Get #1, p, doc_rec
       doc_man_file = Trim(doc_rec.DOCNAME)
       finds = InStr(doc_man_file, "\")
       Do While finds > 0
          doc_man_file = Right(doc_man_file, Len(doc_man_file) - finds)
          finds = InStr(doc_man_file, "\")
       Loop
       finds = InStr(doc_man_file, ".")
       If finds > 0 Then
          doc_man_file = Left(doc_man_file, finds - 1)
       End If
       If UCase(report_file) = UCase(doc_man_file) Then
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
    Close #1
'    p = 1
'    Get #1, p, doc_rec
'    While Not EOF(1) And UCase(msg) <> Mid(UCase(Trim(doc_rec.DOCNAME)), 3)
'          Get #1, p, doc_rec
'          p = p + 74
'          Seek #1, p
'    Wend

    Dim OBJ As String, DIST As String, BACK As String, WETHER As String
  If Is_Exist = True Then
    Select Case doc_rec.TESTOBJECT
           Case "0"
                 OBJ = "区域"
           Case "1"
                 OBJ = "基站"
           Case "2"
                 OBJ = "路段"
           Case "3"
                 OBJ = "室内"
    End Select
    Select Case doc_rec.TESTDIST
           Case "0"
                 DIST = "一般调查"
           Case "1"
                 DIST = "覆盖调查"
           Case "2"
                 DIST = "切换带调整"
           Case "3"
                 DIST = "话务量调整"
           Case "4"
                 DIST = "干扰调整"
           Case "5"
                 DIST = "网络评估"
           Case "6"
                 DIST = "异网评估"
           Case "7"
                 DIST = "客户投诉"
    End Select
    Select Case doc_rec.TESTBACK
           Case "0"
                 BACK = "射频功率调整"
           Case "1"
                 BACK = "天线馈线系统调整"
           Case "2"
                 BACK = "小区参数调整及验证"
           Case "3"
                 BACK = "扩容工程"
           Case "4"
                 BACK = "上级普查"
           Case "5"
                 BACK = "气侯影响"
    End Select
    Select Case doc_rec.WEATHER
           Case "0"
                 WETHER = "晴好"
           Case "1"
                 WETHER = "阴雨"
           Case "2"
                 WETHER = "夜晚"
    End Select
    finds = InStr(doc_rec.Partner, Chr(0))
    If finds > 0 Then
       doc_rec.Partner = Left(doc_rec.Partner, finds - 1)
    End If
  End If
    
    Close
    If Cell_Report = True Then
       For i = 1 To stre_num
           stre_save(i) = stre_tab(i)
           stre_tab(i) = "my_temp" & Format(i)
       Next
'       stre_s(10) = False
    End If
    cc_all = 0
    For i = 1 To stre_num
        cc_all = cc_all + mapinfo.eval("tableinfo(" + stre_tab(i) + ",8)")
    Next
       
hua:                    ' hua_add
    Set word = CreateObject("Word.application")
    word.Visible = True
    word.Documents.Add
    If word.ActiveWindow.View.SplitSpecial = 0 Then
       word.ActiveWindow.ActivePane.View.type = 3
       word.ActiveWindow.ActivePane.View.Zoom.Percentage = 100
    Else
       word.ActiveWindow.View.type = 3
       word.ActiveWindow.View.Zoom.Percentage = 100
    End If
    word.selection.ParagraphFormat.Alignment = 1
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Size = 12
    word.selection.Font.Bold = -1
    word.selection.typetext Text:="GSM 移动电话无线网络质量调查报告"
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    
    word.selection.Font.Size = 9
    word.selection.ParagraphFormat.Alignment = 3
    word.selection.typetext Text:="网络运行局名称：" + USERNAME
    word.selection.TypeParagraph
    word.selection.typetext Text:="测  试  时  间：" + doc_rec.DATE
    word.selection.TypeParagraph
    word.selection.typetext Text:="测  试  人  员：" + doc_rec.Partner
    word.selection.TypeParagraph
    word.selection.typetext Text:="测  试  目  的：" + DIST
    word.selection.TypeParagraph
    word.selection.typetext Text:="测  试  对  象：" + OBJ
    word.selection.TypeParagraph
    word.selection.typetext Text:="测  试  背  景：" + BACK
    word.selection.TypeParagraph
    word.selection.typetext Text:="测  试  天  气：" + WETHER
    word.selection.TypeParagraph
    word.selection.typetext Text:="路  段  名  称："
    word.selection.TypeParagraph
    pri_tbl = ""
    For i = 1 To stre_num
        pri_tbl = pri_tbl + stre_tab(i) + ".tab"
        If i < stre_num Then pri_tbl = pri_tbl + "；"
    Next
    word.selection.typetext Text:="测量数据文件名：" + pri_tbl
    word.selection.TypeParagraph
    word.selection.typetext Text:="测  量  距  离："
    word.selection.TypeParagraph
    word.selection.typetext Text:="有 效 覆 盖 率："
    word.selection.TypeParagraph
    If sys = 0 Then
       word.selection.typetext Text:="测  量  模  式：通话测量"
       word.selection.TypeParagraph
    Else
       word.selection.typetext Text:="测  量  模  式：扫频测量"
       word.selection.TypeParagraph
    End If

'    Call esceed_a(4)
    word.selection.InsertBreak type:=7
    
 If stre_s(0) = True Then
    word.selection.Font.Size = 10.5
    word.selection.ParagraphFormat.Alignment = 1
    word.selection.Font.Bold = -1
    word.selection.typetext Text:="无线参数统计报告"
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.ParagraphFormat.Alignment = 3
    word.selection.Font.Size = 9
    word.selection.Font.Bold = -1
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    'word.selection.typetext Text:="城市街道测量 RELEV_Full"
    word.selection.typetext Text:="街道信道场强测量(Full)"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    rxlev_fuc
    word.selection.MoveUp Count:=8
    Call fill("56", "200", "22", True, False, False, False)
    Call fill("46", "57", "22", False, False, False, False)
    Call fill("36", "47", "22", False, False, False, False)
    Call fill("26", "37", "22", False, False, False, False)
    r_f83 = m_per
    Call fill("16", "27", "22", False, False, False, False)
    r_f93 = m_per
    Call fill("6", "17", "22", False, False, False, False)
    Call fill("-1", "7", "22", False, False, False, False)
    word.selection.EndKey unit:=5
    word.selection.Font.Bold = -1
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    word.selection.typetext Text:=Chr(9) + all_0 + Chr(9) + putin + Chr(9) + putin1 + Chr(9) + putin2
    word.selection.typetext Text:=Chr(9) + "/" + Chr(9) + "/"
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
 End If
 
 If stre_s(2) = True Then
    If stre_s(0) = False Then
       word.selection.Font.Size = 10.5
       word.selection.ParagraphFormat.Alignment = 1
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="无线参数统计报告"
       word.selection.Font.Bold = 0
       word.selection.TypeParagraph
       word.selection.TypeParagraph
       word.selection.ParagraphFormat.Alignment = 3
    End If
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.Font.Size = 9
    word.selection.Font.Bold = -1
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.typetext Text:="街道信号误码测量(Full)"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    rxqual_fuc
    word.selection.MoveUp Count:=8
    Call fill("0", "0", "23", True, True, False, False)
    Call fill("1", "0", "23", False, True, False, False)
    Call fill("2", "0", "23", False, True, False, False)
    Call fill("3", "0", "23", False, True, False, False)
    rq_f = m_per
    Call fill("4", "0", "23", False, True, False, False)
    Call fill("5", "0", "23", False, True, False, False)
    Call fill("6", "0", "23", False, True, False, False)
    Call fill("7", "0", "23", False, True, False, False)
    word.selection.EndKey unit:=5
    word.selection.Font.Bold = -1
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    word.selection.typetext Text:=Chr(9) + all_0 + Chr(9) + putin1
    word.selection.typetext Text:=Chr(9) + "/" + Chr(9) + "/"
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
 End If
 
 If stre_s(4) = True Then
    If stre_s(0) = False And stre_s(2) = False Then
       word.selection.Font.Size = 10.5
       word.selection.ParagraphFormat.Alignment = 1
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="无线参数统计报告"
       word.selection.Font.Bold = 0
       word.selection.TypeParagraph
       word.selection.TypeParagraph
       word.selection.ParagraphFormat.Alignment = 3
    End If
 
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.Font.Size = 9
    word.selection.Font.Bold = -1
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.typetext Text:="小区服务范围统计(TIMING ADVANCE)"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    word.selection.Font.Size = 9
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(3), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.7), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(6.39), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(8.08), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(9.79), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(11.85), Alignment:=1, Leader:=0
    word.selection.typetext Text:=Chr(9) & "测量数" & Chr(9) & "最大值" & Chr(9) & "平均值" & Chr(9) & "最小值" & Chr(9) & "百分比" & Chr(9) & "累计百分比"
    word.selection.TypeParagraph
    'word.selection.typetext Text:="0<=X<7"
    word.selection.typetext Text:="X=0"
    word.selection.TypeParagraph
    'word.selection.typetext Text:="7<=X<15"
    word.selection.typetext Text:="X=1"
    word.selection.TypeParagraph
    'word.selection.typetext Text:="15<=X<23"
    word.selection.typetext Text:="X=2"
    word.selection.TypeParagraph
    'word.selection.typetext Text:="23<=X<31"
    word.selection.typetext Text:="X=3"
    word.selection.TypeParagraph
    'word.selection.typetext Text:="31<=X<39"
    word.selection.typetext Text:="X=4"
    word.selection.TypeParagraph
    'word.selection.typetext Text:="39<=X<47"
    word.selection.typetext Text:="X=5"
    word.selection.TypeParagraph
    'word.selection.typetext Text:="47<=X<55"
    word.selection.typetext Text:="X=6"
    word.selection.TypeParagraph
    'word.selection.typetext Text:="55<=X<63"
    word.selection.typetext Text:="X=7"
    word.selection.TypeParagraph
    word.selection.typetext Text:="7<X<=30"
    word.selection.TypeParagraph
    word.selection.typetext Text:="30<X<=63"
    word.selection.TypeParagraph
    word.selection.typetext Text:="总计"
    word.selection.MoveUp Count:=10
    'Call fill("-1", "7", "26", True, False, True, False)
    'Call fill("6", "15", "26", False, False, True, False)
    'Call fill("14", "23", "26", False, False, True, False)
    'Call fill("22", "31", "26", False, False, True, False)
    'Call fill("30", "39", "26", False, False, True, False)
    'Call fill("38", "47", "26", False, False, True, False)
    'Call fill("46", "55", "26", False, False, True, False)
    'Call fill("54", "63", "26", False, False, True, False)
    Call fill("0", "0", "26", True, False, True, False)
    Call fill("1", "0", "26", False, False, True, False)
    Call fill("2", "0", "26", False, False, True, False)
    Call fill("3", "0", "26", False, False, True, False)
    Call fill("4", "0", "26", False, False, True, False)
    Call fill("5", "0", "26", False, False, True, False)
    Call fill("6", "0", "26", False, False, True, False)
    Call fill("7", "0", "26", False, False, True, False)
    Call fill("7", "31", "26", False, False, True, False)
    Call fill("30", "64", "26", False, False, True, False)
    word.selection.EndKey unit:=5
    word.selection.Font.Bold = -1
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / ta_all
    putin1 = Format$(all_avg, "fixed")
    mta = putin1
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    word.selection.typetext Text:=Chr(9) + all_0 + Chr(9) + putin + Chr(9) + putin1 + Chr(9) + putin2
    word.selection.typetext Text:=Chr(9) + "/" + Chr(9) + "/"
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
 End If
 
  If stre_s(1) = True Then
    If stre_s(0) = False And stre_s(2) = False And stre_s(4) = False Then
       word.selection.Font.Size = 10.5
       word.selection.ParagraphFormat.Alignment = 1
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="无线参数统计报告"
       word.selection.Font.Bold = 0
       word.selection.TypeParagraph
       word.selection.TypeParagraph
       word.selection.ParagraphFormat.Alignment = 3
    End If
    
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.Font.Size = 9
    word.selection.Font.Bold = -1
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.typetext Text:="街道信道场强测量(Sub)"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    rxlev_fuc
    word.selection.MoveUp Count:=8
    Call fill("56", "200", "24", True, False, False, False)
    Call fill("46", "57", "24", False, False, False, False)
    Call fill("36", "47", "24", False, False, False, False)
    Call fill("26", "37", "24", False, False, False, False)
    r_s83 = m_per
    Call fill("16", "27", "24", False, False, False, False)
    r_s93 = m_per
    Call fill("6", "17", "24", False, False, False, False)
    Call fill("-1", "7", "24", False, False, False, False)
    word.selection.EndKey unit:=5
    word.selection.Font.Bold = -1
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    word.selection.typetext Text:=Chr(9) + all_0 + Chr(9) + putin + Chr(9) + putin1 + Chr(9) + putin2
    word.selection.typetext Text:=Chr(9) + "/" + Chr(9) + "/"
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
  End If
   
  If stre_s(3) = True Then
    If stre_s(0) = False And stre_s(2) = False And stre_s(4) = False And stre_s(3) = False Then
       word.selection.Font.Size = 10.5
       word.selection.ParagraphFormat.Alignment = 1
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="无线参数统计报告"
       word.selection.Font.Bold = 0
       word.selection.TypeParagraph
       word.selection.TypeParagraph
       word.selection.ParagraphFormat.Alignment = 3
    End If
  
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.Font.Size = 9
    word.selection.Font.Bold = -1
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.typetext Text:="街道信号误码测量(Sub)"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    rxqual_fuc
    word.selection.MoveUp Count:=8
    Call fill("0", "0", "25", True, True, False, False)
    Call fill("1", "0", "25", False, True, False, False)
    Call fill("2", "0", "25", False, True, False, False)
    Call fill("3", "0", "25", False, True, False, False)
    rq_s = m_per
    Call fill("4", "0", "25", False, True, False, False)
    Call fill("5", "0", "25", False, True, False, False)
    Call fill("6", "0", "25", False, True, False, False)
    Call fill("7", "0", "25", False, True, False, False)
    word.selection.EndKey unit:=5
    word.selection.Font.Bold = -1
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    word.selection.typetext Text:=Chr(9) + all_0 + Chr(9) + putin1
    word.selection.typetext Text:=Chr(9) + "/" + Chr(9) + "/"
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
  End If
  
  If stre_s(5) = True Then
    If stre_s(0) = False And stre_s(2) = False And stre_s(4) = False And stre_s(3) = False And stre_s(1) = False Then
       word.selection.Font.Size = 10.5
       word.selection.ParagraphFormat.Alignment = 1
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="无线参数统计报告"
       word.selection.Font.Bold = 0
       word.selection.TypeParagraph
       word.selection.TypeParagraph
       word.selection.ParagraphFormat.Alignment = 3
    End If
    
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.Font.Size = 9
    word.selection.Font.Bold = -1
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.typetext Text:="手机功率衰减统计(MOBILE TRANSMITTING PROWER(dBm))"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    word.selection.Font.Size = 9
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(2.82), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.72), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(6.58), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(8.7), Alignment:=1, Leader:=0
    word.selection.typetext Text:=Chr(9) & "测量数" & Chr(9) & "平均值" & Chr(9) & "百分比" & Chr(9) & "累计百分比"
    word.selection.TypeParagraph
    word.selection.typetext Text:="2 (39dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="3 (37dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="4 (35dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="5 (33dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="6 (31dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="7 (29dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="8 (27dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="9 (25dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="10 (23dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="11 (21dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="12 (19dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="13 (17dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="14 (15dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="15 (13dBm)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="总计"
    word.selection.MoveUp Count:=14
    Call fill("2", "0", "27", True, True, True, False)
    Call fill("3", "0", "27", False, True, True, False)
    Call fill("4", "0", "27", False, True, True, False)
    Call fill("5", "0", "27", False, True, True, False)
    Call fill("6", "0", "27", False, True, True, False)
    Call fill("7", "0", "27", False, True, True, False)
    Call fill("8", "0", "27", False, True, True, False)
    Call fill("9", "0", "27", False, True, True, False)
    Call fill("10", "0", "27", False, True, True, False)
    Call fill("11", "0", "27", False, True, True, False)
    Call fill("12", "0", "27", False, True, True, False)
    Call fill("13", "0", "27", False, True, True, False)
    Call fill("14", "0", "27", False, True, True, False)
    Call fill("15", "0", "27", False, True, True, False)
    
    word.selection.EndKey unit:=5
    word.selection.Font.Bold = -1
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    word.selection.typetext Text:=Chr(9) + all_0 + Chr(9) + putin1
    word.selection.typetext Text:=Chr(9) + "/" + Chr(9) + "/"
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
  End If
    
  If stre_s(6) = True Then
    word.selection.Font.Size = 10.5
    word.selection.ParagraphFormat.Alignment = 1
    word.selection.Font.Bold = -1
    word.selection.typetext Text:="信令事件统计报告"
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.ParagraphFormat.Alignment = 3
  
    word.selection.Font.Size = 9
    word.selection.Font.Bold = -1
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.typetext Text:="通话统计"
    word.selection.Font.Underline = 0
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Size = 9
    
    setup_n = mess_num("SETUP")
    word.selection.typetext Text:="Call Attempts： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(setup_n)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：SETUP)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    Call xt_time(False, tmp2_n, True)
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Successes： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp2_n)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：成对的 CONNECT...DISCONNECT)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    tmp1_n = setup_n - tmp2_n
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call failures： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp1_n)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：CONNECT...DISCONNECT 不成对)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    
    word.selection.Font.Bold = -1
    word.selection.Font.Underline = 1
    word.selection.typetext Text:="通话建立（SETUP）"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    Call xt_time(True, my_enum, False)
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Setup Successes： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(my_enum)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：成对的 SETUP...ASSIGNMENT COMPLETE)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    tmp1_n = mess_num("ASSIGNMENT FAILURES")
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Setup Failures： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp1_n)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：ASSIGNMENT FAILURES)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    tmp2_n = setup_n - my_enum - tmp1_n
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Unknow Call Setup： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp2_n)
    word.selection.TypeParagraph
    
    word.selection.TypeParagraph
    word.selection.Font.Underline = 1
    word.selection.typetext Text:="通话振铃（ALERT）"
    word.selection.Font.Underline = 0
    tmp1_n = mess_num("ALERTING")
    tmp2_n = setup_n - tmp1_n
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Alert Successes： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp1_n)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：ALERTING)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Alert Failures： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp2_n)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：SETUP 后没有 ALERTING 建立的过程)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    tmp1_n = setup_n - tmp1_n - tmp2_n
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Alert /Unknow： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp1_n)
    
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Underline = 1
    word.selection.typetext Text:="通话接续（CONNECT）"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    tmp1_n = mess_num("CONNECT")
    Call xt_time(True, tmp2_n, True)
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Connect Successes： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp1_n)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：CONNECT)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Connect Failures： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp2_n)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：SETUP 后没有 CONNECT 建立的过程)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    tmp1_n = setup_n - tmp1_n - tmp2_n
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Unknown Connect： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp1_n)
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    
    word.selection.Font.Underline = 1
    word.selection.typetext Text:="通话位置更新（LOCATION UPDATING）"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    tmp1_n = mess_num("LOCATION UPDATING REQUEST")
    If tmp1_n = 0 Then
       myb = "N/A"
    Else
       tmp2_n = mess_num("LOCATION UPDATING ACCEPT")
       tmp3_n = tmp1_n - tmp2_n
       myb = str(tmp2_n)
    End If
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Location Updating Attempts： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(tmp1_n)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：LOCATION UPDATING REQUEST)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Location Updating Successes： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=myb
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：LOCATION UPDATING ACCEPT)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    
    If tmp1_n = 0 Then
       myb = "N/A"
    Else
       tmp2_n = mess_num("LOCATION UPDATING REJECT")
       tmp3_n = tmp3_n - tmp2_n
       myb = str(tmp2_n)
    End If
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Location Updating Failures： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=myb
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：LOCATION UPDATING REJECT)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Unknown Location Updating： "
    word.selection.Font.Bold = -1
    If tmp1_n = 0 Then
       myb = "N/A"
    Else
       myb = str(tmp3_n)
    End If
    word.selection.typetext Text:=myb
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.TypeParagraph
  End If
    
  If stre_s(7) = True Then
    word.selection.Font.Size = 10.5
    word.selection.ParagraphFormat.Alignment = 1
    word.selection.Font.Bold = -1
    word.selection.typetext Text:="切换及系统评估报告"
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.ParagraphFormat.Alignment = 3
  
    word.selection.Font.Size = 9
    word.selection.Font.Bold = -1
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.typetext Text:="切换（HANDOVER）统计"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Size = 9
    Call hand_time(enum1, enum2, enum3)
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="切换发起（Handover Command）： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(enum1)
    com_hc = str(enum1)
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="切换成功（Handover Successes）： "
    word.selection.Font.Bold = -1
    If enum1 = 0 Then
       word.selection.typetext Text:="N/A"
    Else
       word.selection.typetext Text:=str(enum2)
       com_hs = str(enum2)
    End If
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="切换失败（Handover Failures）： "
    word.selection.Font.Bold = -1
    If enum1 = 0 Then
       word.selection.typetext Text:="N/A"
    Else
       word.selection.typetext Text:=str(enum3)
       com_hf = str(enum3)
    End If
    If enum1 = 0 Then GoTo no_time
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    hand_zz
    
    word.selection.Font.Size = 9
    word.selection.Font.Bold = -1
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.typetext Text:="切换性能评估表"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：HANDOVER COMMAND 与 HANDOVER COMPLETE 或 HANDOVER FAILURES)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    word.selection.Font.Bold = -1
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(3), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.7), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(6.56), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(8.68), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(10.37), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(12.28), Alignment:=1, Leader:=0
    word.selection.typetext Text:=Chr(9) & "测量数" & Chr(9) & "最大值(ms)" & Chr(9) & "平均值(ms)" & Chr(9) & "最小值(ms)" & Chr(9) & "百分比" & Chr(9) & "累计百分比"
    word.selection.TypeParagraph
    
    Call wor_fi("0s<=x<0.1s", 1)
    Call wor_fi("0.1s<=x<0.2s", 2)
    Call wor_fi("0.2<=x<0.3s", 3)
    Call wor_fi("0.3s<=x<0.5s", 4)
    Call wor_fi("0.5s<=x<1s", 5)
    Call wor_fi("1s<=x<2s", 6)
    Call wor_fi("2s<=x<5s", 7)
    Call wor_fi("5s<=x<15s", 8)
    Call wor_fi("总计", 9)
    word.selection.Font.Bold = 0
    com_hmax = table_f(9, 2)
    com_havg = table_f(9, 3)
    com_hmin = table_f(9, 4)
    
    If enum1 < 2 Then GoTo no_time
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    For i = 1 To 9
        For j = 1 To 6
            table_s(i, j) = tibeh(i, j)
        Next
    Next
    hand_zz
    word.selection.Font.Underline = 1
    word.selection.Font.Bold = -1
    word.selection.typetext Text:="切换间隔时间统计表"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：HANDOVER COMMAND 与下一个 HANDOVER COMMAND)"
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
    word.selection.Font.Bold = -1
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(3), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.7), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(6.56), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(8.68), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(10.37), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(12.28), Alignment:=1, Leader:=0
    word.selection.typetext Text:=Chr(9) & "测量数" & Chr(9) & "最大值(ms)" & Chr(9) & "平均值(ms)" & Chr(9) & "最小值(ms)" & Chr(9) & "百分比" & Chr(9) & "累计百分比"
    word.selection.TypeParagraph
    
    Call wor_fi("0s<=x<1s", 1)
    Call wor_fi("1s<=x<2s", 2)
    Call wor_fi("2<=x<4s", 3)
    Call wor_fi("4s<=x<10s", 4)
    Call wor_fi("10s<=x<120s", 5)
    Call wor_fi("2min<=x<20min", 6)
    Call wor_fi("总计", 9)
    
    com_qmax = table_f(9, 2)
    com_qavg = table_f(9, 3)
    com_qmin = table_f(9, 4)
no_time:
    word.selection.TypeParagraph
    word.selection.TypeParagraph
  End If
  
  If stre_s(9) = True Then
    word.selection.Font.Size = 9
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.Font.Bold = -1
    word.selection.typetext Text:="系统响应时间统计表"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:="    (信令过程：CHANNEL REQUEST 到 ASSIGNMENT COMMAND)"
    word.selection.Font.colorindex = 0
    word.selection.Font.Bold = -1
    word.selection.Font.Size = 9
    word.selection.TypeParagraph
    Call xt_time(False, tt, False)
    hand_zz
    If table_s(9, 1) = 0 Then
       com_xmax = "N/A"
       com_xavg = "N/A"
       com_xmin = "N/A"
    Else
       com_xmax = table_f(9, 2)
       com_xavg = table_f(9, 3)
       com_xmin = table_f(9, 4)
    End If
    If table_s(9, 1) = 0 Then
       word.selection.Font.Bold = 0
       word.selection.typetext Text:="无 CHANNEL REQUEST"
       word.selection.Font.Bold = -1
       word.selection.TypeParagraph
       GoTo ewi
    End If
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(3), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.7), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(6.56), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(8.68), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(10.37), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(12.28), Alignment:=1, Leader:=0
    word.selection.typetext Text:=Chr(9) & "测量数" & Chr(9) & "最大值(ms)" & Chr(9) & "平均值(ms)" & Chr(9) & "最小值(ms)" & Chr(9) & "百分比" & Chr(9) & "累计百分比"
    word.selection.TypeParagraph
    Call wor_fi("0s<=x<0.1s", 1)
    Call wor_fi("0.1s<=x<0.2s", 2)
    Call wor_fi("0.2<=x<0.3s", 3)
    Call wor_fi("0.3s<=x<0.5s", 4)
    Call wor_fi("0.5s<=x<1s", 5)
    Call wor_fi("1s<=x<2s", 6)
    Call wor_fi("2s<=x<5s", 7)
    Call wor_fi("5s<=x<15s", 8)
    Call wor_fi("总计", 9)
    com_xmax = table_f(9, 2)
    com_xavg = table_f(9, 3)
    com_xmin = table_f(9, 4)
ewi:
    word.selection.TypeParagraph
    word.selection.TypeParagraph
  End If
    
  If stre_s(8) = True Then
    word.selection.Font.Size = 10.5
    word.selection.ParagraphFormat.Alignment = 1
    word.selection.Font.Bold = -1
    word.selection.typetext Text:="综合指标统计报告"
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.ParagraphFormat.Alignment = 3
    
    'word.selection.Font.Size = 9
    'If word.selection.Font.Underline = 0 Then
    '   word.selection.Font.Underline = 1
    'End If
    'word.selection.Font.Bold = -1
    'word.selection.typetext Text:="综合指标统计项"
    'word.selection.typeparagraph
    word.selection.Font.Size = 9
    word.selection.TypeParagraph
    'word.selection.typetext Text:="物理参数统计"
    word.selection.typetext Text:="网络无线参数统计表"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="测量数目： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=str(cc_all)
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.typetext Text:="Rxlev_Full (>=-83dBm)： "
    If stre_s(0) = True Then
       word.selection.Font.Bold = -1
       word.selection.typetext Text:=r_f83
    Else
       'Call fill("26", "200", "22", True, False, False, True)
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="N/A"
    End If
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Rxlev_Full (>=-93dBm)： "
    If stre_s(0) = True Then
       word.selection.Font.Bold = -1
       word.selection.typetext Text:=r_f93
    Else
       'Call fill("16", "200", "22", True, False, False, True)
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="N/A"
    End If
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="RxQual_Full (<=3BER)： "
    If stre_s(2) = True Then
       word.selection.Font.Bold = -1
       word.selection.typetext Text:=rq_f
    Else
       'Call fill("-1", "4", "23", True, False, False, True)
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="N/A"
    End If
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Rxlev_Sub (>=-83dBm)： "
    If stre_s(1) = True Then
       word.selection.Font.Bold = -1
       word.selection.typetext Text:=r_s83
    Else
       'Call fill("26", "200", "24", True, False, False, True)
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="N/A"
    End If
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Rxlev_Sub (>=-93dBm)： "
    If stre_s(1) = True Then
       word.selection.Font.Bold = -1
       word.selection.typetext Text:=r_s93
    Else
       'Call fill("16", "200", "24", True, False, False, True)
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="N/A"
    End If
    
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="RxQual_Sub (<=3BER)： "
    If stre_s(3) = True Then
       word.selection.Font.Bold = -1
       word.selection.typetext Text:=rq_s
    Else
       'Call fill("-1", "4", "25", True, False, False, True)
       word.selection.Font.Bold = -1
       word.selection.typetext Text:="N/A"
    End If
    word.selection.TypeParagraph
    
    If stre_s(4) = True Then
       myb = mta
    Else
       ttnum = 0
       ttall = 0
       tt_ta = 0
       For i = 1 To stre_num
           msg = "select * from " + stre_tab(i) + " where col26 <> " + Chr(34) + Chr(34) + " into temp"
           mapinfo.do msg
           ttnum = mapinfo.eval("tableinfo(temp,8)")
           ttall = ttnum + ttall
           mapinfo.do "select avg(col26) from temp into mytemp"
           tt_ta = mapinfo.eval("mytemp.col1") * ttnum + tt_ta
       Next
       tt_ta = tt_ta / ttall
       myb = Format(tt_ta, "fixed")
    End If
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Mean Timing Advance： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=myb
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Underline = 1
    word.selection.typetext Text:="通话统计"
    word.selection.Font.Underline = 0
    word.selection.TypeParagraph
    setup_n = mess_num("SETUP")
    Call xt_time(False, tmp2_n, True)
'    tmp1_n = mess_num("TMSI REALLOCATION COMMAND")
'    tmp2_n = setup_n - tmp1_n
    If setup_n <> 0 Then
       perce = tmp2_n / setup_n
    End If
    If setup_n = 0 Then
       myb = "N/A"
    Else
       myb = Format(perce, "percent")
    End If
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Successes： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=myb
    word.selection.TypeParagraph
    tmp1_n = mess_num("ALERTING")
    If setup_n <> 0 Then
       perce = tmp1_n / setup_n
    End If
    If setup_n = 0 Then
       myb = "N/A"
    Else
       myb = Format(perce, "percent")
    End If
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Alert Successes： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=myb
    word.selection.TypeParagraph
    tmp1_n = mess_num("CONNECT")
    If setup_n <> 0 Then
       perce = tmp1_n / setup_n
    End If
    If setup_n = 0 Then
       myb = "N/A"
    Else
       myb = Format(perce, "percent")
    End If
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Call Connect Successes： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=myb
    word.selection.TypeParagraph
    tmp1_n = mess_num("LOCATION UPDATING REQUEST")
    tmp2_n = mess_num("LOCATION UPDATING ACCEPT")
    If tmp1_n <> 0 Then
       perce = tmp2_n / tmp1_n
    End If
    If tmp1_n = 0 Then
       myb = "N/A"
    Else
       myb = Format(perce, "percent")
    End If
    word.selection.Font.Bold = 0
    word.selection.typetext Text:="Location Updating Successes： "
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=myb
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Underline = 1
    word.selection.typetext Text:="切换统计"
    word.selection.Font.Underline = 0
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    If stre_s(7) = False Then
       Call hand_time(enum1, enum2, enum3)
       com_hc = enum1
       com_hs = enum2
       com_hf = enum3
       hand_zz
       com_hmax = table_f(9, 2)
       com_havg = table_f(9, 3)
       com_hmin = table_f(9, 4)
       For i = 1 To 9
           For j = 1 To 6
               table_s(i, j) = tibeh(i, j)
           Next
       Next
       hand_zz
       com_qmax = table_f(9, 2)
       com_qavg = table_f(9, 3)
       com_qmin = table_f(9, 4)
    
    End If
    If stre_s(9) = fasle Then
       Call xt_time(False, tt, False)
       hand_zz
       If table_s(9, 1) = 0 Then
          com_xmax = "N/A"
          com_xavg = "N/A"
          com_xmin = "N/A"
       Else
          com_xmax = table_f(9, 2)
          com_xavg = table_f(9, 3)
          com_xmin = table_f(9, 4)
       End If
    
    End If
    If Val(com_hc) = 0 Then
       com_hs = "N/A"
       com_hf = "N/A"
       com_hmax = "N/A"
       com_havg = "N/A"
       com_hmin = "N/A"
       com_qmax = "N/A"
       com_qavg = "N/A"
       com_qmin = "N/A"
    Else
       If Val(com_hc) < 2 Then
          com_qmax = "N/A"
          com_qavg = "N/A"
          com_qmin = "N/A"
       End If
    End If
    Call w_insert("切换发起（Handover Command）： ", com_hc)
    Call w_insert("切换成功（Handover Successes）： ", com_hs)
    Call w_insert("切换失败（Handover Failures）： ", com_hf)
    Call w_insert("最大切换时间（ms）： ", com_hmax)
    Call w_insert("平均切换时间（ms）： ", com_havg)
    Call w_insert("最小切换时间（ms）： ", com_hmin)
    word.selection.TypeParagraph
    word.selection.Font.Bold = -1
    word.selection.Font.Underline = 1
    word.selection.typetext Text:="切换间隔时间统计"
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.Underline = 0
    Call w_insert("最长间隔时间（ms）： ", com_qmax)
    Call w_insert("平均间隔时间（ms）： ", com_qavg)
    Call w_insert("最短间隔时间（ms）： ", com_qmin)
    word.selection.TypeParagraph
    word.selection.Font.Bold = -1
    word.selection.Font.Underline = 1
    word.selection.typetext Text:="系统响应时间"
    word.selection.TypeParagraph
    word.selection.Font.Bold = 0
    word.selection.Font.Underline = 0
    Call w_insert("最长响应时间（ms）： ", com_xmax)
    Call w_insert("平均响应时间（ms）： ", com_xavg)
    Call w_insert("最短响应时间（ms）： ", com_xmin)
    word.selection.TypeParagraph
    word.selection.TypeParagraph
  End If
    
  If stre_s(10) = True Then
    cellci = space$(5)
    oldci = space$(5)
    cellall = mapinfo.eval("tableinfo(cell,8)")
    ReDim Rcellname(1 To cellall) As String
    ReDim dtx(1 To cellall) As String
    ReDim Rep_Ci(1 To cellall) As String * 5
    ReDim lon(1 To cellall) As String
    ReDim lat(1 To cellall) As String
    ReDim down(1 To cellall) As String
    ReDim bearing(1 To cellall) As String
    ReDim No_Ncell(1 To cellall) As String
    ReDim No_Ncell_Percent(1 To cellall) As String
    
'******************************************************************************
    ReDim My_Bearing(1 To cellall) As String
    Dim Point1_Lon As Single, Point1_Lat As Single
    Dim Point2_Lon As Single, Point2_Lat As Single
    Dim Cell_Lon As Single, Cell_Lat As Single
    Dim Point_Lon As Single, Point_Lat As Single
    Dim GetPoint As Boolean, Point_Mark As Boolean
'*****************************************************************************
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.Font.Underline = 0
    word.selection.Font.Bold = -1
    'word.selection.typetext Text:="车载测量报告"
    word.selection.Font.Size = 10.5
    word.selection.ParagraphFormat.Alignment = 1
    word.selection.typetext Text:="街道测量报告"
    word.selection.TypeParagraph
    word.selection.ParagraphFormat.Alignment = 3
    
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Underline = 0
    word.selection.Font.Bold = 0
    word.selection.Font.Size = 9
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(0.65), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(1.42), Alignment:=0, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.66), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(6.98), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(9.38), Alignment:=1, Leader:=0
    'word.selection.ParagraphFormat.TabStops.ADD Position:=word.CentimetersToPoints(11.22), Alignment:=1, Leader:=0
    'word.selection.typetext Text:=Chr(9) & "序号" & Chr(9) & "天线名称" & Chr(9) & "DTX状态" & Chr(9) & "覆盖状况" & Chr(9) & "干扰来源" & Chr(9) & "说明"
    word.selection.typetext Text:=Chr(9) & "序号" & Chr(9) & "天线名称" & Chr(9) & "DTX状态" & Chr(9) & "无邻小区测量数" & Chr(9) & "所占百分比"
    word.selection.TypeParagraph
    RCellNo = 1
    For j = 1 To stre_num
        str_all = mapinfo.eval("tableinfo(" + stre_tab(j) + ",8)")
        mapinfo.do "fetch first from " & stre_tab(j)
        GetPoint = False   '1014
        Point_Mark = False
        For i = 1 To str_all
            mapci = mapinfo.eval(stre_tab(j) + ".col16")
   '        mapci = "&h" & mapci
            mapci = Val(mapci)
            cellci = mapci
           If cellci <> oldci Then
              If RCellNo > 1 Then
                 For p = 1 To RCellNo - 1
                     If Rep_Ci(p) = cellci Then
                        If Point_Mark = True Then
                        
                 Point2_Lon = mapinfo.eval(stre_tab(j) + ".lon")
                 Point2_Lat = mapinfo.eval(stre_tab(j) + ".lat")
                 GetPoint = False
                 Point_Lon = (Abs(Point2_Lon + Point1_Lon)) / 2
                 Point_Lat = (Abs(Point2_Lat + Point1_Lat)) / 2
                 My_Bearing(RCellNo - 1) = Format(180 / 3.1415 * Atn(Abs((Point_Lon - Cell_Lon) / (Point_Lat - Cell_Lat))), "###")
                 If Point_Lon > Cell_Lon And Point_Lat < Cell_Lat Then
                    My_Bearing(RCellNo - 1) = Format(90 - Val(My_Bearing(RCellNo - 1)) + 90, "###")
                 Else
                    If Point_Lon < Cell_Lon And Point_Lat < Cell_Lat Then
                       My_Bearing(RCellNo - 1) = Format(Val(My_Bearing(RCellNo - 1)) + 180, "###")
                    Else
                       If Point_Lon < Cell_Lon And Point_Lat > Cell_Lat Then
                          My_Bearing(RCellNo - 1) = Format(90 - Val(My_Bearing(RCellNo - 1)) + 270, "###")
                       Else
                          My_Bearing(RCellNo - 1) = Format(Val(My_Bearing(RCellNo - 1)), "###")
                       End If
                    End If
                 End If
                        
                        End If
                        GetPoint = False
                        Point_Mark = False
                        GoTo ddt_mov
                     End If
                 Next
              End If
              Rep_Ci(RCellNo) = cellci
              Rcellname(RCellNo) = Findcell(cellci)
              If Rcellname(RCellNo) = "" Then
                 Rcellname(RCellNo) = "CI=" & Rep_Ci(RCellNo)
              Else
                 finds = InStr(Rcellname(RCellNo), Chr(0))
                 If finds > 0 Then
                    Rcellname(RCellNo) = Left(Rcellname(RCellNo), finds - 1)
                 End If
              End If
              dtx(RCellNo) = mapinfo.eval(stre_tab(j) + ".col40")
              lon(RCellNo) = mapinfo.eval("cell.lon")
              lat(RCellNo) = mapinfo.eval("cell.lat")
              down(RCellNo) = mapinfo.eval("cell.col16")
              bearing(RCellNo) = mapinfo.eval("cell.col6")
              word.selection.typetext Text:=Chr(9) + str(RCellNo)
              word.selection.typetext Text:=Chr(9) + Rcellname(RCellNo)
              word.selection.typetext Text:=Chr(9) + dtx(RCellNo)
              word.selection.TypeParagraph
              oldci = cellci
              
              Point_Mark = True
              If GetPoint = False Then   '1014
                 Point1_Lon = mapinfo.eval(stre_tab(j) + ".lon")   '1014
                 Point1_Lat = mapinfo.eval(stre_tab(j) + ".lat")   '1014
                 Cell_Lon = Val(lon(RCellNo))
                 Cell_Lat = Val(lat(RCellNo))
                 GetPoint = True
              Else
                 Point2_Lon = mapinfo.eval(stre_tab(j) + ".lon")
                 Point2_Lat = mapinfo.eval(stre_tab(j) + ".lat")
                 Point_Lon = (Abs(Point2_Lon + Point1_Lon)) / 2
                 Point_Lat = (Abs(Point2_Lat + Point1_Lat)) / 2
                 My_Bearing(RCellNo - 1) = Format(180 / 3.1415 * Atn(Abs((Point_Lon - Cell_Lon) / (Point_Lat - Cell_Lat))), "###")
                 If Point_Lon > Cell_Lon And Point_Lat < Cell_Lat Then
                    My_Bearing(RCellNo - 1) = Format(90 - Val(My_Bearing(RCellNo - 1)) + 90, "###")
                 Else
                    If Point_Lon < Cell_Lon And Point_Lat < Cell_Lat Then
                       My_Bearing(RCellNo - 1) = Format(Val(My_Bearing(RCellNo - 1)) + 180, "###")
                    Else
                       If Point_Lon < Cell_Lon And Point_Lat > Cell_Lat Then
                          My_Bearing(RCellNo - 1) = Format(90 - Val(My_Bearing(RCellNo - 1)) + 270, "###")
                       Else
                          My_Bearing(RCellNo - 1) = Format(Val(My_Bearing(RCellNo - 1)), "###")
                       End If
                    End If
                 End If
                 Point1_Lon = Point2_Lon
                 Point1_Lat = Point2_Lat
                 Cell_Lon = Val(lon(RCellNo))
                 Cell_Lat = Val(lat(RCellNo))
              End If   '1014
              RCellNo = RCellNo + 1
           End If
ddt_mov:
           mapinfo.do "fetch next from " & stre_tab(j)
       Next
       If Point_Mark = True Then
                 Point2_Lon = mapinfo.eval(stre_tab(j) + ".lon")
                 Point2_Lat = mapinfo.eval(stre_tab(j) + ".lat")
                 GetPoint = False
                 Point_Lon = (Abs(Point2_Lon + Point1_Lon)) / 2
                 Point_Lat = (Abs(Point2_Lat + Point1_Lat)) / 2
                 My_Bearing(RCellNo - 1) = Format(180 / 3.1415 * Atn(Abs((Point_Lon - Cell_Lon) / (Point_Lat - Cell_Lat))))
                 If Point_Lon > Cell_Lon And Point_Lat < Cell_Lat Then
                    My_Bearing(RCellNo - 1) = Format(90 - Val(My_Bearing(RCellNo - 1)) + 90, "###")
                 Else
                    If Point_Lon < Cell_Lon And Point_Lat < Cell_Lat Then
                       My_Bearing(RCellNo - 1) = Format(Val(My_Bearing(RCellNo - 1)) + 180, "###")
                    Else
                       If Point_Lon < Cell_Lon And Point_Lat > Cell_Lat Then
                          My_Bearing(RCellNo - 1) = Format(90 - Val(My_Bearing(RCellNo - 1)) + 270, "###")
                       Else
                          My_Bearing(RCellNo - 1) = Format(Val(My_Bearing(RCellNo - 1)), "###")
                       End If
                    End If
                 End If
       End If
    Next
    
    For i = 1 To RCellNo - 1
        NTempNum = 0
        NTempTotal = 0
        For j = 1 To stre_num
            mapinfo.do "select * from " + stre_tab(j) + " where ci_serv = " + Chr(34) + Trim(Rep_Ci(i)) + Chr(34) + " and ncell_num = 0 into temp"
            If Val(mapinfo.eval("tableinfo(temp,8)")) > 0 Then
               NTempNum = NTempNum + Val(mapinfo.eval("tableinfo(temp,8)"))
               mapinfo.do "select * from " + stre_tab(j) + " where ci_serv = " + Chr(34) + Trim(Rep_Ci(i)) + Chr(34) + " into temp"
               NTempTotal = NTempTotal + Val(mapinfo.eval("tableinfo(temp,8)"))
            End If
        Next
        No_Ncell(i) = Format(NTempNum)
        If NTempTotal = 0 Then
           No_Ncell_Percent(i) = "0.00%"
        Else
           No_Ncell_Percent(i) = Format(NTempNum / NTempTotal, "Percent")
        End If
    Next
    word.selection.MoveUp unit:=5, Count:=RCellNo - 1
    word.selection.EndKey unit:=5
    For i = 1 To RCellNo - 1
        word.selection.typetext Text:=Chr(9) & No_Ncell(i) & Chr(9) & No_Ncell_Percent(i)
        word.selection.MoveDown unit:=5, Count:=1
    Next
    
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.Font.Size = 9
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.Font.Bold = -1
    word.selection.typetext Text:="天线测试及整治报告"
    word.selection.Font.Underline = 0
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Size = 9
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(0.65), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(1.42), Alignment:=0, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.13), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(6.15), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(8.67), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(11.04), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(13.7), Alignment:=1, Leader:=0
    word.selection.typetext Text:=Chr(9) & "序号" & Chr(9) & "天线名称" & Chr(9) & "DTX" & Chr(9) & "天线位置" & Chr(9) & "天线下倾角" & Chr(9) & "规划天线方向角" & Chr(9) & "测试天线方向角"
    word.selection.TypeParagraph
    For i = 1 To RCellNo - 1
        word.selection.typetext Text:=Chr(9) + str(i) + Chr(9) + Rcellname(i)
        word.selection.typetext Text:=Chr(9) + dtx(i)
        word.selection.typetext Text:=Chr(9) + lon(i) + "  " + lat(i) + Chr(9) + down(i) + Chr(9) + bearing(i) + Chr(9)
        word.selection.Font.colorindex = 2
        word.selection.typetext Text:=My_Bearing(i)
        word.selection.Font.colorindex = 0
        word.selection.TypeParagraph
    Next
    
    word.selection.TypeParagraph
    
    Call esceed_a(3)
    
    For i = 1 To RCellNo - 1
        Call RxQual_Percent(Rep_Ci(i), Get_Percent)
        word.selection.typetext Text:=Chr(9) + str(i) + Chr(9) + Rcellname(i) + Chr(9) + Get_Percent
        word.selection.TypeParagraph
    Next
    If Cell_Report = False And RCellNo > 2 Then
       word.selection.MoveUp unit:=5, Count:=RCellNo - 1
       word.selection.HomeKey unit:=5
       word.selection.MoveDown unit:=5, Count:=RCellNo - 1, Extend:=1
       word.selection.Sort ExcludeHeader:=False, FieldNumber:="域 4", SortFieldType:=3, SortOrder:=1, FieldNumber2:="", SortFieldType2:=5, SortOrder2:=0, FieldNumber3:="", SortFieldType3:=5, SortOrder3:=0, Separator:=0, SortColumn:=False, CaseSensitive:=False, LanguageID:=0
       word.selection.MoveDown unit:=5, Count:=1
    End If
    word.selection.Font.colorindex = 15
    word.selection.typetext Text:=Chr(9) + Chr(9) + Chr(9) + Chr(9) + "1. 上基站查看" + Chr(9) + "1. S1 天线断裂" + Chr(9) + "1. 调整下倾角"
    word.selection.TypeParagraph
    word.selection.typetext Text:=Chr(9) + Chr(9) + Chr(9) + Chr(9) + "2. 频率重新安排" + Chr(9) + "2. 来自澳门的干扰" + Chr(9) + "2. 调整基站位置"
    word.selection.TypeParagraph
    word.selection.typetext Text:=Chr(9) + Chr(9) + Chr(9) + Chr(9) + "3. 重新做邻小区表" + Chr(9) + "3. 覆盖太小" + Chr(9) + "3. 改变天线方向"
    word.selection.TypeParagraph
    word.selection.typetext Text:=Chr(9) + "范例" + Chr(9) + "九洲港" + Chr(9) + Chr(9) + "4. 减小 TX 功率" + Chr(9) + "4. 信号过界" + Chr(9) + "4. 增加一扇天线"
    word.selection.TypeParagraph
    word.selection.typetext Text:=Chr(9) + Chr(9) + Chr(9) + Chr(9) + "5. 检查天线损耗" + Chr(9) + "5. 天线太低" + Chr(9) + "5. 增加直放站"
    word.selection.TypeParagraph
    word.selection.typetext Text:=Chr(9) + Chr(9) + Chr(9) + Chr(9) + "6. 换掉 RTC 组合器" + Chr(9) + "6. 建筑物遮挡" + Chr(9) + "6. 增加微蜂窝"
    word.selection.TypeParagraph
    word.selection.typetext Text:=Chr(9) + Chr(9) + Chr(9) + Chr(9) + "7. 检查 Rx/Tx 天线" + Chr(9) + "7. 无下倾角"
    word.selection.TypeParagraph
    word.selection.Font.Size = 9
    word.selection.Font.colorindex = 0
    word.selection.TypeParagraph
 End If
 If stre_s(11) = True Then
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.Font.Size = 9
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.Font.Bold = -1
    word.selection.typetext Text:="同频干扰统计"
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Underline = 0
    word.selection.Font.Bold = 0
    word.selection.Font.Size = 9
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(0.65), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(1.42), Alignment:=0, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.76), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(7.8), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(9.84), Alignment:=0, Leader:=0
    word.selection.typetext Text:=Chr(9) & "序号" & Chr(9) & "天线名称" & Chr(9) & "受干扰测量数" & Chr(9) & "占天线测量数百分比" & Chr(9) & "干扰来源"
    word.selection.TypeParagraph
    If stre_s(10) = False Then
       cellci = space$(5)
       oldci = space$(5)
       cellall = mapinfo.eval("tableinfo(cell,8)")
       ReDim Rcellname(1 To cellall) As String
       ReDim dtx(1 To cellall) As String
       ReDim Rep_Ci(1 To cellall) As String * 5
       RCellNo = 1
       For j = 1 To stre_num
           str_all = mapinfo.eval("tableinfo(" + stre_tab(j) + ",8)")
           mapinfo.do "fetch first from " & stre_tab(j)
           For i = 1 To str_all
               cellci = mapinfo.eval(stre_tab(j) + ".col16")
               If cellci <> oldci Then
                  If RCellNo > 1 Then
                     For p = 1 To RCellNo - 1
                         If Rep_Ci(p) = cellci Then
                            GoTo arfcn_mov
                         End If
                     Next
                  End If
                  Rep_Ci(RCellNo) = cellci
                  Rcellname(RCellNo) = Findcell(cellci)
                  If Rcellname(RCellNo) = "" Then
                     Rcellname(RCellNo) = "CI=" & Rep_Ci(RCellNo)
                  Else
                     finds = InStr(Rcellname(RCellNo), Chr(0))
                     If finds > 0 Then
                        Rcellname(RCellNo) = Left(Rcellname(RCellNo), finds - 1)
                     End If
                  End If
                  oldci = cellci
                  RCellNo = RCellNo + 1
               End If
arfcn_mov:
               mapinfo.do "fetch next from " & stre_tab(j)
           Next
       Next
    End If
    For i = 1 To RCellNo - 1
        ArfcnBsicNum = 0
        Rtotal_Num = 0
        For j = 1 To 10
            Rselect_Num(j) = 0
            My_Source(j) = ""
            CellArfcn(j) = ""
            CellBsic(j) = ""
        Next
        OldArfcn = ""
        OldBsic = ""
        For j = 1 To stre_num
            For k = 1 To 6
                If Report_Full = True Then
                   mapinfo.do "select * from " + stre_tab(j) + " where (bcch_serv > 0 ) and (ci_serv = " + Chr(34) + Trim(Rep_Ci(i)) + Chr(34) + ") and (bcch_serv= Bcch_N" & k & ") AND (abs(RXLEV_F-Rxlev_n" & k & ") < " & Report_Rxlev1 & ") into my_temp"
                Else
                   mapinfo.do "select * from " + stre_tab(j) + " where (bcch_serv > 0 ) and (ci_serv = " + Chr(34) + Trim(Rep_Ci(i)) + Chr(34) + ") and (bcch_serv= Bcch_N" & k & ") AND (abs(RXLEV_s-Rxlev_n" & k & ") < " & Report_Rxlev1 & ") into my_temp"
                End If
                temp_all = Val(mapinfo.eval("tableinfo(my_temp,8)"))
                mapinfo.do "fetch first from my_temp"
                For p = 1 To temp_all
                    If Trim(mapinfo.eval("my_temp.Bcch_n" & j)) <> Trim(OldArfcn) And Trim(mapinfo.eval("my_temp.Rxlev_n" & j)) <> Trim(OldBsic) Then
                       If ArfcnBsicNum > 0 Then
                          For q = 1 To ArfcnBsicNum
                              If Trim(mapinfo.eval("my_temp.Bcch_n" & j)) = Trim(CellArfcn(q)) And Trim(mapinfo.eval("my_temp.Rxlev_n" & j)) = Trim(CellBsic(q)) Then
                                 Rselect_Num(q) = Rselect_Num(q) + 1
                                 GoTo NextRecord
                              End If
                          Next
                       End If
                       ArfcnBsicNum = ArfcnBsicNum + 1
                       Rselect_Num(ArfcnBsicNum) = Rselect_Num(ArfcnBsicNum) + 1
                       CellArfcn(ArfcnBsicNum) = Trim(mapinfo.eval("my_temp.Bcch_n" & j))
                       CellBsic(ArfcnBsicNum) = Trim(mapinfo.eval("my_temp.bsic_n" & j))
                       My_Source(ArfcnBsicNum) = FindSource(CellArfcn(ArfcnBsicNum), CellBsic(ArfcnBsicNum))
                       OldArfcn = CellArfcn(ArfcnBsicNum)
                       OldBsic = CellBsic(ArfcnBsicNum)
                    Else
                       Rselect_Num(ArfcnBsicNum) = Rselect_Num(ArfcnBsicNum) + 1
                    End If
NextRecord:
                    mapinfo.do "fetch next from my_temp"
                Next
                
            Next
            mapinfo.do "select * from " + stre_tab(j) + " where ci_serv = " + Chr(34) + Trim(Rep_Ci(j)) + Chr(34) + " into temp"
            Rtotal_Num = Rtotal_Num + Val(mapinfo.eval("tableinfo(temp,8)"))
        Next
        word.selection.typetext Text:=Chr(9) & Format(i) & Chr(9) & Rcellname(i) & Chr(9) & Format(Rselect_Num(1)) & Chr(9) & Format(Rselect_Num(1) / Rtotal_Num, "percent") & Chr(9) & My_Source(1)
        word.selection.TypeParagraph
        If ArfcnBsicNum > 1 Then
           For j = 2 To ArfcnBsicNum
               word.selection.typetext Text:=Chr(9) & Chr(9) & Chr(9) & Format(Rselect_Num(j)) & Chr(9) & Format(Rselect_Num(j) / Rtotal_Num, "percent") & Chr(9) & My_Source(j)
               word.selection.TypeParagraph
           Next
        End If
    Next
    
    'word.selection.MoveUp Unit:=5, Count:=RCellNo - 1
    'word.selection.HomeKey Unit:=5
    'word.selection.MoveDown Unit:=5, Count:=RCellNo - 1, Extend:=1
    'word.selection.Sort ExcludeHeader:=False, FieldNumber:="域 5", SortFieldType:=3, SortOrder:=1, FieldNumber2:="", SortFieldType2:=5, SortOrder2:=0, FieldNumber3:="", SortFieldType3:=5, SortOrder3:=0, Separator:=0, SortColumn:=False, CaseSensitive:=False, LanguageID:=0
    'word.selection.MoveDown Unit:=5, Count:=1
    word.selection.TypeParagraph
    
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
    word.selection.Font.Bold = -1
    word.selection.typetext Text:="邻频干扰统计"
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Underline = 0
    word.selection.Font.Bold = 0
    word.selection.Font.Size = 9
    word.selection.typetext Text:=Chr(9) & "序号" & Chr(9) & "天线名称" & Chr(9) & "受干扰测量数" & Chr(9) & "占天线测量数百分比" & Chr(9) & "干扰来源"
    word.selection.TypeParagraph
    For i = 1 To RCellNo - 1
        ArfcnBsicNum = 0
        Rtotal_Num = 0
        For j = 1 To 10
            Rselect_Num(j) = 0
            My_Source(j) = ""
            CellArfcn(j) = ""
            CellBsic(j) = ""
        Next
        OldArfcn = ""
        OldBsic = ""
        For j = 1 To stre_num
            For k = 1 To 6
                If Report_Full = True Then
                   mapinfo.do "select * from " + stre_tab(j) + " where (bcch_serv > 0 ) and (ci_serv = " + Chr(34) + Trim(Rep_Ci(i)) + Chr(34) + ") and (abs(bcch_serv- Bcch_N" & k & ")=1) AND (abs(RXLEV_F-Rxlev_n" & k & ") < " & Report_Rxlev2 & ") into my_temp"
                Else
                   mapinfo.do "select * from " + stre_tab(j) + " where (bcch_serv > 0 ) and (ci_serv = " + Chr(34) + Trim(Rep_Ci(i)) + Chr(34) + ") and (abs(bcch_serv= Bcch_N" & k & ")=1) AND (abs(RXLEV_s-Rxlev_n" & k & ") < " & Report_Rxlev2 & ") into my_temp"
                End If
                temp_all = Val(mapinfo.eval("tableinfo(my_temp,8)"))
                mapinfo.do "fetch first from my_temp"
                For p = 1 To temp_all
                    If Trim(mapinfo.eval("my_temp.Bcch_n" & j)) <> Trim(OldArfcn) And Trim(mapinfo.eval("my_temp.Rxlev_n" & j)) <> Trim(OldBsic) Then
                       If ArfcnBsicNum > 0 Then
                          For q = 1 To ArfcnBsicNum
                              If Trim(mapinfo.eval("my_temp.Bcch_n" & j)) = Trim(CellArfcn(q)) And Trim(mapinfo.eval("my_temp.Rxlev_n" & j)) = Trim(CellBsic(q)) Then
                                 Rselect_Num(q) = Rselect_Num(q) + 1
                                 GoTo NextRecord2
                              End If
                          Next
                       End If
                       ArfcnBsicNum = ArfcnBsicNum + 1
                       Rselect_Num(ArfcnBsicNum) = Rselect_Num(ArfcnBsicNum) + 1
                       CellArfcn(ArfcnBsicNum) = Trim(mapinfo.eval("my_temp.Bcch_n" & j))
                       CellBsic(ArfcnBsicNum) = Trim(mapinfo.eval("my_temp.bsic_n" & j))
                       My_Source(ArfcnBsicNum) = FindSource(CellArfcn(ArfcnBsicNum), CellBsic(ArfcnBsicNum))
                       OldArfcn = CellArfcn(ArfcnBsicNum)
                       OldBsic = CellBsic(ArfcnBsicNum)
                    Else
                       Rselect_Num(ArfcnBsicNum) = Rselect_Num(ArfcnBsicNum) + 1
                    End If
NextRecord2:
                    mapinfo.do "fetch next from my_temp"
                Next
                
            Next
            mapinfo.do "select * from " + stre_tab(j) + " where ci_serv = " + Chr(34) + Trim(Rep_Ci(j)) + Chr(34) + " into temp"
            Rtotal_Num = Rtotal_Num + Val(mapinfo.eval("tableinfo(temp,8)"))
        Next
        word.selection.typetext Text:=Chr(9) & Format(i) & Chr(9) & Rcellname(i) & Chr(9) & Format(Rselect_Num(1)) & Chr(9) & Format(Rselect_Num(1) / Rtotal_Num, "percent") & Chr(9) & My_Source(1)
        word.selection.TypeParagraph
        If ArfcnBsicNum > 1 Then
           For j = 2 To ArfcnBsicNum
               word.selection.typetext Text:=Chr(9) & Chr(9) & Chr(9) & Format(Rselect_Num(j)) & Chr(9) & Format(Rselect_Num(j) / Rtotal_Num, "percent") & Chr(9) & My_Source(j)
               word.selection.TypeParagraph
           Next
        End If
    Next
    word.selection.TypeParagraph
    word.selection.TypeParagraph
 
 End If
 
    On Error Resume Next
    word.selection.Font.Size = 9
    If word.ActiveWindow.View.SplitSpecial <> 0 Then
       word.ActiveWindow.Panes(2).Close
    End If
    If word.ActiveWindow.ActivePane.View.type = 1 Or word.ActiveWindow.ActivePane.View.type = 2 Or word.ActiveWindow.ActivePane.View.type = 5 Then
        word.ActiveWindow.ActivePane.View.type = 3
    End If
    word.ActiveWindow.ActivePane.View.SeekView = 9
    
    Gsm_FileName = Gsm_Path + "\bmp\ant_sign.bmp"
    word.selection.InlineShapes.AddPicture filename:=Gsm_FileName, LinkToFile:=True, SaveWithDocument:=True
    word.selection.ParagraphFormat.Alignment = 2
    
    If word.selection.HeaderFooter.IsHeader = True Then
       word.ActiveWindow.ActivePane.View.SeekView = 10
    Else
       word.ActiveWindow.ActivePane.View.SeekView = 9
    End If
    
    word.selection.HeaderFooter.Shapes.AddLine(90#, 797.4, 504#, 797.4).Select
    word.selection.MoveDown unit:=5, Count:=3
    word.selection.TypeParagraph
    word.selection.HeaderFooter.Shapes("line 1").Select
    word.selection.ShapeRange.IncrementTop -8
    word.selection.MoveDown unit:=5, Count:=3
    word.selection.typetext Text:=USERNAME
    
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(7.32), Alignment:=3, Leader:=0
    
    word.selection.typetext Text:=Chr(9)
    word.selection.Fields.Add Range:=word.selection.Range, type:=33
    word.ActiveWindow.ActivePane.View.SeekView = 0
    
    'word.selection.typetext Text:=USERNAME
    'word.ActiveWindow.ActivePane.View.SeekView = 0
    'word.selection.Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:=2, FirstPage:=True
    On Error Resume Next
    word.ChangeFileOpenDirectory Gsm_Path + "\user\"
    word.ActiveDocument.SaveAs

    If Cell_Report = True Then
       For i = 1 To stre_num
           mapinfo.do "close table " & stre_save(i)
       Next
    Else
       For i = 1 To stre_num
           mapinfo.do "close table " & stre_tab(i)
       Next
    End If
    mapinfo.do "close table cell"
End Sub

Sub fill(a As String, b As String, col As String, ByVal sta As Boolean, ByVal rxq As Boolean, ByVal va As Boolean, ByVal x9 As Boolean)
    Dim num, er_max, er_avg, er_min, er3, er4, er
    Dim num_z As Integer, max_z As Integer, avg_z As Single, min_z As Integer
    Dim msg As String
    Dim zero As Boolean
    Dim Is_Exist As Boolean
    Static perc
    On Error Resume Next
    num_z = 0
    zero = True
    max_z = 0
    If sta = True Then
       all_max = 0
       all_min = 0
       all_avg = 0
       perc = 0
       ta_all = 0
    End If
    Is_Exist = False
    If perc = 1 Then
       word.selection.MoveDown unit:=5, Count:=1
       Exit Sub
    End If
    
    For pp = 1 To stre_num
        If rxq = True Then
           If va = True Then
              msg = "select * from " + stre_tab(pp) + " where val(col" + col + ") = " + a + " into temp"
           Else
              msg = "select * from " + stre_tab(pp) + " where col" + col + " = " + a + " into temp"
           End If
        Else
           If va = True Then
              msg = "select * from " + stre_tab(pp) + " where col" + col + " <> " + Chr(34) + Chr(34) + " into tep"
              mapinfo.do msg
              If sta = True Then
                 ta_all = mapinfo.eval("tableinfo(tep,8)") + ta_all
              End If
              If b = "0" Then
                 msg = "select * from tep where val(col" + col + ") = " + a + " into temp"
              Else
                 msg = "select * from tep where val(col" + col + ") > " + a + " and val(col" + col + ") < " + b + " into temp"
              End If
'              msg = "select * from " + stre_tab(pp) + " where val(col" + col + ") > " + a + " and val(col" + col + ") < " + b + " into temp"
           Else
              msg = "select * from " + stre_tab(pp) + " where col" + col + " > " + a + " and col" + col + " < " + b + " into temp"
           End If
        End If
        mapinfo.do msg
        num = mapinfo.eval("tableinfo(temp,8)")
        num_z = num_z + num
     If Val(num) > 0 Then
        If rxq = False Then
           msg = "select max(col" + col + ") from temp into mytemp"
           mapinfo.do msg
           er_max = mapinfo.eval("mytemp.col1")
           mapinfo.do "select avg(col" + col + ") from temp into mytemp"
           er_avg = mapinfo.eval("mytemp.col1")
           mapinfo.do "select min(col" + col + ") from temp into mytemp"
           er_min = mapinfo.eval("mytemp.col1")
        Else
           er_max = Val(a)
           er_avg = Val(a)
           er_min = Val(a)
        End If
           If zero = True Then
              max_z = er_max
              avg_z = Val(er_avg) * num
              min_z = er_min
           Else
              If max_z < er_max Then max_z = er_max
'              avg_z = (avg_z + Val(er_avg)) / 2
              If min_z > er_min Then min_z = er_min
              avg_z = Val(er_avg) * num + avg_z
           End If
           zero = False
           
     End If
    Next
    If zero = False Then
       If perc = 0 Then
          all_max = max_z
          all_min = min_z
          all_avg = avg_z '* num
       Else
          If all_max < max_z Then all_max = max_z
          If all_min > min_z Then all_min = min_z
'          all_avg = (avg_z + all_avg) / 2
          all_avg = avg_z + all_avg
       End If
       num = LTrim$(str(num_z))
       er_max = LTrim$(str(max_z))
       er_avg = LTrim$(str(avg_z / num_z))
       er_min = LTrim$(str(min_z))
       er_avg = Format(er_avg, "fixed")
       er3 = Format(num_z / cc_all, "percent")
       er4 = Format(num_z / cc_all + perc, "percent")
       perc = num_z / cc_all + perc
       Is_Exist = True
    End If
    If x9 = True Then
       word.selection.Font.Bold = 0
       word.selection.typetext Text:=er3
    Else
       If zero = True Then
          word.selection.MoveDown unit:=5, Count:=1
       Else
          word.selection.EndKey unit:=5
          word.selection.Font.Bold = 0
          If rxq = True Then
             word.selection.typetext Text:=Chr(9) + num + Chr(9) + er_avg + Chr(9) + er3 + Chr(9) + er4
          Else
             word.selection.typetext Text:=Chr(9) + num + Chr(9) + er_max + Chr(9) + er_avg + Chr(9) + er_min + Chr(9) + er3 + Chr(9) + er4
          End If
          word.selection.MoveDown unit:=5, Count:=1
       End If
    End If
    If Is_Exist = False Then
       m_per = Format(perc, "Percent")
    Else
       m_per = er4
    End If
End Sub


Sub rxlev_fuc()
    On Error Resume Next
    word.selection.Font.Size = 9
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.5), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(6.2), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(7.91), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(9.59), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(11.32), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(13.33), Alignment:=1, Leader:=0
    word.selection.typetext Text:=Chr(9) & "测量数" & Chr(9) & "最大值" & Chr(9) & "平均值" & Chr(9) & "最小值" & Chr(9) & "百分比" & Chr(9) & "累计百分比"
    word.selection.TypeParagraph
    word.selection.typetext Text:="57-63 (-53<=dBm<-47)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="47-57 (-63<=dBm<-53)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="37-47 (-73<=dBm<-63)"
    word.selection.TypeParagraph
    word.selection.Font.colorindex = 2
    word.selection.typetext Text:="27-37 (-83<=dBm<-73)"
    word.selection.TypeParagraph
    word.selection.Font.colorindex = 6
    word.selection.typetext Text:="17-27 (-93<=dBm<-83)"
    word.selection.TypeParagraph
    word.selection.Font.colorindex = 0
    word.selection.typetext Text:="7-17 (-103<=dBm<-93)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="0-7 (-110<=dBm<-103)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="总计"
    word.selection.TypeParagraph

End Sub

Sub rxqual_fuc()
    On Error Resume Next
    word.selection.Font.Size = 9
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.23), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(6.14), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(8.04), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(10.16), Alignment:=1, Leader:=0
    word.selection.typetext Text:=Chr(9) & "测量数" & Chr(9) & "平均值" & Chr(9) & "百分比" & Chr(9) & "累计百分比"
    word.selection.TypeParagraph
    word.selection.typetext Text:="0 (BER<0.2%)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="1 (0.2%<BER<0.4%)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="2 (0.4%<BER<0.8%)"
    word.selection.TypeParagraph
    word.selection.Font.colorindex = 6
    word.selection.typetext Text:="3 (0.8%<BER<1.6%)"
    word.selection.TypeParagraph
    word.selection.Font.colorindex = 0
    word.selection.typetext Text:="4 (1.6%<BER<3.2%)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="5 (3.2%<BER<6.4%)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="6 (6.4%<BER<12.8%)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="7 (12.8%<BER)"
    word.selection.TypeParagraph
    word.selection.typetext Text:="总计"

End Sub



Function lencell(cellname) As Integer
    Dim i As Integer
    DoEvents
    If Len(cellname) = 0 Then
       lencell = 0
    Else
       lencell = Len(cellname)
       For i = 1 To Len(cellname)
           nc = Mid(cellname, i, 1)
           If Asc(nc) > 255 Or Asc(nc) < 0 Then
              lencell = lencell + 1
           End If
       Next
    End If
End Function

Function mess_num(ByVal tabna As String) As Integer
    Dim msg As String
    Dim menum
    On Error GoTo errend
    DoEvents
    menum = 0
    For i = 1 To stre_num
        msg = "select * from " + stre_tab(i) + " where col5 = " + Chr(34) + tabna + Chr(34) + " into temp"
        mapinfo.do msg
        DoEvents
        menum = mapinfo.eval("tableinfo(temp,8)") + menum
    Next
    mess_num = menum
errend:
End Function

Sub hand_time(getnum1, getnum2, getnum3)
    Dim h_comm() As Single, h_comp() As Single, h_fail() As Single
    Dim tp(1 To 1000) As Single
    Dim menum, tt, all
    Dim dd As String
    Dim hav As Boolean
    Dim ggg As Single, sss As Single
    Dim mid_dd As Long
    Dim abc As String
    Dim ssi As Integer
    ReDim h_comm(1 To 1000) As Single
    ReDim h_comp(1 To 1000) As Single
    ReDim h_fail(1 To 1000) As Single
    On Error Resume Next
    DoEvents
    menum = 0
    getnum1 = 0
    getnum2 = 0
    getnum3 = 0
    For j = 1 To 9
        For k = 1 To 6
            table_s(j, k) = 0
            table_f(j, k) = ""
        Next k
    Next j
    
    For hh = 1 To stre_num
        For pp = 1 To 1000
            h_comm(pp) = 0
            h_comp(pp) = 0
            h_fail(pp) = 0
        Next
    For qq = 1 To 3
        DoEvents
        If qq = 1 Then abc = "HANDOVER COMMAND"
        If qq = 2 Then abc = "HANDOVER COMPLETE"
        If qq = 3 Then abc = "HANDOVER FAILURE"
        mapinfo.do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + abc + Chr(34) + "into temp order by col1"
        menum = mapinfo.eval("tableinfo(temp,8)")
        If qq = 1 Then getnum1 = getnum1 + menum
        If qq = 2 Then getnum2 = getnum2 + menum
        If qq = 3 Then getnum3 = getnum3 + menum
        
        If menum > 0 Then
           mapinfo.do "fetch first from temp"
           Select Case abc
               Case "HANDOVER COMMAND":
                    For i = 1 To menum
                        tt = mapinfo.eval("temp.col1")
                        dd = tt
                        If Len(dd) <= 8 Then GoTo nohourw
                        finds = InStr(dd, ":")
                        mid_dd = Val(Left(dd, finds - 1)) * 3600
                        h_comm(i) = mid_dd
                        dd = Right(dd, Len(dd) - finds)
nohourw:
                        finds = InStr(dd, ":")
                        mid_dd = Val(Left(dd, finds - 1)) * 60
                        h_comm(i) = h_comm(i) + mid_dd
                        dd = Right(dd, Len(dd) - finds)
                        finds = InStr(dd, ".")
                        mid_dd = Val(Left(dd, finds - 1))
                        h_comm(i) = h_comm(i) + mid_dd
                        dd = Right(dd, Len(dd) - finds)
                        h_comm(i) = h_comm(i) + Val(dd) / 100
                        If i < menum Then
                           mapinfo.do "fetch next from temp"
                        End If
                    Next
'                    h_comm(i) = -1
               Case "HANDOVER COMPLETE":
                    For i = 1 To menum
                        tt = mapinfo.eval("temp.col1")
                        dd = tt
                        If Len(dd) <= 8 Then GoTo nohour2
                        finds = InStr(dd, ":")
                        mid_dd = Val(Left(dd, finds - 1)) * 3600
                        h_comp(i) = mid_dd
                        dd = Right(dd, Len(dd) - finds)
nohour2:
                        finds = InStr(dd, ":")
                        mid_dd = Val(Left(dd, finds - 1)) * 60
                        h_comp(i) = h_comp(i) + mid_dd
                        dd = Right(dd, Len(dd) - finds)
                        finds = InStr(dd, ".")
                        mid_dd = Val(Left(dd, finds - 1))
                        h_comp(i) = h_comp(i) + mid_dd
                        dd = Right(dd, Len(dd) - finds)
                        h_comp(i) = h_comp(i) + Val(dd) / 100
                        If i < menum Then
                           mapinfo.do "fetch next from temp"
                        End If
                    Next
 '                   h_comp(i) = -1
               Case "HANDOVER FAILURE":
                     For i = 1 To menum
                         tt = mapinfo.eval("temp.col1")
                         dd = tt
                         If Len(dd) <= 8 Then GoTo nohour3
                         finds = InStr(dd, ":")
                         mid_dd = Val(Left(dd, finds - 1)) * 3600
                         h_fail(i) = mid_dd
                         dd = Right(dd, Len(dd) - finds)
nohour3:
                         finds = InStr(dd, ":")
                         mid_dd = Val(Left(dd, finds - 1)) * 60
                         h_fail(i) = h_fail(i) + mid_dd
                         dd = Right(dd, Len(dd) - finds)
                         finds = InStr(dd, ".")
                         mid_dd = Val(Left(dd, finds - 1))
                         h_fail(i) = h_fail(i) + mid_dd
                         dd = Right(dd, Len(dd) - finds)
                         h_fail(i) = h_fail(i) + Val(dd) / 100
                         If i < menum Then
                            mapinfo.do "fetch next from temp"
                         End If
                     Next
  '                   h_fail(i) = -1
           End Select
        End If
    Next qq
    
    i = 1
    j = 1
    k = 1
    For bbb = 1 To getnum1
        If h_fail(j) <> 0 Then
           If h_comp(i) < h_fail(j) Then
              tp(k) = h_comp(i)
              k = k + 1
              i = i + 1
           Else
              If h_comp(i) = h_fail(j) Then
                 tp(k) = h_comp(i)
                 k = k + 1
                 i = i + 1
                 j = j + 1
              Else
                 tp(k) = h_fail(j)
                 k = k + 1
                 j = j + 1
              End If
           End If
        Else
           tp(k) = h_comp(i)
           k = k + 1
           i = i + 1
        End If
    Next
    If j - 1 < getnum3 Then
       For M = j To getnum3
           tp(k) = h_fail(M)
           k = k + 1
           M = M + 1
       Next
    End If
    
           i = 1
           j = 1
           k = 1
           tal = 0
           ssi = 1
           Do While h_comm(i) <> 0
           
              If h_comm(ssi + 1) <> 0 Then
                 sss = h_comm(ssi + 1) - h_comm(ssi)
                 ssi = ssi + 1
                 Select Case sss
                     Case Is >= 120
                          tber = 6
                     Case Is >= 10
                          tber = 5
                     Case Is >= 4
                          tber = 4
                     Case Is >= 2
                          tber = 3
                     Case Is >= 1
                          tber = 2
                     Case Is >= 0
                          tber = 1
                 End Select
                 If tibeh(tber, 1) = 0 Then
                    tibeh(tber, 1) = 1
                    tibeh(tber, 2) = sss
                    tibeh(tber, 3) = sss
                    tibeh(tber, 4) = sss
                 Else
                    tibeh(tber, 1) = tibeh(tber, 1) + 1
                    If tibeh(tber, 2) < sss Then tibeh(tber, 2) = sss
                    If tibeh(tber, 4) > sss Then tibeh(tber, 4) = sss
                       tibeh(tber, 3) = (tibeh(tber, 3) + sss) / 2
                 End If
              End If
              
              hav = False
              If h_comm(i) <= tp(j) Then
                 If h_comm(i + 1) > tp(j) Or h_comm(i + 1) = 0 Then
                    ggg = tp(j) - h_comm(i)
                    j = j + 1
                    i = i + 1
                    hav = True
                 Else
                    i = i + 1
                    If h_comm(i) = 0 Then Exit Do
                 End If
              Else
                 j = j + 1
                 If tp(j) = 0 Then Exit Do
              End If
                 
              If hav = True Then
                 Select Case ggg
                     Case Is >= 5
                          tal = 8
                     Case Is >= 2
                          tal = 7
                     Case Is >= 1
                          tal = 6
                     Case Is >= 0.5
                          tal = 5
                     Case Is >= 0.3
                          tal = 4
                     Case Is >= 0.2
                          tal = 3
                     Case Is >= 0.1
                          tal = 2
                     Case Is >= 0
                          tal = 1
                 End Select
                 
                 If table_s(tal, 1) = 0 Then
                    table_s(tal, 1) = 1
                    table_s(tal, 2) = ggg
                    table_s(tal, 3) = ggg
                    table_s(tal, 4) = ggg
                 Else
                    table_s(tal, 1) = table_s(tal, 1) + 1
                    If table_s(tal, 2) < ggg Then table_s(tal, 2) = ggg
                    If table_s(tal, 4) > ggg Then table_s(tal, 4) = ggg
                       table_s(tal, 3) = (table_s(tal, 3) + ggg) / 2
                 End If
              End If
           Loop
    
    Next hh
errend:
End Sub

Sub hand_zz()
   DoEvents
   On Error GoTo errend
           For j = 1 To 8
               table_s(9, 1) = table_s(9, 1) + table_s(j, 1)
           Next
           tawri = False
           For j = 1 To 8
               If table_s(j, 1) > 0 Then
                  table_s(j, 5) = table_s(j, 1) / table_s(9, 1)
                  If j = 1 Or tawri = False Then
                     table_s(j, 6) = table_s(j, 5)
                  Else
                     table_s(j, 6) = table_s(j, 5) + table_s(j - 1, 6)
                  End If
                  If tawri = False Then
                     tawri = True
                     table_s(9, 2) = table_s(j, 2)
                  '   table_s(9, 3) = table_s(j, 3)
                     table_s(9, 4) = table_s(j, 4)
                  Else
                     If table_s(9, 2) < table_s(j, 2) Then table_s(9, 2) = table_s(j, 2)
                     If table_s(9, 4) > table_s(j, 4) Then table_s(9, 4) = table_s(j, 4)
'                     table_s(9, 3) = (table_s(9, 3) + table_s(j, 3)) / 2
                  End If
               Else
                  table_s(j, 5) = 0
                  If j > 1 Then
                     table_s(j, 6) = table_s(j - 1, 6)
                  Else
                     table_s(j, 6) = 0
                  End If
               End If
           Next
           DoEvents

           If table_s(9, 1) > 0 Then
              table_s(9, 3) = table_s(1, 3) * table_s(1, 1) + table_s(2, 3) * table_s(2, 1) + table_s(3, 3) * table_s(3, 1)
              table_s(9, 3) = table_s(9, 3) + table_s(4, 3) * table_s(4, 1) + table_s(5, 3) * table_s(5, 1) + table_s(6, 1) * table_s(6, 3)
              table_s(9, 3) = (table_s(9, 3) + table_s(7, 3) * table_s(7, 1) + table_s(8, 3) * table_s(8, 1)) / table_s(9, 1)
           Else
              table_s(9, 3) = 0
           End If
              
                       
           For i = 1 To 9
               table_s(i, 2) = table_s(i, 2) * 1000
               table_s(i, 3) = table_s(i, 3) * 1000
               table_s(i, 4) = table_s(i, 4) * 1000
               
               table_f(i, 1) = str(table_s(i, 1))
               table_f(i, 2) = Format(table_s(i, 2), "fixed")
               table_f(i, 3) = Format(table_s(i, 3), "fixed")
               table_f(i, 4) = Format(table_s(i, 4), "fixed")
               table_f(i, 5) = Format(table_s(i, 5), "percent")
               table_f(i, 6) = Format(table_s(i, 6), "percent")
            Next
        
errend:
End Sub
Sub wor_fi(ByVal gg As String, ByVal kk As Integer)
    On Error Resume Next
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=gg
    If table_s(kk, 1) = 0 Then GoTo tr
    If kk <> 9 Then
       word.selection.Font.Bold = 0
    End If
    word.selection.typetext Text:=Chr(9) + table_f(kk, 1) + Chr(9) + table_f(kk, 2) + Chr(9)
    word.selection.typetext Text:=table_f(kk, 3) + Chr(9) + table_f(kk, 4)
    If kk = 9 Then
       GoTo tr
    Else
       word.selection.typetext Text:=Chr(9) + table_f(kk, 5) + Chr(9) + table_f(kk, 6)
    End If
tr:
    word.selection.TypeParagraph
End Sub


Sub xt_time(ByVal hu As Boolean, assnum, ByVal hg As Boolean)
    Dim h_comm() As Single, h_comp() As Single
    Dim menum, tt, all
    Dim dd As String
    Dim hav As Boolean
    Dim ggg As Single
    Dim mid_dd As Long
    Dim i As Integer, j As Integer, assi As Integer, assj As Integer
'    all = mapinfo.eval("tableinfo(cell,8)")
    ReDim h_comm(1 To 1000) As Single
    ReDim h_comp(1 To 1000) As Single
    On Error GoTo errend
    DoEvents
    menum = 0
    assnum = 0
    For j = 1 To 9
        For k = 1 To 6
            table_s(j, k) = 0
        Next k
    Next j
    
    For hh = 1 To stre_num
        For pp = 1 To 1000
            h_comm(pp) = 0
            h_comp(pp) = 0
        Next
        If hu = False Then
           If hg = True Then
              mapinfo.do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + "CONNECT" + Chr(34) + "into temp order by col1"
           Else
              mapinfo.do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + "CHANNEL REQUEST" + Chr(34) + "into temp order by col1"
           End If
        Else
           mapinfo.do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + "SETUP" + Chr(34) + "into temp order by col1"
        End If
        DoEvents
        menum = mapinfo.eval("tableinfo(temp,8)")
        
        
        
        If menum > 0 Then
           mapinfo.do "fetch first from temp"
           For i = 1 To menum
               tt = mapinfo.eval("temp.col1")
               dd = tt
               If Len(dd) <= 8 Then GoTo nohourw
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 3600
               h_comm(i) = mid_dd
               dd = Right(dd, Len(dd) - finds)
nohourw:
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 60
               h_comm(i) = h_comm(i) + mid_dd
               dd = Right(dd, Len(dd) - finds)
               finds = InStr(dd, ".")
               mid_dd = Val(Left(dd, finds - 1))
               h_comm(i) = h_comm(i) + mid_dd
               dd = Right(dd, Len(dd) - finds)
               h_comm(i) = h_comm(i) + Val(dd) / 100
               If i < menum Then
                  mapinfo.do "fetch next from temp"
               End If
           Next
           h_comm(i) = -1
        End If
      If hg = False Then
        If hu = False Then
           mapinfo.do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + "ASSIGNMENT COMMAND" + Chr(34) + "into temp order by col1"
        Else
           mapinfo.do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + "ASSIGNMENT COMPLETE" + Chr(34) + "into temp order by col1"
        End If
      Else
         If hu = False Then
            mapinfo.do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + "DISCONNECT" + Chr(34) + "into temp order by col1"
         Else
           mapinfo.do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + "CONNECT" + Chr(34) + "into temp order by col1"
         End If
      End If
      DoEvents
        menum = mapinfo.eval("tableinfo(temp,8)")
        If menum > 0 Then
           mapinfo.do "fetch first from temp"
           For i = 1 To menum
               tt = mapinfo.eval("temp.col1")
               dd = tt
               If Len(dd) <= 8 Then GoTo nohour2
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 3600
               h_comp(i) = mid_dd
               dd = Right(dd, Len(dd) - finds)
nohour2:
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 60
               h_comp(i) = h_comp(i) + mid_dd
               dd = Right(dd, Len(dd) - finds)
               finds = InStr(dd, ".")
               mid_dd = Val(Left(dd, finds - 1))
               h_comp(i) = h_comp(i) + mid_dd
               dd = Right(dd, Len(dd) - finds)
               h_comp(i) = h_comp(i) + Val(dd) / 100
               If i < menum Then
                   mapinfo.do "fetch next from temp"
               End If
           Next
           h_comp(i) = -1
        End If
        
        If hu = False And hg = False Then GoTo efid
           assi = 1
           assj = 1
           Do While h_comm(assi) <> -1 And h_comm(assi) <> 0
              If h_comm(assi) <= h_comp(assj) Then
                 If h_comm(assi + 1) > h_comp(assj) Or h_comm(assi + 1) = -1 Or h_comm(assi + 1) = 0 Then
                    If hg = False And hu = True Then assnum = assnum + 1
                    If hg = True And hu = False Then assnum = assnum + 1
                    assi = assi + 1
                    assj = assj + 1
                 Else
                    If hg = True And hu = True Then assnum = assnum + 1
                    assi = assi + 1
                    If h_comm(assi) = -1 Or h_comm(assi) = 0 Then Exit Do
                 End If
              Else
                 assj = assj + 1
                 If h_comp(assj) = -1 Or h_comp(assj) = 0 Then Exit Do
              End If
           Loop
           Exit Sub
efid:
           i = 1
           j = 1
           tal = 0
           Do While h_comm(i) <> -1 And h_comm(i) <> 0
              hav = False
              If h_comm(i) <= h_comp(j) Then
                 If h_comm(i + 1) > h_comp(j) Or h_comm(i + 1) = -1 Or h_comm(i + 1) = 0 Then
                    ggg = h_comp(j) - h_comm(i)
                    j = j + 1
                    i = i + 1
                    hav = True
                 Else
                    i = i + 1
                 End If
              Else
                 j = j + 1
                 If h_comp(j) = -1 Or h_comp(j) = 0 Then Exit Do
              End If
              If hu = True Then Exit Sub
              If hav = True Then
                 Select Case ggg
                     Case Is >= 5
                          tal = 8
                     Case Is >= 2
                          tal = 7
                     Case Is >= 1
                          tal = 6
                     Case Is >= 0.5
                          tal = 5
                     Case Is >= 0.3
                          tal = 4
                     Case Is >= 0.2
                          tal = 3
                     Case Is >= 0.1
                          tal = 2
                     Case Is >= 0
                          tal = 1
                 End Select
                 
                 If table_s(tal, 1) = 0 Then
                    table_s(tal, 1) = 1
                    table_s(tal, 2) = ggg
                    table_s(tal, 3) = ggg
                    table_s(tal, 4) = ggg
                 Else
                    table_s(tal, 1) = table_s(tal, 1) + 1
                    If table_s(tal, 2) < ggg Then table_s(tal, 2) = ggg
                    If table_s(tal, 4) > ggg Then table_s(tal, 4) = ggg
                       table_s(tal, 3) = (table_s(tal, 3) + ggg) / 2
                 End If
              End If
           Loop
    
    Next hh
errend:
End Sub

Sub w_insert(ByVal bg As String, ByVal hj As String)
    word.selection.typetext Text:=bg
    word.selection.Font.Bold = -1
    word.selection.typetext Text:=hj
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
End Sub

Sub esceed_a(ByVal p As Integer)
 On Error GoTo errend
 Select Case p
 Case 3:
    word.selection.TypeParagraph
    word.selection.ParagraphFormat.TabStops.ClearAll
    word.ActiveDocument.DefaultTabStop = word.CentimetersToPoints(0.75)
    word.selection.Font.Size = 9
    word.selection.Font.Bold = -1
    If word.selection.Font.Underline = 0 Then
       word.selection.Font.Underline = 1
    End If
'    word.selection.typetext Text:= "天线优化建议"
    word.selection.typetext Text:="天线下行测试质量报告"
    word.selection.Font.Underline = 0
    word.selection.Font.Bold = 0
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(0.21), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(1.06), Alignment:=0, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(4.13), Alignment:=1, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(5.31), Alignment:=0, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(8.89), Alignment:=0, Leader:=0
    word.selection.ParagraphFormat.TabStops.Add Position:=word.CentimetersToPoints(12.06), Alignment:=0, Leader:=0
    word.selection.typetext Text:=Chr(9) & "序号" & Chr(9) & "天线名称" & Chr(9) & "误码率" & Chr(9) & "行动方案" & Chr(9) & "故障来源" & Chr(9) & "改善建议"
    word.selection.TypeParagraph
    word.selection.Font.Size = 9
 Case 4:
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.TypeParagraph
    word.selection.Font.Size = 9
    word.centerpara
    word.selection.typetext Text:="GSM 网络测试优化流程"
    word.selection.TypeParagraph
    
    word.leftpara
    word.selection.TypeParagraph
    word.FontSize 10
    word.selection.typetext Text:="测试手段    工程内容                   测试时间     完成情况"
    word.selection.TypeParagraph
    word.selection.typetext Text:="车载测量    确定现存状态，识别问题来源"
    word.selection.TypeParagraph
    word.selection.typetext Text:="　　　　    修改现存邻小区表"
    word.selection.TypeParagraph
    word.selection.typetext Text:="车载测量    基站和天线审视"
    word.selection.TypeParagraph
    word.selection.typetext Text:="　　　　    频率重安排"
    word.selection.TypeParagraph
    word.selection.typetext Text:="车载测量    细调邻小区切换"
    word.selection.TypeParagraph
    word.selection.typetext Text:="车载测量    完成系统参数记录"
    word.Linedown 1
    word.selection.typetext Text:=pagebreak
    word.selection.typetext Text:=Chr(13) + Chr(10)

 End Select
errend:
End Sub

Sub RxQual_Percent(My_Ci As String, My_Percent As String)
    Dim i As Integer
    Dim All_Count As Long
    Dim Check_Count As Integer
    Dim Check_Percent As Single
    On Error Resume Next
        
    All_Count = 0
    Check_Count = 0
    For i = 1 To stre_num
        mapinfo.do "select * from " + stre_tab(i) + " where col16 = " + Chr(34) + Trim(My_Ci) + Chr(34) + " into ci_temp"
        All_Count = All_Count + mapinfo.eval("tableinfo(ci_temp,8)")
        If mapinfo.eval("tableinfo(ci_temp,8)") > 0 Then
           If Report_Full = True Then
              mapinfo.do "select * from ci_temp where col23 > " + Format(Report_Qual) + " into temp"
           Else
              mapinfo.do "select * from ci_temp where col25 > " + Format(Report_Qual) + " into temp"
           End If
           Check_Count = Check_Count + mapinfo.eval("tableinfo(temp,8)")
        End If
    Next
    If All_Count = 0 Then
       My_Percent = ""
       Exit Sub
    End If
    Check_Percent = Check_Count / All_Count
    My_Percent = Format(Check_Percent, "Percent")
End Sub

Function FindSource(my_arfcn, my_bsic)
    Dim Mysource As String, myFind As String
    Dim i As Integer, finds As Integer
    Dim TempRow As Variant
    
    On Error Resume Next
    Mysource = ""
    mapinfo.do "select * from cell where arfcn = " & my_arfcn & " and bsic = " & my_bsic & " into temp"
    TempRow = mapinfo.eval("tableinfo(temp,8)")
    mapinfo.do "fetch first from temp"
    If TempRow > 0 Then
       For i = 1 To TempRow
           myFind = mapinfo.eval("temp.cell_name")
           finds = InStr(myFind, Chr(0))
           If finds > 0 Then
              myFind = Trim(Left(myFind, finds - 1))
           End If
           mapinfo.do "fetch next from temp"
           If i < TempRow Then
              Mysource = Mysource + myFind + " 或 "
           Else
              Mysource = Mysource + myFind
           End If
       Next
    End If
    FindSource = Mysource
End Function
