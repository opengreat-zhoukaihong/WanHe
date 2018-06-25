Attribute VB_Name = "report1"
Option Explicit

Public RangeNum As Integer, pp As Integer '
Public RxLevRange(1 To 2, 1 To 16) As String '
Public Data_Report As Boolean
Public stre_s(0 To 12) As Boolean
Public stre_tab(1 To 50) As String
Dim stre_tab_cell(1 To 50) As String
Public stre_num As Integer
Public Cell_Report As Boolean
Public all_max As Integer, all_min As Integer
Public all_avg As Single
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
Dim FileNumber As Integer '
Public RxlevFullFlag As Boolean '
Public TaFlag As Boolean
Public C1C2Flag As Boolean
Public IsQuickConvert As Boolean '
Dim MyExcel As Object
Public stcname As String
Public Gsm900Dcs1800Flag As Boolean
Dim CellRownum As Integer
Dim AssignmntFlag As Boolean


Sub TEST_REPORT()
    Dim putin, putin1, putin2, putin3, putin5
    Dim Cell_5, Cell_6, Hand_5, Hand_6, Alert_4, Alert_5, Loca_6, Loca_7, Loca_8, Conn_4, Conn_5
    Dim Block_2, Dropp_2, Gsm_n1, Dcs_n1
    Dim CellName() As String, dtx() As String
    Dim Gsm_n As Long, Dcs_n As Long, GsmDcs_n As Integer
    'Dim mapci
    'Dim CellCi As String * 5, oldci As String * 5
    'Dim cellno As Integer
    'Dim ci() As String * 5
    Dim setup_n As Integer, tmp1_n As Integer, tmp2_n As Integer, tmp3_n As Integer, tmp4_n As Integer
    Dim tmp_1 As Integer, tmp_2 As Integer, tmp_3 As Integer, tmp_4 As Integer, tmp_5 As Integer
    Dim Alert_1 As Integer, Alert_2 As Integer, Alert_3 As Integer, tmp5_n As Integer
    Dim Conn_1 As Integer, Conn_2 As Integer, Conn_3 As Integer
    Dim Loca_1 As Integer, Loca_2 As Integer, Loca_3 As Integer, Loca_4 As Integer, Loca_5 As Integer
    Dim Block_1 As Integer, Dropp_1 As Integer, putin4 As Integer
    Dim Hand_1 As Integer, Hand_2 As Integer, Hand_3 As Integer
    Dim Hand_4 As Integer
    Dim Cell_1 As Integer, Cell_2 As Integer, Cell_3 As Integer, Cell_4 As Integer
    Dim Nenum As Integer, tt As Integer, setup_n1 As Integer
    Dim myb As String, msg1 As String
    Dim com_hc As String, com_hs As String, com_hf As String
    Dim com_hmax As String, com_hmin As String, com_havg As String
    Dim com_qmax As String, com_qmin As String, com_qavg As String
    Dim com_xmax As String, com_xmin As String, com_xavg As String
    Dim pri_tbl As String
    Dim i As Integer
    Dim MyTempPath As String
    Dim all_0 As String
    Dim MyTableName As String, MyDbName As String
    Dim dbs As Database, rst As Recordset
    Dim lngRecords As Long, lngFields As Long
    Dim lngRecCount As Long, lngFldCount As Long
    Dim RowNum As Integer

    On Error Resume Next
    
    Screen.MousePointer = 11 '呼叫统计
    AssignmntFlag = False
    MyTempPath = Gsm_Path + "\user\"
    If dir(MyTempPath, 16) <> "" Then
       ChDir MyTempPath
    Else
       MkDir MyTempPath
    End If
    stcname = Gsm_Path + "\user\" + stre_tab(1) + ".xls"
    cc_all = 0
    pri_tbl = ""
    For i = 1 To stre_num
        MyTableName = convert_filename(i)
        MyDbName = ""
        Do While InStr(MyTableName, "\") > 0
            MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
            MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
        Loop
            MyDbName = Left(MyDbName, Len(MyDbName) - 1)
        pri_tbl = pri_tbl + MyTableName
        If i < stre_num Then pri_tbl = pri_tbl + ";"
        Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
        Set rst = dbs.OpenRecordset("SELECT  " _
        & " count(*) as countrxlev FROM " & MyTableName)
        If rst.RecordCount <> 0 Then
            rst.MoveLast
        End If
        lngRecords = rst.RecordCount
        lngFields = rst.Fields.Count
        cc_all = cc_all + rst.Fields(0).Value
    
    Next
    
    setup_n = mess_num("SETUP") '试呼次数
    setup_n1 = mess_num("EMERGENCY SETUP")
    setup_n = setup_n + setup_n1
    'tmp_1 = Mark_num("Call Successes") '呼叫成功次数
    Call xt_time(False, tmp_1, True)
    
    If setup_n > 0 Then
        putin = Format(tmp_1 / setup_n, "percent") '呼叫成功率
    End If
    
    'tmp_2 = Mark_num("Call Setup Succe") '呼叫建立成功次数
    Call xt_time(True, tmp_2, False)
    tmp_3 = Mark_num("Call Setup Fail") '呼叫建立失败次数
    If setup_n > 0 Then
        putin2 = Format(tmp_2 / setup_n, "percent") '呼叫建立成功率
        putin3 = Format(tmp_3 / setup_n, "percent") '呼叫建立失败率
    End If
    putin4 = setup_n - (tmp_2 + tmp_3) '求知
    If putin4 < 0 Then putin4 = 0
    Alert_1 = mess_num("ALERTING") '呼叫振铃成功次数
    If Alert_1 > setup_n Then Alert_1 = setup_n
    Alert_2 = Mark_num("Call Alert Fail") '呼叫振铃失败次数
    Alert_2 = Alert_2 + tmp_3
    If setup_n > 0 Then
        Alert_4 = Format(Alert_1 / setup_n, "percent") '呼叫振铃成功率
        Alert_5 = Format(Alert_2 / setup_n, "percent") '呼叫振铃失败率
    End If
    Alert_3 = setup_n - (Alert_1 + Alert_2) '求知
    If Alert_3 < 0 Then Alert_3 = 0
    Conn_1 = mess_num("CONNECT ACKNOWLEDGE") '呼叫连接成功次数
    If Conn_1 > setup_n Then Conn_1 = setup_n
    Conn_2 = Mark_num("Call Conn Failur") ''呼叫连接失败次数
    Conn_2 = Conn_2 + Alert_2
    If setup_n > 0 Then
        Conn_4 = Format(Conn_1 / setup_n, "percent") '呼叫连接成功率
        Conn_5 = Format(Conn_2 / setup_n, "percent") '呼叫连接失败率
    End If
    Conn_3 = setup_n - (Conn_1 + Conn_2) '求知
    If Conn_3 < 0 Then Conn_3 = 0
    tmp_4 = setup_n - Conn_1 '呼叫失败次数
    If setup_n > 0 Then
        putin5 = Format(tmp_4 / setup_n, "percent") '呼叫失败率
    End If
    tmp_5 = setup_n - (tmp_1 + tmp_4) '求知
    If tmp_5 < 0 Then tmp_5 = 0
    Loca_1 = mess_num("LOCATION UPDATING REQUEST") '位置更新次数
    Loca_2 = mess_num("LOCATION UPDATING ACCEPT") '位置更新成功次数
    If Loca_2 > Loca_1 Then Loca_2 = Loca_1
    Loca_3 = mess_num("LOCATION UPDATING REJECT") '位置更新失败次数
    Loca_4 = Mark_num("Loca Updata Tend") '位置更新终止次数
    If Loca_4 > Loca_1 Then Loca_4 = Loca_1
    Loca_5 = Loca_1 - (Loca_2 + Loca_3 + Loca_4) '未知
    If Loca_1 > 0 Then
        Loca_6 = Format(Loca_2 / Loca_1, "percent") '位置更新成功率
        Loca_7 = Format(Loca_3 / Loca_1, "percent") '位置更新失败率
        Loca_8 = Format(Loca_4 / Loca_1, "percent") '位置更新终止率
    End If
    
    Block_1 = Mark_num("Blocked Call") '拥塞次数
    If setup_n > 0 Then
        Block_2 = Format(Block_1 / setup_n, "percent") '拥塞率
    End If
    
    Dropp_1 = Mark_num("Dropped Call") '掉话次数
    If setup_n > 0 Then
        Dropp_2 = Format(Dropp_1 / setup_n, "percent") '掉话率
    End If
    
    Hand_1 = mess_num("HANDOVER COMMAND") '切换次数
    'Hand_2 = mess_num("HANDOVER COMPLETE") '切换成功次数
    Hand_2 = Mark_num("Handover Complet") '切换成功次数
    If Hand_2 > Hand_1 Then Hand_2 = Hand_1
    Hand_3 = mess_num("HANDOVER FAILURE") '切换失败次数
    If Hand_3 > Hand_1 Then Hand_3 = Hand_1
    Hand_4 = Hand_1 - (Hand_2 + Hand_3) '未知
    If Hand_4 < 0 Then Hand_4 = 0
    If Hand_1 > 0 Then
        Hand_5 = Format(Hand_2 / Hand_1, "percent") '成功切换率
        Hand_6 = Format(Hand_3 / Hand_1, "percent") '失败切换率
    End If
    
    Cell_1 = Mark_num("Intracell Atte") '小区内切换次数
    Cell_2 = Mark_num("Intracell Succe") '小区内成功切换次数
    If Cell_2 > Cell_1 Then Cell_2 = Cell_1
    Cell_3 = Mark_num("Intracell Fail") '小区内失败切换次数
    If Cell_3 > Cell_1 Then Cell_3 = Cell_1
    Cell_4 = Cell_1 - (Cell_2 + Cell_3) '未知
    If Cell_4 < 0 Then Cell_4 = 0
    If Cell_1 > 0 Then
        Cell_5 = Format(Cell_2 / Cell_1, "percent") '小区内成功切换率
        Cell_6 = Format(Cell_3 / Cell_1, "percent") '小区内失败切换率
    End If
    Screen.MousePointer = 0
    Set MyExcel = CreateObject("excel.application")
    MyExcel.Visible = True
    MyExcel.Workbooks.ADD
    MyExcel.Application.DisplayAlerts = False
    MyExcel.Sheets("Sheet3").Select
    MyExcel.ActiveWindow.SelectedSheets.Delete
    MyExcel.Sheets("Sheet1").Select
    MyExcel.Columns("A:A").ColumnWidth = 28.88 '19.75 '29.63
    MyExcel.Columns("B:B").ColumnWidth = 8.38 ' 9.38
    MyExcel.ActiveWindow.ScrollColumn = 4
    MyExcel.ActiveWindow.SmallScroll ToRight:=-2
    MyExcel.Columns("C:C").ColumnWidth = 19.25 '55.75
    MyExcel.ActiveWindow.ScrollColumn = 1
    MyExcel.Sheets("Sheet1").Name = "呼叫统计"
    MyExcel.Range("A1").Select
    'MyExcel.Rows("1:1").RowHeight = 27.75
    BoldfacedSize '粗体
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 16
    MyExcel.ActiveCell.FormulaR1C1 = "ANT 信令统计报告"
 
    MyExcel.Range("A2").Select
    MyExcel.Application.CutCopyMode = False
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 3
    FontSize '9号宋体
    MyExcel.ActiveCell.FormulaR1C1 = "...通话信令统计..."
        
    MyExcel.Range("A3").Select
    FontSize '9号宋体
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "1.通话信令统计(Call ATTEMPT)"
    
    
    MyExcel.Range("B3").Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    HorVerSize '中间
    MyExcel.Selection.Merge
    FontSize '9号宋体
    MyExcel.ActiveCell.FormulaR1C1 = "统计数值"
    
    MyExcel.Range("C3").Select
    HorVerSize '中间
    MyExcel.Selection.Merge
    'MyExcel.Selection.Interior.ColorIndex = 15
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    FontSize '9号宋体
    MyExcel.ActiveCell.FormulaR1C1 = "信令"
    
    MyExcel.Range("A4").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("A14")
    MyExcel.Range("A4:A14").Select
    FontSize '9号宋体
    MyExcel.Range("A4").Select
    MyExcel.ActiveCell.FormulaR1C1 = "通话次数(Call Attempts):"
    MyExcel.Range("A5").Select
    MyExcel.ActiveCell.FormulaR1C1 = "完整通话信令次数(Call Successes):"
    MyExcel.Range("A6").Select
    MyExcel.ActiveCell.FormulaR1C1 = "不完整通话信令次数(Call Failures):"
    MyExcel.Range("A12").Select
    MyExcel.ActiveCell.FormulaR1C1 = "未知通话信令次数(Unknown Call endings):"
    MyExcel.Range("A13").Select
    MyExcel.ActiveCell.FormulaR1C1 = "完整信令率:"
    MyExcel.Range("A14").Select
    MyExcel.ActiveCell.FormulaR1C1 = "不完整信令率:"
    
    MyExcel.Range("C4").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("C13")
    MyExcel.Range("C4:C13").Select
    FontSize 'Arial体
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.Range("C4").Select
    MyExcel.ActiveCell.FormulaR1C1 = """CHANNEL REQUEST""&""SETUP"""
        
    
    MyExcel.Range("C5").Select
    MyExcel.ActiveCell.FormulaR1C1 = """CONNECT ACKNOWLEDGE""&""DISCONNECT""&""成对"""
    
    MyExcel.Range("C6").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """SET UP""&""CHANNEL RELEASE"" OR""CHANNEL REQUEST"""
    
    MyExcel.Range("C7").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """ASSIGNMENT COMMAND""&""CHANNEL RELEASE""OR""CHANNEL REQUEST"""
    MyExcel.Range("C8").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """ALERTING""&""CHANNEL RELEASE"" OR""CHANNEL REQUEST"""
    MyExcel.Range("C9").Select
    MyExcel.ActiveCell.FormulaR1C1 = """CONNECT""&""CHANNEL RELEASE""OR""CHANNEL REQUEST"""
    MyExcel.Range("C10").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """CONNECT ACKNOWLEDGE""&""CHANNEL RELEASE"" OR""CHANNEL REQUEST"""
    MyExcel.Range("C11").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """ASSIGNMENT COMMAND""&""CHANNEL RELEASE""OR""CHANNEL REQUEST"""
    MyExcel.Range("C12").Select
    MyExcel.ActiveCell.FormulaR1C1 = """CHANNEL REQUEST""&""SETUP""&""后，信令不完整"""
    
    '*******************************************************
    MyExcel.Range("B4").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("B15")
    MyExcel.Range("B4:B15").Select
    FontSize 'Arial体
    HorVerSize '中间
    ' MyExcel.Selection.Font.ColorIndex = 7
    MyExcel.Range("B4").Select
    MyExcel.ActiveCell.FormulaR1C1 = setup_n
    MyExcel.Range("B5").Select
    MyExcel.ActiveCell.FormulaR1C1 = tmp_1
    MyExcel.Range("B6").Select
    MyExcel.ActiveCell.FormulaR1C1 = tmp_4
    MyExcel.Range("B12").Select
    MyExcel.ActiveCell.FormulaR1C1 = tmp_5
    MyExcel.Range("B13").Select
    MyExcel.ActiveCell.FormulaR1C1 = putin
    MyExcel.Range("B14").Select
    MyExcel.ActiveCell.FormulaR1C1 = putin5
    MyExcel.Rows("13:14").Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    '*******************************************************
    MyExcel.Range("A16").Select
    
    FontStyle
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "2.建立信令(Call SETUP)"
    
    MyExcel.Range("A17").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("A22")
    MyExcel.Range("A17:A23").Select
    FontSize
    MyExcel.Range("A17").Select
    MyExcel.ActiveCell.FormulaR1C1 = "建立信令成功次数(Call Setup Successes):"
    MyExcel.Range("A20").Select
    MyExcel.ActiveCell.FormulaR1C1 = "建立信令失败次数(Call Setup Failures):"
    MyExcel.Range("A21").Select
    MyExcel.ActiveCell.FormulaR1C1 = "未知建立信令次数(Unknown Call Setup):"
    MyExcel.Range("A22").Select
    MyExcel.ActiveCell.FormulaR1C1 = "建立信令成功率:"
    MyExcel.Range("A23").Select
    MyExcel.ActiveCell.FormulaR1C1 = "建立信令失败率:"
    
    MyExcel.Range("C17").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("C21")
    MyExcel.Range("C17:C21").Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.Range("C17").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """SET UP""&""ASSGNMENT COMMAND"" "
    MyExcel.Range("C18").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """ALERTING""&""ASSGNMENT COMMAND"" "
    MyExcel.Range("C19").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """CONNECT""&""ASSGNMENT COMMAND"" "
    MyExcel.Range("C20").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """SET UP""&""CHANNEL RELEASE"" OR ""CHANNEL REQUEST"" "
    MyExcel.Range("C21").Select
    MyExcel.ActiveCell.FormulaR1C1 = """SETUP后信令不完整"""

    '********************************************************
    MyExcel.Range("B17").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("B24")
    MyExcel.Range("B17:B24").Select
    FontSize
    HorVerSize
   ' MyExcel.Selection.Font.ColorIndex = 7
    MyExcel.Range("B17").Select
    MyExcel.ActiveCell.FormulaR1C1 = tmp_2
    MyExcel.Range("B20").Select
    MyExcel.ActiveCell.FormulaR1C1 = tmp_3
    MyExcel.Range("B21").Select
    MyExcel.ActiveCell.FormulaR1C1 = putin4
    MyExcel.Range("B22").Select
    MyExcel.ActiveCell.FormulaR1C1 = putin2
    MyExcel.Range("B23").Select
    MyExcel.ActiveCell.FormulaR1C1 = putin3
    MyExcel.Rows("22:23").Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    '*******************************************************
    
    MyExcel.Range("A25").Select
    FontStyle
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "3.振铃信令(CALL ALERTING)"
    
    MyExcel.Range("A26").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("A32")
    MyExcel.Range("A26:A33").Select
    FontSize
    MyExcel.Range("A26").Select
    MyExcel.ActiveCell.FormulaR1C1 = "振铃信令成功次数(Call Alert Successes):"
    MyExcel.Range("A30").Select
    MyExcel.ActiveCell.FormulaR1C1 = "振铃信令失败次数(Call Alert Failures):"
    MyExcel.Range("A31").Select
    MyExcel.ActiveCell.FormulaR1C1 = "未知振铃信令次数(Unknown Connect):"
    MyExcel.Range("A32").Select
    MyExcel.ActiveCell.FormulaR1C1 = "振铃信令成功率:"
    MyExcel.Range("A33").Select
    MyExcel.ActiveCell.FormulaR1C1 = "振铃信令失败率:"
    
    MyExcel.Range("C26").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("C31")
    MyExcel.Range("C26:C31").Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.Range("C26").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """SET UP""&""ALERTING"" "
    MyExcel.Range("C27").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """CHANNEL REQUEST""&""ALERTING"" "
    MyExcel.Range("C28").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """ASSGNMENT COMMAND""&""ALERTING"" "
    MyExcel.Range("C29").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """CONNECT""&""ALERTING""  "
    MyExcel.Range("C30").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """ALERTING""&""CHANNEL RELEASE"" OR ""CHANNEL REQUEST"" "
    MyExcel.Range("C31").Select
    MyExcel.ActiveCell.FormulaR1C1 = """ALERTING后信令不完整"""

    '********************************************************
    MyExcel.Range("B26").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("B34")
    MyExcel.Range("B26:B34").Select
    FontSize
    HorVerSize
    MyExcel.Range("B26").Select
    MyExcel.ActiveCell.FormulaR1C1 = Alert_1
    MyExcel.Range("B30").Select
    MyExcel.ActiveCell.FormulaR1C1 = Alert_2
    MyExcel.Range("B31").Select
    MyExcel.ActiveCell.FormulaR1C1 = Alert_3
    MyExcel.Range("B32").Select
    MyExcel.ActiveCell.FormulaR1C1 = Alert_4
    MyExcel.Range("B33").Select
    MyExcel.ActiveCell.FormulaR1C1 = Alert_5
    MyExcel.Rows("32:33").Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    '*******************************************************
    MyExcel.Range("A35").Select
    FontStyle
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "4.连接信令(CALL CONNECT)"
    
    MyExcel.Range("A36").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("A42")
    MyExcel.Range("A36:A42").Select
    FontSize
    MyExcel.Range("A36").Select
    MyExcel.ActiveCell.FormulaR1C1 = "连接信令成功次数(Call Connect Successes):"
    MyExcel.Range("A39").Select
    MyExcel.ActiveCell.FormulaR1C1 = "连接信令失败次数(Call Connect Failures):"
    MyExcel.Range("A40").Select
    MyExcel.ActiveCell.FormulaR1C1 = "未知连接信令次数(Unknown Connect):"
    MyExcel.Range("A41").Select
    MyExcel.ActiveCell.FormulaR1C1 = "连接信令成功率:"
    MyExcel.Range("A42").Select
    MyExcel.ActiveCell.FormulaR1C1 = "连接信令失败率:"
    
    MyExcel.Range("C36").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("C41")
    MyExcel.Range("C36:C42").Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.Range("C36").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """ASSGNMENT COMMAND""&""CONNECT ACKNOWLEDGE"" "
    MyExcel.Range("C37").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """CONNECT""&""CONNECT ACKNOWLEDGE"" "
    MyExcel.Range("C38").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """CHANNEL REQUEST""&""CONNECT ACKNOWLEDGE"" "
    MyExcel.Range("C39").Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """CONNECT""&""CHANNEL RELEASE"" OR ""CHANNEL REQUEST"" "
    MyExcel.Range("C40").Select
    MyExcel.ActiveCell.FormulaR1C1 = """CONNECT后信令不完整"""

    '********************************************************
    MyExcel.Range("B36").Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("B42")
    MyExcel.Range("B36:B42").Select
    FontSize
    HorVerSize
    MyExcel.Range("B36").Select
    MyExcel.ActiveCell.FormulaR1C1 = Conn_1
    MyExcel.Range("B39").Select
    MyExcel.ActiveCell.FormulaR1C1 = Conn_2
    MyExcel.Range("B40").Select
    MyExcel.ActiveCell.FormulaR1C1 = Conn_3
    MyExcel.Range("B41").Select
    MyExcel.ActiveCell.FormulaR1C1 = Conn_4
    MyExcel.Range("B42").Select
    MyExcel.ActiveCell.FormulaR1C1 = Conn_5
    MyExcel.Rows("41:42").Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    '**********************
    RowNum = 44
    MyExcel.Range("A" & RowNum).Select
    FontSize
    MyExcel.ActiveCell.FormulaR1C1 = "通话建立拥塞次数:"
    MyExcel.Range("A" & RowNum + 1).Select
    FontSize
    MyExcel.ActiveCell.FormulaR1C1 = "通话建立拥塞率:"
    MyExcel.Range("A" & RowNum + 2).Select
    FontSize
    MyExcel.ActiveCell.FormulaR1C1 = "掉话次数:"
    MyExcel.Range("A" & RowNum + 3).Select
    FontSize
    MyExcel.ActiveCell.FormulaR1C1 = "掉话率:"
    MyExcel.Range("C" & RowNum).Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("C" & RowNum + 4)
    MyExcel.Range("C" & RowNum & ":" & "C" & RowNum + 4).Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.Range("C" & RowNum).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """CONNECT""NO""CONNECT ACKNOWLEDGE"""

    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """CONNECT ACKNOWLEDGE""&""CHANNEL RELEASE"" OR ""CHANNEL REQUEST"""
    MyExcel.Range("B" & RowNum).Select
    FontSize
    HorVerSize
    MyExcel.ActiveCell.FormulaR1C1 = Block_1
    MyExcel.Range("B" & RowNum + 1).Select
    FontSize
    HorVerSize
    MyExcel.ActiveCell.FormulaR1C1 = Block_2
    MyExcel.Range("B" & RowNum + 2).Select
    FontSize
    HorVerSize
    MyExcel.ActiveCell.FormulaR1C1 = Dropp_1
    MyExcel.Range("B" & RowNum + 3).Select
    FontSize
    HorVerSize
    MyExcel.ActiveCell.FormulaR1C1 = Dropp_2
    MyExcel.Rows(RowNum & ":" & RowNum + 3).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    
    RowNum = RowNum + 5
    
    '*******************************************************
    MyExcel.Range("A" & RowNum).Select
    FontStyle
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "5.位置更新信令(LOCATION UPDATING)"
    
    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("A" & RowNum + 8)
    MyExcel.Range("A" & RowNum + 1 & ":" & "A" & RowNum + 8).Select
    FontSize
    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.ActiveCell.FormulaR1C1 = "位置更新次数(LOCATION UPDATING Attempts):"
    MyExcel.Range("A" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "位置更新成功次数(LOCATION UPDATING Successes):"
    MyExcel.Range("A" & RowNum + 3).Select
    MyExcel.ActiveCell.FormulaR1C1 = "位置更新失败次数(LOCATION UPDATING Failures):"
    MyExcel.Range("A" & RowNum + 4).Select
    MyExcel.ActiveCell.FormulaR1C1 = "位置更新终止次数(LOCATION UPDATING Ternimated):"
    MyExcel.Range("A" & RowNum + 5).Select
    MyExcel.ActiveCell.FormulaR1C1 = "未知位置更新(Unknown LOCATION UPDATING):"
    MyExcel.Range("A" & RowNum + 6).Select
    MyExcel.ActiveCell.FormulaR1C1 = "位置更新成功率:"
    MyExcel.Range("A" & RowNum + 7).Select
    MyExcel.ActiveCell.FormulaR1C1 = "位置更新失败率:"
    MyExcel.Range("A" & RowNum + 8).Select
    MyExcel.ActiveCell.FormulaR1C1 = "位置更新终止率:"
    
    MyExcel.Range("C" & RowNum + 1).Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("C" & RowNum + 8)
    MyExcel.Range("C" & RowNum + 1 & ":" & "C" & RowNum + 8).Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.Range("C" & RowNum + 1).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """LOCATION UPDATING REQUEST"""
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """LOCATION UPDATING REQUEST""&""LOCATION UPDATING ACCEPT"" "
    MyExcel.Range("C" & RowNum + 3).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """LOCATION UPDATING REQUEST""&""LOCATION UPDATING REJECT"" "
    MyExcel.Range("C" & RowNum + 4).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """LOCATION UPDATING REQUEST""&""CHANNEL REQUEST"" "
    MyExcel.Range("C" & RowNum + 5).Select
    MyExcel.ActiveCell.FormulaR1C1 = """LOCATION UPDATING REQUEST后不完整"""

    
    '********************************************************
    MyExcel.Range("B" & RowNum + 1).Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("B" & RowNum + 8)
    MyExcel.Range("B" & RowNum + 1 & ":" & "B" & RowNum + 8).Select
    FontSize
    HorVerSize
    MyExcel.Range("B" & RowNum + 1).Select
    MyExcel.ActiveCell.FormulaR1C1 = Loca_1
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = Loca_2
    MyExcel.Range("B" & RowNum + 3).Select
    MyExcel.ActiveCell.FormulaR1C1 = Loca_3
    MyExcel.Range("B" & RowNum + 4).Select
    MyExcel.ActiveCell.FormulaR1C1 = Loca_4
    MyExcel.Range("B" & RowNum + 5).Select
    MyExcel.ActiveCell.FormulaR1C1 = Loca_5
    MyExcel.Range("B" & RowNum + 6).Select
    MyExcel.ActiveCell.FormulaR1C1 = Loca_6
    MyExcel.Range("B" & RowNum + 7).Select
    MyExcel.ActiveCell.FormulaR1C1 = Loca_7
    MyExcel.Range("B" & RowNum + 8).Select
    MyExcel.ActiveCell.FormulaR1C1 = Loca_8
    MyExcel.Rows(RowNum + 6 & ":" & RowNum + 8).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    RowNum = RowNum + 10
    
    '*******************************************************
    MyExcel.Range("A" & RowNum).Select
    MyExcel.Application.CutCopyMode = False
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 3
    FontSize
    MyExcel.ActiveCell.FormulaR1C1 = "...切换信令统计..."
        
    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    FontSize
    MyExcel.ActiveCell.FormulaR1C1 = "1.切换信令(HANDOVER)"
    
    
    MyExcel.Range("B" & RowNum + 1).Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    HorVerSize
    MyExcel.Selection.Merge
    FontSize
    
    MyExcel.ActiveCell.FormulaR1C1 = "统计数值"
    
    MyExcel.Range("C" & RowNum + 1).Select
    HorVerSize
    MyExcel.Selection.Merge
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    FontSize
    
    MyExcel.ActiveCell.FormulaR1C1 = "信令"
    
    MyExcel.Range("A" & RowNum + 2).Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("A" & RowNum + 14)
    MyExcel.Range("A" & RowNum + 2 & ":" & "A" & RowNum + 14).Select
    FontSize
    MyExcel.Range("A" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "切换信令次数(Handover Attempts):"
    MyExcel.Range("A" & RowNum + 3).Select
    MyExcel.ActiveCell.FormulaR1C1 = "切换信令成功次数(Handover Successes):"
    MyExcel.Range("A" & RowNum + 4).Select
    MyExcel.ActiveCell.FormulaR1C1 = "切换信令失败次数(Handover Failures):"
    MyExcel.Range("A" & RowNum + 5).Select
    MyExcel.ActiveCell.FormulaR1C1 = "未知切换信令次数(Unknown Handover):"
    MyExcel.Range("A" & RowNum + 6).Select
    MyExcel.ActiveCell.FormulaR1C1 = "切换信令成功率:"
    MyExcel.Range("A" & RowNum + 7).Select
    MyExcel.ActiveCell.FormulaR1C1 = "切换信令失败率:"
    MyExcel.Range("A" & RowNum + 9).Select
    MyExcel.ActiveCell.FormulaR1C1 = "小区切换信令次数(Intra Cell Handover Attempts):"
    MyExcel.Range("A" & RowNum + 10).Select
    MyExcel.ActiveCell.FormulaR1C1 = "小区内切换信令成功次数(Intra Cell Handover Succeses):"
    MyExcel.Range("A" & RowNum + 11).Select
    MyExcel.ActiveCell.FormulaR1C1 = "小区内切换信令失败次数(Intra Cell Handover Failures):"
    MyExcel.Range("A" & RowNum + 12).Select
    MyExcel.ActiveCell.FormulaR1C1 = "未知小区内切换次数(Unknown Intracell Handover):"
    MyExcel.Range("A" & RowNum + 13).Select
    MyExcel.ActiveCell.FormulaR1C1 = "小区内切换成功率():"
    MyExcel.Range("A" & RowNum + 14).Select
    MyExcel.ActiveCell.FormulaR1C1 = "小区内切换失败率():"
    
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("C" & RowNum + 14)
    MyExcel.Range("C" & RowNum + 2 & ":" & "C" & RowNum + 14).Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """HANDOVER COMMAND"""
    MyExcel.Range("C" & RowNum + 3).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """HANDOVER COMPLETE""NO""HANDOVER FAILUER"" IN 1,5S "
    MyExcel.Range("C" & RowNum + 4).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """HANDOVER COMMAND""&""HANDOVER FAILUER"" "
    MyExcel.Range("C" & RowNum + 5).Select
    MyExcel.ActiveCell.FormulaR1C1 = """HANDOVER COMMAND后信令不完整"""
    MyExcel.Range("C" & RowNum + 9).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """ASSIGNMENT COMMAND"""
    MyExcel.Range("C" & RowNum + 10).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """ASSIGNMENT COMMAND""&""ASSIGNMENT COMPLETE"" "

    MyExcel.Range("C" & RowNum + 11).Select
    MyExcel.ActiveCell.FormulaR1C1 = _
        """ASSIGNMENT COMMAND""&""ASSIGNMENT FAILUER"""
    MyExcel.Range("C" & RowNum + 12).Select
    MyExcel.ActiveCell.FormulaR1C1 = """ASSIGNMENT COMMAND后信令不完整"""

    '********************************************************
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.Selection.Cut Destination:=MyExcel.Range("B" & RowNum + 14)
    MyExcel.Range("B" & RowNum + 2 & ":" & "B" & RowNum + 14).Select
    FontSize
    HorVerSize
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = Hand_1
    MyExcel.Range("B" & RowNum + 3).Select
    MyExcel.ActiveCell.FormulaR1C1 = Hand_2
    MyExcel.Range("B" & RowNum + 4).Select
    MyExcel.ActiveCell.FormulaR1C1 = Hand_3
    MyExcel.Range("B" & RowNum + 5).Select
    MyExcel.ActiveCell.FormulaR1C1 = Hand_4
    MyExcel.Range("B" & RowNum + 6).Select
    MyExcel.ActiveCell.FormulaR1C1 = Hand_5
    MyExcel.Range("B" & RowNum + 7).Select
    MyExcel.ActiveCell.FormulaR1C1 = Hand_6
    MyExcel.Rows(RowNum + 6 & ":" & RowNum + 7).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    MyExcel.Range("B" & RowNum + 9).Select
    MyExcel.ActiveCell.FormulaR1C1 = Cell_1
    MyExcel.Range("B" & RowNum + 10).Select
    MyExcel.ActiveCell.FormulaR1C1 = Cell_2
    MyExcel.Range("B" & RowNum + 11).Select
    MyExcel.ActiveCell.FormulaR1C1 = Cell_3
    MyExcel.Range("B" & RowNum + 12).Select
    MyExcel.ActiveCell.FormulaR1C1 = Cell_4
    MyExcel.Range("B" & RowNum + 13).Select
    MyExcel.ActiveCell.FormulaR1C1 = Cell_5
    MyExcel.Range("B" & RowNum + 14).Select
    MyExcel.ActiveCell.FormulaR1C1 = Cell_6
    MyExcel.Rows(RowNum + 13 & ":" & RowNum + 14).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    RowNum = RowNum + 16
    
    MyExcel.Range("A" & RowNum).Select
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 3
    MyExcel.ActiveCell.FormulaR1C1 = "...呼叫总表..."
    MyExcel.Rows(RowNum + 1 & ":" & RowNum + 1).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.ActiveCell.FormulaR1C1 = "文件名"
    MyExcel.Range("B" & RowNum + 1).Select
    MyExcel.ActiveCell.FormulaR1C1 = "启动时间"
    MyExcel.Range("C" & RowNum + 1).Select
    MyExcel.ActiveCell.FormulaR1C1 = "结束时间"
    MyExcel.Range("D" & RowNum + 1).Select
    MyExcel.ActiveCell.FormulaR1C1 = "时长(秒)"
    MyExcel.Range("E" & RowNum + 1).Select
    MyExcel.ActiveCell.FormulaR1C1 = "状态"
    MyExcel.Range("F" & RowNum + 1).Select
    Call Call_Attemp(RowNum + 2)
    
    
    
    '****************************************'ANT测量统计报告

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
   
    '******************
    MyExcel.Sheets("Sheet2").Select
    MyExcel.ActiveWindow.ScrollColumn = 4
    MyExcel.ActiveWindow.SmallScroll ToRight:=-2
    MyExcel.ActiveWindow.ScrollColumn = 1
    MyExcel.Sheets("Sheet2").Name = "测量统计分表"
    MyExcel.Range("A1").Select
    BoldfacedSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 16
    MyExcel.ActiveCell.FormulaR1C1 = "ANT 测量统计报告"
    
    MyExcel.Range("A2").Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...系统响应时间统计表..."
    
    MyExcel.Range("A3").Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.ActiveCell.FormulaR1C1 = "（信令过程：CHANNEL REQUEST 与ASSIGNMENT COMMAND）"
    If table_s(9, 1) = 0 Then
       MyExcel.Range("A4").Select
       MyExcel.Selection.Font.ColorIndex = 3
       FontSize
       MyExcel.ActiveCell.FormulaR1C1 = "无 CHANNEL REQUEST"
       RowNum = 6
       GoTo ewi
    End If
    MyExcel.Columns("A:A").Select
    MyExcel.Selection.ColumnWidth = 24.5 '21
    MyExcel.Columns("B:B").Select
    MyExcel.Selection.ColumnWidth = 5.63 '8.8
    MyExcel.Columns("C:C").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("D:D").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("E:E").ColumnWidth = 5.63 ' 8.88
    MyExcel.Columns("F:F").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("G:G").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("H:H").ColumnWidth = 5.63 '8.88
    MyExcel.Range("A4").Select
    MyExcel.ActiveCell.FormulaR1C1 = ""
    
    MyExcel.Rows("4:4").Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B4").Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    'MyExcel.Range("C4").Select
    'MyExcel.ActiveCell.FormulaR1C1 = "累结数"
    MyExcel.Range("C4").Select
    MyExcel.ActiveCell.FormulaR1C1 = "最大数"
    MyExcel.Range("D4").Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    MyExcel.Range("E4").Select
    MyExcel.ActiveCell.FormulaR1C1 = "最小数"
    MyExcel.Range("F4").Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("G4").Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    Call Row_Col("0s<=x<1s", 1, 5, "B")
    Call Row_Col("0.1s<=x<0.2s", 2, 6, "B")
    Call Row_Col("0.2<=x<0.3s", 3, 7, "B")
    Call Row_Col("0.3s<=x<0.5s", 4, 8, "B")
    Call Row_Col("0.5s<=x<1s", 5, 9, "B")
    Call Row_Col("1s<=x<2s", 6, 10, "B")
    Call Row_Col("2s<=x<5s", 7, 11, "B")
    Call Row_Col("5s<=x<15s", 8, 12, "B")
    Call Row_Col("总计", 9, 13, "B")
    com_xmax = table_f(9, 2)
    com_xavg = table_f(9, 3)
    com_xmin = table_f(9, 4)
    RowNum = 15
ewi:
     '***************呼叫切换统计表
    Call xtt_time(False, tt, False)
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
    MyExcel.Range("A" & RowNum).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...每呼叫切换频度统计表..."
    MyExcel.Range("A" & RowNum + 1).Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.ActiveCell.FormulaR1C1 = "（信令过程：ASSIGNMENT COMMAND与DISCONNECT"
    If table_s(9, 1) = 0 Then
        MyExcel.Range("A" & RowNum + 2).Select
        MyExcel.Selection.Font.ColorIndex = 3
        FontSize
        MyExcel.ActiveCell.FormulaR1C1 = "无ASSIGNMENT COMMAND"
        RowNum = RowNum + 4
        GoTo ei
    End If
    MyExcel.Range("A" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = ""
    
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    'MyExcel.Range("C" & RowNum + 2).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "最大数"
    'MyExcel.Range("D" & RowNum + 2).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "均值"
    'MyExcel.Range("E" & RowNum + 2).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "最小数"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    AssignmntFlag = True
    Call Row_Col("0m<=x<1m", 1, RowNum + 3, "B")
    Call Row_Col("1m<=x<2m", 2, RowNum + 4, "B")
    Call Row_Col("2m<=x<3m", 3, RowNum + 5, "B")
    Call Row_Col("3m<=x<5m", 4, RowNum + 6, "B")
    Call Row_Col("5m<=x<6m", 5, RowNum + 7, "B")
    Call Row_Col("6m<=x<7m", 6, RowNum + 8, "B")
    Call Row_Col("7m<=x<8m", 7, RowNum + 9, "B")
    Call Row_Col("x>=8m", 8, RowNum + 10, "B")
    Call Row_Col("总计", 9, RowNum + 11, "B")
    AssignmntFlag = False
    com_hmax = table_f(9, 2)
    com_havg = table_f(9, 3)
    com_hmin = table_f(9, 4)
    RowNum = RowNum + 13
ei:
     '***************
    
    '*************************切换性能评估统计表
    Dim enum1, enum2, enum3
    Call hand_time(enum1, enum2, enum3)
    MyExcel.Range("A" & RowNum).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...切换性能评估统计表..."
    
    MyExcel.Range("A" & RowNum + 1).Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.ActiveCell.FormulaR1C1 = "（信令过程：HANDOVER COMMAND与HANDOVER COMPLETE或HANDOVER COMMAND FAILUER之间）"
    If enum1 <= 0 Then
        MyExcel.Range("A" & RowNum + 2).Select
        MyExcel.Selection.Font.ColorIndex = 3
        FontSize
        MyExcel.ActiveCell.FormulaR1C1 = "无HANDOVER COMMAND"
        RowNum = RowNum + 4
        GoTo no_time
    End If
    MyExcel.Columns("A:A").Select
    MyExcel.Selection.ColumnWidth = 24.4
    MyExcel.Columns("B:B").Select
    MyExcel.Selection.ColumnWidth = 5.63 '8.8
    MyExcel.Columns("C:C").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("D:D").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("E:E").ColumnWidth = 5.63 ' 8.88
    MyExcel.Columns("F:F").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("G:G").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("H:H").ColumnWidth = 5.63 '8.88
    MyExcel.Range("A" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = ""
    
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最大数"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最小数"
    MyExcel.Range("F" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("G" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    hand_zz
    Call Row_Col("0s<=x<0.1s", 1, RowNum + 3, "B")
    Call Row_Col("0.1s<=x<0.2s", 2, RowNum + 4, "B")
    Call Row_Col("0.2<=x<0.3s", 3, RowNum + 5, "B")
    Call Row_Col("0.3s<=x<0.5s", 4, RowNum + 6, "B")
    Call Row_Col("0.5s<=x<1s", 5, RowNum + 7, "B")
    Call Row_Col("1s<=x<2s", 6, RowNum + 8, "B")
    Call Row_Col("2s<=x<5s", 7, RowNum + 9, "B")
    Call Row_Col("5s<=x<15s", 8, RowNum + 10, "B")
    Call Row_Col("总计", 9, RowNum + 11, "B")
    com_hmax = table_f(9, 2)
    com_havg = table_f(9, 3)
    com_hmin = table_f(9, 4)
    RowNum = RowNum + 13

    '***********************'切换间隔时间统计表
    Dim j
    
    For i = 1 To 9
        For j = 1 To 6
            table_s(i, j) = tibeh(i, j)
        Next
    Next
    hand_zz
    com_qmax = table_f(9, 2)
    com_qavg = table_f(9, 3)
    com_qmin = table_f(9, 4)
    Dim q1, q2, q3
       
    MyExcel.Range("A" & RowNum).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...切换间隔时间统计表..."
    
    MyExcel.Range("A" & RowNum + 1).Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.ActiveCell.FormulaR1C1 = "（信令过程：HANDOVER COMMAND与下一个HANDOVER COMMAND 之间）"
    If enum1 < 2 Then
        MyExcel.Range("A" & RowNum + 2).Select
        MyExcel.Selection.Font.ColorIndex = 3
        FontSize
        MyExcel.ActiveCell.FormulaR1C1 = "只有一个HANDOVER COMMAND"
        RowNum = RowNum + 4
        GoTo no_time
    End If
    MyExcel.Columns("A:A").Select
    MyExcel.Selection.ColumnWidth = 24.5
    MyExcel.Columns("B:B").Select
    MyExcel.Selection.ColumnWidth = 5.63 ' 8.8
    MyExcel.Columns("C:C").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("D:D").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("E:E").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("F:F").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("G:G").ColumnWidth = 5.63 ' 8.88
    MyExcel.Columns("H:H").ColumnWidth = 5.63 '8.88
    MyExcel.Range("A" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = ""
    
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最大数"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最小数"
    MyExcel.Range("F" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("G" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
   ' hand_zz
    Call Row_Col("0s<=x<1s", 1, RowNum + 3, "B")
    Call Row_Col("1s<=x<2s", 2, RowNum + 4, "B")
    Call Row_Col("2<=x<4s", 3, RowNum + 5, "B")
    Call Row_Col("4s<=x<10s", 4, RowNum + 6, "B")
    Call Row_Col("10s<=x<120s", 5, RowNum + 7, "B")
    Call Row_Col("2min<=x<20min", 6, RowNum + 8, "B")
    Call Row_Col("总计", 9, RowNum + 9, "B")
    com_hmax = table_f(9, 2)
    com_havg = table_f(9, 3)
    com_hmin = table_f(9, 4)
    RowNum = RowNum + 11
no_time:
     '*****************'双频测试统计表
    MyExcel.Range("A" & RowNum).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...双频测试统计表..."
    MyExcel.Rows(RowNum + 1 & ":" & RowNum + 1).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 1).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    MyExcel.Range("C" & RowNum + 1).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    Gsm_n = Gsm_Dcs("0", "125")
    Dcs_n = Gsm_Dcs("512", "886")
    GsmDcs_n = Gsm_n + Dcs_n
    Gsm_n1 = Format(Gsm_n / GsmDcs_n, "percent")
    Dcs_n1 = Format(Dcs_n / GsmDcs_n, "percent")
    
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontSize
    MyExcel.Range("A" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "GSM900"
    MyExcel.Range("B" & RowNum + 2 & ":" & "C" & RowNum + 2).Select
    LeftSize
    FontSize
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = Gsm_n
    MyExcel.Range("C" & RowNum + 2).Select
    FontSize
    MyExcel.ActiveCell.FormulaR1C1 = Gsm_n1
    '***********
    MyExcel.Rows(RowNum + 3 & ":" & RowNum + 3).Select
    FontSize
    MyExcel.Range("A" & RowNum + 3).Select
    MyExcel.ActiveCell.FormulaR1C1 = "DCS1800"
    MyExcel.Range("B" & RowNum + 3 & ":" & "C" & RowNum + 3).Select
    LeftSize
    FontSize
    MyExcel.Range("B" & RowNum + 3).Select
    MyExcel.ActiveCell.FormulaR1C1 = Dcs_n
    MyExcel.Range("C" & RowNum + 3).Select
    MyExcel.ActiveCell.FormulaR1C1 = Dcs_n1
    MyExcel.Rows(RowNum + 3 & ":" & RowNum + 3).Select
    MyExcel.Range("A" & RowNum + 4).Select
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.ActiveCell.FormulaR1C1 = "总计"
    MyExcel.Range("B" & RowNum + 4 & ":" & "C" & RowNum + 4).Select
    LeftSize
    FontSize
    MyExcel.Selection.Font.Bold = True
    
    MyExcel.Range("B" & RowNum + 4).Select
    MyExcel.ActiveCell.FormulaR1C1 = GsmDcs_n
    MyExcel.Rows(Format(RowNum + 4) & ":" & Format(RowNum + 4)).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    RowNum = RowNum + 6
    If Dcs_n = 0 Then GoTo Dcs
    
    '***************************'***双频互相切换统计表***
   ' MyExcel.Range("A" & RowNum).Select
   ' MyExcel.Application.CutCopyMode = False
   '  FontSize
   ' MyExcel.Selection.Font.Bold = True
    'MyExcel.Selection.Font.ColorIndex = 5
    ''MyExcel.ActiveCell.FormulaR1C1 = "...双频互相切换统计表..."
   '
   ' MyExcel.Range("A" & RowNum + 1).Select
    'FontSize
    'MyExcel.Selection.Font.ColorIndex = 10
    'MyExcel.ActiveCell.FormulaR1C1 = "(信令过程:HANDOVER COMMAND的ARFCN属GSM900,HANDOVER COMPLETE后的ARFCN属DCS1800,反之亦然)"
    'MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    'fontstyle
    'horversize
    'MyExcel.Selection.Font.Bold = True
    'MyExcel.Range("B" & RowNum + 2).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    'MyExcel.Range("C" & RowNum + 2).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "%"
    'MyExcel.Rows(RowNum + 3 & ":" & RowNum + 3).Select
    'fontsize
    'MyExcel.Range("A" & RowNum + 3).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "GSM900->DCS1800"
    'MyExcel.Range("B" & RowNum + 3 & ":" & "C" & RowNum + 3).Select
     'leftsize
     'MyExcel.Range("B" & RowNum + 3).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "" ' Gsm_n
    'MyExcel.Range("C" & RowNum + 3).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "" 'Gsm_n1
    '***********
    'MyExcel.Rows(RowNum + 4 & ":" & RowNum + 4).Select
    'fontsize
    'MyExcel.Range("A" & RowNum + 4).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "DCS1800->GSM900"
    'MyExcel.Range("B" & RowNum + 4 & ":" & "C" & RowNum + 4).Select
    'leftsize
    'MyExcel.Range("B" & RowNum + 4).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "" 'Dcs_n
    'MyExcel.Range("C" & RowNum + 4).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "" 'Dcs_n1
    'MyExcel.Rows(RowNum + 5 & ":" & RowNum + 5).Select
    'fontsize
    'MyExcel.Range("A" & RowNum + 5).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "总计"
    'MyExcel.Range("B" & RowNum + 5 & ":" & "C" & RowNum + 5).Select
    'leftsize
    'MyExcel.Range("B" & RowNum + 5).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "" 'GsmDcs_n
    'RowNum = RowNum + 7
  
  '****************************'Gsm900***手机发送功率统计表***
Dcs:
    MyExcel.Range("A" & RowNum).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...手机发送功率统计表..."
    
    MyExcel.Range("A" & RowNum + 1).Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.ActiveCell.FormulaR1C1 = "GSM900"
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    Gsm900Dcs1800Flag = False
    Call st_fill("0", "0", "27", True, True, True, "0 (43dBm)", False, RowNum + 3, "B")
    Call st_fill("1", "0", "27", False, True, True, "1 (41dBm)", False, RowNum + 4, "B")
    Call st_fill("2", "0", "27", False, True, True, "2 (39dBm)", False, RowNum + 5, "B")
    Call st_fill("3", "0", "27", False, True, True, "3 (37dBm)", False, RowNum + 6, "B")
    Call st_fill("4", "0", "27", False, True, True, "4 (35dBm)", False, RowNum + 7, "B")
    Call st_fill("5", "0", "27", False, True, True, "5 (33dBm)", False, RowNum + 8, "B")
    Call st_fill("6", "0", "27", False, True, True, "6 (31dBm)", False, RowNum + 9, "B")
    Call st_fill("7", "0", "27", False, True, True, "7 (29dBm)", False, RowNum + 10, "B")
    Call st_fill("8", "0", "27", False, True, True, "8 (27dBm)", False, RowNum + 11, "B")
    Call st_fill("9", "0", "27", False, True, True, "9 (25dBm)", False, RowNum + 12, "B")
    Call st_fill("10", "0", "27", False, True, True, "10 (23dBm)", False, RowNum + 13, "B")
    Call st_fill("11", "0", "27", False, True, True, "11 (21dBm)", False, RowNum + 14, "B")
    Call st_fill("12", "0", "27", False, True, True, "12 (19dBm)", False, RowNum + 15, "B")
    Call st_fill("13", "0", "27", False, True, True, "13 (17dBm)", False, RowNum + 16, "B")
    Call st_fill("14", "0", "27", False, True, True, "14 (15dBm)", False, RowNum + 17, "B")
    Call st_fill("15", "0", "27", False, True, True, "15 (13dBm)", False, RowNum + 18, "B")
    Call st_fill("16", "0", "27", False, True, True, "16 (11dBm)", False, RowNum + 19, "B")
    Call st_fill("17", "0", "27", False, True, True, "17 (9dBm)", False, RowNum + 20, "B")
    Call st_fill("18", "0", "27", False, True, True, "18 (7dBm)", False, RowNum + 21, "B")
    Call st_fill("19", "0", "27", False, True, True, "19 (5dBm)", False, RowNum + 22, "B")
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    
    MyExcel.Rows(RowNum + 23 & ":" & RowNum + 23).Select
    FontSize
    MyExcel.Range("A" & RowNum + 23).Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.ActiveCell.FormulaR1C1 = "总计"
    
    MyExcel.Range("B" & RowNum + 23 & ":" & "C" & RowNum + 23).Select
    LeftSize
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 23).Select
    MyExcel.ActiveCell.FormulaR1C1 = all_0
    MyExcel.Range("C" & RowNum + 23).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin1
    MyExcel.Rows(Format(RowNum + 23) & ":" & Format(RowNum + 23)).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    RowNum = RowNum + 24
    
    '****************************'***Gsm1800手机发送功率统计表***
    If Dcs_n = 0 Then GoTo TA
    MyExcel.Range("A" & RowNum + 1).Select
    FontSize
    MyExcel.Selection.Font.ColorIndex = 10
    MyExcel.ActiveCell.FormulaR1C1 = "DCS1800"
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    Gsm900Dcs1800Flag = True
    all_avg = 0
    Call st_fill("0", "0", "27", True, True, True, "0 (30dBm)", False, RowNum + 3, "B")
    Call st_fill("1", "0", "27", False, True, True, "1 (28dBm)", False, RowNum + 4, "B")
    Call st_fill("2", "0", "27", False, True, True, "2 (26dBm)", False, RowNum + 5, "B")
    Call st_fill("3", "0", "27", False, True, True, "3 (24dBm)", False, RowNum + 6, "B")
    Call st_fill("4", "0", "27", False, True, True, "4 (22dBm)", False, RowNum + 7, "B")
    Call st_fill("5", "0", "27", False, True, True, "5 (20dBm)", False, RowNum + 8, "B")
    Call st_fill("6", "0", "27", False, True, True, "6 (18dBm)", False, RowNum + 9, "B")
    Call st_fill("7", "0", "27", False, True, True, "7 (16dBm)", False, RowNum + 10, "B")
    Call st_fill("8", "0", "27", False, True, True, "8 (14dBm)", False, RowNum + 11, "B")
    Call st_fill("9", "0", "27", False, True, True, "9 (12dBm)", False, RowNum + 12, "B")
    Call st_fill("10", "0", "27", False, True, True, "10 (10dBm)", False, RowNum + 13, "B")
    Call st_fill("11", "0", "27", False, True, True, "11 (8dBm)", False, RowNum + 14, "B")
    Call st_fill("12", "0", "27", False, True, True, "12 (6dBm)", False, RowNum + 15, "B")
    Call st_fill("13", "0", "27", False, True, True, "13 (4dBm)", False, RowNum + 16, "B")
    Call st_fill("14", "0", "27", False, True, True, "14 (2dBm)", False, RowNum + 17, "B")
    Call st_fill("15", "0", "27", False, True, True, "15 (0dBm)", False, RowNum + 18, "B")
    all_0 = LTrim$(str(cc_all))
    If all_avg <> 0 Then
        all_avg = all_avg / cc_all
        putin1 = Format$(all_avg, "fixed")
    Else
        putin1 = 0
    End If
    MyExcel.Rows(RowNum + 19 & ":" & RowNum + 19).Select
    FontSize
    MyExcel.Range("A" & RowNum + 19).Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.ActiveCell.FormulaR1C1 = "总计"
    MyExcel.Range("B" & RowNum + 19 & ":" & "C" & RowNum + 19).Select
    LeftSize
    FontSize
    
    MyExcel.Selection.Font.Bold = True
    
    MyExcel.Range("B" & RowNum + 19).Select
    MyExcel.ActiveCell.FormulaR1C1 = all_0
    MyExcel.Range("C" & RowNum + 19).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin1
    MyExcel.Rows(RowNum + 19 & ":" & RowNum + 19).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    RowNum = RowNum + 20
   
   '*********************"***RXQUAL_FULL统计表****"
TA:
    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...RXQUAL_Full统计表..."
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    'MyExcel.Range("C" & RowNum + 2).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "最大值"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    'MyExcel.Range("E" & RowNum + 2).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "最小值"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    RxlevFullFlag = True
    Call st_fill("0", "0", "23", True, True, False, "0 (BER<0.2%)", False, RowNum + 3, "B")
    Call st_fill("1", "0", "23", False, True, False, "1 (0.2%<BER<0.4%)", False, RowNum + 4, "B")
    Call st_fill("2", "0", "23", False, True, False, "2 (0.4%<BER<0.8%)", False, RowNum + 5, "B")
    Call st_fill("3", "0", "23", False, True, False, "3 (0.8%<BER<1.6%)", False, RowNum + 6, "B")
    Call st_fill("4", "0", "23", False, True, False, "4 (1.6%<BER<3.2%)", False, RowNum + 7, "B")
    Call st_fill("5", "0", "23", False, True, False, "5 (3.2%<BER<6.4%)", False, RowNum + 8, "B")
    Call st_fill("6", "0", "23", False, True, False, "6 (6.4%<BER<12.8%)", False, RowNum + 9, "B")
    Call st_fill("7", "0", "23", False, True, False, "7 (12.8%<BER)", False, RowNum + 10, "B")
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    MyExcel.Rows(RowNum + 11 & ":" & RowNum + 11).Select
    FontSize
    MyExcel.Range("A" & RowNum + 11).Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.ActiveCell.FormulaR1C1 = "总计"
    
    MyExcel.Range("B" & RowNum + 11 & ":" & "C" & RowNum + 11).Select
    LeftSize
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 11).Select
    MyExcel.ActiveCell.FormulaR1C1 = all_0
    MyExcel.Range("C" & RowNum + 11).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin1
    MyExcel.Rows(RowNum + 11 & ":" & RowNum + 11).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    RowNum = RowNum + 12
     '*********************"***RXQUAL_SUB统计表****"
    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...RXQUAL_SUB统计表..."
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    'MyExcel.Range("C" & RowNum + 2).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "最大值"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    'MyExcel.Range("E" & RowNum + 2).Select
    'MyExcel.ActiveCell.FormulaR1C1 = "最小值"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    RxlevFullFlag = False
    Call st_fill("0", "0", "23", True, True, False, "0 (BER<0.2%)", False, RowNum + 3, "B")
    Call st_fill("1", "0", "23", False, True, False, "1 (0.2%<BER<0.4%)", False, RowNum + 4, "B")
    Call st_fill("2", "0", "23", False, True, False, "2 (0.4%<BER<0.8%)", False, RowNum + 5, "B")
    Call st_fill("3", "0", "23", False, True, False, "3 (0.8%<BER<1.6%)", False, RowNum + 6, "B")
    Call st_fill("4", "0", "23", False, True, False, "4 (1.6%<BER<3.2%)", False, RowNum + 7, "B")
    Call st_fill("5", "0", "23", False, True, False, "5 (3.2%<BER<6.4%)", False, RowNum + 8, "B")
    Call st_fill("6", "0", "23", False, True, False, "6 (6.4%<BER<12.8%)", False, RowNum + 9, "B")
    Call st_fill("7", "0", "23", False, True, False, "7 (12.8%<BER)", False, RowNum + 10, "B")
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    MyExcel.Rows(RowNum + 11 & ":" & RowNum + 11).Select
    FontSize
    MyExcel.Range("A" & RowNum + 11).Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.ActiveCell.FormulaR1C1 = "总计"
    
    MyExcel.Range("B" & RowNum + 11 & ":" & "C" & RowNum + 11).Select
    LeftSize
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 11).Select
    MyExcel.ActiveCell.FormulaR1C1 = all_0
    MyExcel.Range("C" & RowNum + 11).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin1
    MyExcel.Rows(RowNum + 11 & ":" & RowNum + 11).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    RowNum = RowNum + 12
    
    '********************************'RXLEV_F统计表
    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...RXLEV_FULL统计表..."
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最大值"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最小值"
    MyExcel.Range("F" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("G" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    RxlevFullFlag = True
    RangeNum = 3
    RxLevRange(1, 1) = "27"
    RxLevRange(1, 2) = "17"
    RxLevRange(1, 3) = "0"
    RxLevRange(2, 1) = "63"
    RxLevRange(2, 2) = "27"
    RxLevRange(2, 3) = "17"
    '************
    For i = 1 To RangeNum
        If i = 1 Then
            Call st_fill(Format(Val(RxLevRange(1, i))), RxLevRange(2, i), "22", True, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 3, "B")
        Else
            Call st_fill(Format(Val(RxLevRange(1, i))), RxLevRange(2, i), "22", False, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 2 + i, "B")
        End If
    Next
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    MyExcel.Rows(RowNum + 3 + i & ":" & RowNum + 3 + i).Select
    FontSize
    MyExcel.Range("A" & RowNum + 3 + i).Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.ActiveCell.FormulaR1C1 = "总计"
    MyExcel.Range("B" & RowNum + 3 + i & ":" & "E" & RowNum + 3 + i).Select
    LeftSize
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 3 + i).Select
    MyExcel.ActiveCell.FormulaR1C1 = all_0
    MyExcel.Range("C" & RowNum + 3 + i).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin
    MyExcel.Range("D" & RowNum + 3 + i).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin1
    MyExcel.Range("E" & RowNum + 3 + i).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin2
    MyExcel.Rows(RowNum + 3 + i & ":" & RowNum + 3 + i).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    RowNum = RowNum + 4 + i
    '***************************
    
    '********************************'RXLEV_S统计表
    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...RXLEV_SUB统计表..."
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最大值"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最小值"
    MyExcel.Range("F" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("G" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    RxlevFullFlag = False
    RangeNum = 3
    RxLevRange(1, 1) = "27"
    RxLevRange(1, 2) = "17"
    RxLevRange(1, 3) = "0"
    RxLevRange(2, 1) = "63"
    RxLevRange(2, 2) = "27"
    RxLevRange(2, 3) = "17"
    '************
    For i = 1 To RangeNum
        If i = 1 Then
            Call st_fill(Format(Val(RxLevRange(1, i))), RxLevRange(2, i), "22", True, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 3, "B")
        Else
            Call st_fill(Format(Val(RxLevRange(1, i))), RxLevRange(2, i), "22", False, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 2 + i, "B")
        End If
    Next
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    MyExcel.Rows(RowNum + 3 + i & ":" & RowNum + 3 + i).Select
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("A" & RowNum + 3 + i).Select
    MyExcel.ActiveCell.FormulaR1C1 = "总计"
    MyExcel.Range("B" & RowNum + 3 + i & ":" & "E" & RowNum + 3 + i).Select
    LeftSize
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 3 + i).Select
    MyExcel.ActiveCell.FormulaR1C1 = all_0
    MyExcel.Range("C" & RowNum + 3 + i).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin
    MyExcel.Range("D" & RowNum + 3 + i).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin1
    MyExcel.Range("E" & RowNum + 3 + i).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin2
    MyExcel.Rows(RowNum + 3 + i & ":" & RowNum + 3 + i).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    RowNum = RowNum + 4 + i
    '***************************
    '********************************'***Timing Advance(TA)统计表***

    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...Timing Advance(TA)统计表..."
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最大值"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最小值"
    MyExcel.Range("F" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("G" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    TaFlag = True
    Call st_fill("0", "0", "26", True, False, True, "x=0", False, RowNum + 3, "B")
    Call st_fill("1", "0", "26", False, False, True, "x=1", False, RowNum + 4, "B")
    Call st_fill("2", "0", "26", False, False, True, "x=2", False, RowNum + 5, "B")
    Call st_fill("3", "0", "26", False, False, True, "x=3", False, RowNum + 6, "B")
    Call st_fill("4", "0", "26", False, False, True, "x=4", False, RowNum + 7, "B")
    Call st_fill("5", "0", "26", False, False, True, "x=5", False, RowNum + 8, "B")
    Call st_fill("6", "0", "26", False, False, True, "x=6", False, RowNum + 9, "B")
    Call st_fill("7", "0", "26", False, False, True, "x=7", False, RowNum + 10, "B")
    Call st_fill("7", "30", "26", False, False, True, "7<X<=30", False, RowNum + 11, "B")
    Call st_fill("30", "63", "26", False, False, True, "30<X<=63", False, RowNum + 12, "B")
    all_0 = LTrim$(str(cc_all))
    all_0 = all_0 + space$(9 - Len(all_0))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    MyExcel.Rows(RowNum + 13 & ":" & RowNum + 13).Select
    FontSize
    MyExcel.Range("A" & RowNum + 13).Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.ActiveCell.FormulaR1C1 = "总计"
    
    MyExcel.Range("B" & RowNum + 13 & ":" & "E" & RowNum + 13).Select
    LeftSize
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 13).Select
    MyExcel.ActiveCell.FormulaR1C1 = all_0
    MyExcel.Range("C" & RowNum + 13).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin
    MyExcel.Range("D" & RowNum + 13).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin1
    MyExcel.Range("E" & RowNum + 13).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin2
    MyExcel.Rows(RowNum + 13 & ":" & RowNum + 13).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    RowNum = RowNum + 14
    '***************************
    '********************************'***小区选择参数C1统计表**
    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...小区选择参数C1统计表... "
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最大值"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最小值"
    MyExcel.Range("F" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("G" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    TaFlag = False
    C1C2Flag = True
    Call st_fill("0", "0", "26", True, False, True, "C1<0", False, RowNum + 3, "B")
    Call st_fill("1", "0", "26", True, False, True, "C1=0", False, RowNum + 4, "B")
    Call st_fill("1", "10", "26", False, False, True, "1=<C1<10", False, RowNum + 5, "B")
    Call st_fill("10", "20", "26", False, False, True, "10=<C1<20", False, RowNum + 6, "B")
    Call st_fill("20", "30", "26", False, False, True, "20=<C1<30", False, RowNum + 7, "B")
    Call st_fill("30", "40", "26", False, False, True, "30=<C1<40", False, RowNum + 8, "B")
    Call st_fill("40", "50", "26", False, False, True, "40=<C1<50", False, RowNum + 9, "B")
    Call st_fill("50", "60", "26", False, False, True, "50=<C1<60", False, RowNum + 10, "B")
    Call st_fill("2", "60", "26", False, False, True, "C1>=60", False, RowNum + 11, "B")
    all_0 = LTrim$(str(cc_all))
    all_0 = all_0 + space$(9 - Len(all_0))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    MyExcel.Rows(RowNum + 12 & ":" & RowNum + 12).Select
    FontSize
    MyExcel.Range("A" & RowNum + 12).Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.ActiveCell.FormulaR1C1 = "总计"
    MyExcel.Range("B" & RowNum + 12 & ":" & "E" & RowNum + 12).Select
    LeftSize
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 12).Select
    MyExcel.ActiveCell.FormulaR1C1 = all_0
    MyExcel.Range("C" & RowNum + 12).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin
    MyExcel.Range("D" & RowNum + 12).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin1
    MyExcel.Range("E" & RowNum + 12).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin2
    MyExcel.Rows(RowNum + 12 & ":" & RowNum + 12).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    RowNum = RowNum + 13
    
    '********************************'***小区选择参数C2统计表**
    MyExcel.Range("A" & RowNum + 1).Select
    MyExcel.Application.CutCopyMode = False
    FontSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Selection.Font.ColorIndex = 5
    MyExcel.ActiveCell.FormulaR1C1 = "...小区选择参数C2统计表... "
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Select
    FontStyle
    HorVerSize
    MyExcel.Selection.Font.Bold = True
    MyExcel.Range("B" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "测量数"
    MyExcel.Range("C" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最大值"
    MyExcel.Range("D" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "均值"
    MyExcel.Range("E" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "最小值"
    MyExcel.Range("F" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "%"
    MyExcel.Range("G" & RowNum + 2).Select
    MyExcel.ActiveCell.FormulaR1C1 = "累结%"
    TaFlag = False
    C1C2Flag = False
    Call st_fill("0", "0", "26", True, False, True, "C2<0", False, RowNum + 3, "B")
    Call st_fill("1", "0", "26", True, False, True, "C2=0", False, RowNum + 4, "B")
    Call st_fill("1", "10", "26", False, False, True, "1=<C2<10", False, RowNum + 5, "B")
    Call st_fill("10", "20", "26", False, False, True, "10=<C2<20", False, RowNum + 6, "B")
    Call st_fill("20", "30", "26", False, False, True, "20=<C2<30", False, RowNum + 7, "B")
    Call st_fill("30", "40", "26", False, False, True, "30=<C2<40", False, RowNum + 8, "B")
    Call st_fill("40", "50", "26", False, False, True, "40=<C2<50", False, RowNum + 9, "B")
    Call st_fill("50", "60", "26", False, False, True, "50=<C2<60", False, RowNum + 10, "B")
    Call st_fill("60", "80", "26", False, False, True, "60=<C2<80", False, RowNum + 11, "B")
    Call st_fill("80", "100", "26", False, False, True, "80=<C2<100", False, RowNum + 12, "B")
    Call st_fill("100", "150", "26", False, False, True, "100=<C2<150", False, RowNum + 13, "B")
    Call st_fill("150", "200", "26", False, False, True, "150=<C2<200", False, RowNum + 14, "B")
    Call st_fill("2", "200", "26", False, False, True, "C2>=200", False, RowNum + 15, "B")
    
    all_0 = LTrim$(str(cc_all))
    all_0 = all_0 + space$(9 - Len(all_0))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    MyExcel.Rows(RowNum + 16 & ":" & RowNum + 16).Select
    FontSize
    MyExcel.Range("A" & RowNum + 16).Select
    MyExcel.Selection.Font.Bold = True
    MyExcel.ActiveCell.FormulaR1C1 = "总计"
    MyExcel.Range("B" & RowNum + 16 & ":" & "E" & RowNum + 16).Select
    LeftSize
    FontSize
    MyExcel.Selection.Font.Bold = True
    
    MyExcel.Range("B" & RowNum + 16).Select
    MyExcel.ActiveCell.FormulaR1C1 = all_0
    MyExcel.Range("C" & RowNum + 16).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin
    MyExcel.Range("D" & RowNum + 16).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin1
    MyExcel.Range("E" & RowNum + 16).Select
    MyExcel.ActiveCell.FormulaR1C1 = putin2
    MyExcel.Rows(RowNum + 16 & ":" & RowNum + 16).Select
    With MyExcel.Selection.Interior
        .ColorIndex = 15
        .Pattern = 1
    End With
    
    'RowNum = RowNum + 17
    MyExcel.Sheets("呼叫统计").Select
    MyExcel.Range("A1").Select
    
    
    '**************
    'MyExcel.ChangeFileOpenDirectory gsm_path + "\user\"
    MyExcel.ActiveWorkbook.Saveas filename:=stcname
      
End Sub
Sub LeftSize()
    On Error Resume Next
    With MyExcel.Selection
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4107 ' xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
End Sub
Sub ArialSize()
    On Error Resume Next
    With MyExcel.Selection.Font
        .Name = "Arial"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = -4142
       ' .ColorIndex = xlAutomatic
    End With
    
    
End Sub
Sub FontSize()
   On Error Resume Next
   With MyExcel.Selection.Font
        .Name = "宋体"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = -4142
        '.ColorIndex = xlAutomatic
   End With
End Sub

Sub FontStyle()
    With MyExcel.Selection.Font
        .Name = "宋体"
        .FontStyle = "加粗"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = -4142
        '.ColorIndex = xlAutomatic
    End With
End Sub
Sub HorVerSize()
    On Error Resume Next
    With MyExcel.Selection
        .HorizontalAlignment = -4108
        .VerticalAlignment = -4107
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
End Sub
Sub BoldfacedSize()
   On Error Resume Next
   With MyExcel.Selection.Font
        .Name = "黑体"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = -4142
        .ColorIndex = 16
  End With
    
End Sub

Sub Call_Attemp(RowFirst As Integer)
    Dim h_comm() As Single, h_comp() As Single
    Dim menum, tt, all
    Dim dd As String
    Dim hav As Boolean
    Dim ggg As Single
    Dim Mystring As String
    Dim leefind As Integer
    Dim mid_dd As Long
    Dim i As Integer, j As Integer, assi As Integer, assj As Integer, assg As Integer, assb As Integer, assd As Integer
    Dim k As Integer, HH As Integer, finds As Integer, tal As Integer, pp As Integer
    Dim MyTableName As String
    Dim MyDbName As String
    Dim dbs As Database, rst As Recordset
    Dim lngRecords As Long, lngFields As Long
    Dim lngRecCount As Long, lngFldCount As Long
    ReDim h_comm(1 To 1000) As Single
    ReDim h_comp(1 To 1000) As Single
    ReDim h_comg(1 To 1000) As Single
    ReDim h_comd(1 To 1000) As Single
    ReDim h_comb(1 To 1000) As Single
    Dim Result As String
    Dim h_starttime(1 To 1000) As String
    Dim h_stoptime(1 To 1000) As String
    Dim assnum As Integer
    On Error GoTo errend
   
    menum = 0
    assnum = 0
    For HH = 1 To stre_num
        For pp = 1 To 1000
            h_comm(pp) = 0
            h_comp(pp) = 0
            h_comg(pp) = 0
            h_comd(pp) = 0
            h_comb(pp) = 0
            h_starttime(pp) = ""
            h_stoptime(pp) = ""
            
        Next
        
        MyTableName = convert_filename(HH)
        MyDbName = ""
        Do While InStr(MyTableName, "\") > 0
            MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
            MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
        Loop
        MyDbName = Left(MyDbName, Len(MyDbName) - 1)
        Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "SETUP" & """ or message = """ & "EMERGENCY SETUP" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
                menum = lngRecords
                If menum > 0 Then
                    rst.MoveFirst
                    For i = 1 To menum
                        dd = rst.Fields(0).Value
                        h_starttime(i) = dd
                        finds = InStr(dd, ":")
                        mid_dd = Val(Left(dd, finds - 1)) * 3600
                        h_comm(i) = mid_dd
                        dd = Right(dd, Len(dd) - finds)
                        finds = InStr(dd, ":")
                        If finds > 0 Then
                            mid_dd = Val(Left(dd, finds - 1)) * 60
                            h_comm(i) = h_comm(i) + mid_dd
                            dd = Right(dd, Len(dd) - finds)
                        End If
                        finds = InStr(dd, ".")
                        If finds > 0 Then
                            mid_dd = Val(Left(dd, finds - 1))
                            dd = Right(dd, Len(dd) - finds)
                        Else
                            mid_dd = Val(dd)
                            dd = 0
                        End If
                        h_comm(i) = h_comm(i) + mid_dd
                        h_comm(i) = h_comm(i) + Val(dd) / 100
                        If i < menum Then
                            rst.MoveNext
                        End If
                    Next
                    h_comm(i) = -1
                End If
    '***********
          'Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
        
          Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "CHANNEL RELEASE" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
               menum = lngRecords
        If menum > 0 Then
           rst.MoveFirst
           For i = 1 To menum
               dd = rst.Fields(0).Value
               h_stoptime(i) = dd
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 3600
               h_comp(i) = mid_dd
               dd = Right(dd, Len(dd) - finds)
               finds = InStr(dd, ":")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1)) * 60
                    h_comp(i) = h_comp(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
               End If
               finds = InStr(dd, ".")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1))
                    h_comp(i) = h_comp(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
                    h_comp(i) = h_comp(i) + Val(dd) / 100
               Else
                    h_comp(i) = h_comp(i) + Val(dd)
                    dd = 0
               End If
               If i < menum Then
                    rst.MoveNext
               End If
           Next
'           dbs.Close
           h_comp(i) = -1
        End If
        '************
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timem FROM " & MyTableName & " Where  message = """ & "DISCONNECT" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
                  menum = lngRecords
        If menum > 0 Then
           rst.MoveFirst
           For i = 1 To menum
               dd = rst.Fields(0).Value
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 3600
               h_comg(i) = mid_dd
               dd = Right(dd, Len(dd) - finds)
               finds = InStr(dd, ":")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1)) * 60
                    h_comg(i) = h_comg(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
               End If
               finds = InStr(dd, ".")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1))
                    h_comg(i) = h_comg(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
                    h_comg(i) = h_comg(i) + Val(dd) / 100
               Else
                    h_comg(i) = h_comg(i) + Val(dd)
                    dd = 0
               End If
               If i < menum Then
                    rst.MoveNext
               End If
           Next
 '          dbs.Close
           h_comg(i) = -1
        End If
        
        '**********
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  mark = """ & "Dropped Call" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
               menum = lngRecords
        
        If menum > 0 Then
           rst.MoveFirst
           For i = 1 To menum
               dd = rst.Fields(0).Value
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 3600
               h_comd(i) = mid_dd
               dd = Right(dd, Len(dd) - finds)
               finds = InStr(dd, ":")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1)) * 60
                    h_comd(i) = h_comd(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
               End If
               finds = InStr(dd, ".")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1))
                    h_comd(i) = h_comd(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
                    h_comd(i) = h_comd(i) + Val(dd) / 100
               Else
                    h_comd(i) = h_comd(i) + Val(dd)
                    dd = 0
               End If
               If i < menum Then
                    rst.MoveNext
               End If
           Next
  '         dbs.Close
           h_comd(i) = -1
        End If
        
        '***********
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  mark = """ & "Blocked Call" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
               menum = lngRecords
        If menum > 0 Then
           rst.MoveFirst
           For i = 1 To menum
               dd = rst.Fields(0).Value
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 3600
               h_comb(i) = mid_dd
               dd = Right(dd, Len(dd) - finds)
               finds = InStr(dd, ":")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1)) * 60
                    h_comb(i) = h_comb(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
               End If
               finds = InStr(dd, ".")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1))
                    h_comb(i) = h_comb(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
                    h_comb(i) = h_comb(i) + Val(dd) / 100
               Else
                    h_comb(i) = h_comb(i) + Val(dd)
                    dd = 0
               End If
               If i < menum Then
                    rst.MoveNext
               End If
           Next
           dbs.Close
           h_comb(i) = -1
        End If
        
        '*******
           assi = 1
           assj = 1
           assb = 1
           assg = 1
           assd = 1
           Do While h_comm(assi) <> -1 And h_comm(assi) <> 0
              If h_comm(assi) <= h_comp(assj) Then
                 If h_comm(assi + 1) > h_comp(assj) Or h_comm(assi + 1) = -1 Or h_comm(assi + 1) = 0 Then
                    
                    '****************
                    leefind = InStr(MyTableName, ".")
                    Mystring = Left(MyTableName, leefind - 1)
                    MyExcel.Rows(Format(RowFirst)).Select
                    FontSize
                    HorVerSize
                    MyExcel.Range("A" & Format(RowFirst)).Select
                    MyExcel.ActiveCell.FormulaR1C1 = Mystring
                    MyExcel.Range("B" & Format(RowFirst)).Select
                    HorVerSize
                    MyExcel.Selection.Font.Bold = False
                    MyExcel.Range("B" & Format(RowFirst)).Select
                    MyExcel.ActiveCell.FormulaR1C1 = h_starttime(assi)
                
                    MyExcel.Range("C" & Format(RowFirst)).Select
                    HorVerSize
                    MyExcel.Selection.Font.Bold = False
                    MyExcel.ActiveCell.FormulaR1C1 = h_stoptime(assj)
                
                    MyExcel.Range("D" & Format(RowFirst)).Select
                    HorVerSize
                    MyExcel.Selection.Font.Bold = False
                    MyExcel.ActiveCell.FormulaR1C1 = h_comp(assj) - h_comm(assi)
                
                    If h_comg(assj) <= h_comp(assj) And h_comg(assj) <> -1 And h_comg(assi) <> 0 Then
                        Result = "正常"
                    Else
                    If h_comb(assj) = h_comp(assj) And h_comb(assj) <> 0 Then
                        Result = "拥塞"
                    ElseIf h_comd(assj) = h_comp(assj) And h_comd(assj) <> 0 Then
                        Result = "掉话"
                    End If
                End If
                MyExcel.Range("E" & Format(RowFirst)).Select
                HorVerSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.ActiveCell.FormulaR1C1 = Result
                
                RowFirst = RowFirst + 1
                
                              
                    assi = assi + 1
                    assj = assj + 1
                 Else
                    assi = assi + 1
                    If h_comm(assi) = -1 Or h_comm(assi) = 0 Then Exit Do
                 End If
              Else
                 assj = assj + 1
                 If h_comp(assj) = -1 Or h_comp(assj) = 0 Then Exit Do
              End If
           Loop
    
    Next
    '******************************************
    CellRownum = RowFirst
   
errend:
    
End Sub

Sub st_fill(a As String, b As String, col As String, ByVal sta As Boolean, ByVal rxq As Boolean, ByVal va As Boolean, ByVal fillhead As String, ByVal x9 As Boolean, ByVal RowFirst As Integer, ByVal RowNumber As String)
    Dim num, er_max, er_avg, er_min, er3, er4, er
    Dim num_z As Integer, max_z As Integer, avg_z As Single, min_z As Integer
    Dim msg As String
    Dim fillnum As Integer, j As Integer
    Dim f1 As Integer, f2 As Integer
    Dim Zero As Boolean
    Dim myFilename As String
    Dim MyDbName As String, MyTableName As String
    Dim dbs As Database, rst As Recordset
    Dim lngRecords As Long, lngFields As Long
    Dim lngRecCount As Long, lngFldCount As Long
    Dim i As Integer
    Dim Mystring As String
    Static perc
    On Error GoTo errend
    DoEvents
    Zero = True
    num_z = 0
    max_z = 0
    If sta = True Then
       all_max = 0
       all_min = 0
       all_avg = 0
       perc = 0
    End If
    If perc = 1 Then
       MyExcel.Rows(Format(RowFirst)).Select
       FontSize
       MyExcel.Range("A" & Format(RowFirst)).Select
       MyExcel.ActiveCell.FormulaR1C1 = fillhead
    
       'Print #FileNumber, fillhead
       Exit Sub
    End If
    
    For j = 1 To stre_num
        MyTableName = convert_filename(j)
        MyDbName = ""
        Do While InStr(MyTableName, "\") > 0
            MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
            MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
        Loop
        MyDbName = Left(MyDbName, Len(MyDbName) - 1)
        Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
        If rxq = True Then
          If va = True Then
             'fillnum = 15
             If Not Gsm900Dcs1800Flag Then
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(TX_POWER) as counttxpower ," _
                    & " Avg(val(TX_POWER)) " _
                    & "AS Averagta FROM " & MyTableName & " Where  bcch_serv < 125 and TX_POWER = """ & a & """") 'TX_POWER = """ & a & """")
             Else
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(TX_POWER) as counttxpower ," _
                    & " Avg(val(TX_POWER)) " _
                    & "AS Averagta FROM " & MyTableName & " Where  bcch_serv > 511 and TX_POWER = """ & a & """")
             End If
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
                num = rst.Fields(0).Value
                'er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                'er_min = rst.Fields(3).Value
                dbs.Close
          Else
             fillnum = 22 'rxqual
             If RxlevFullFlag Then
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(rxqual_f) as countrxqual ," _
                    & " Avg(rxqual_f) " _
                    & "AS Averagerxqual FROM " & MyTableName & " Where  rxqual_f = " & a)
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
                num = rst.Fields(0).Value
                'er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                'er_min = rst.Fields(3).Value
                dbs.Close
             Else
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(rxqual_s) as countrxqual ," _
                    & " Avg(rxqual_s) " _
                    & "AS Averagerxqual FROM " & MyTableName & " Where  rxqual_s = " & a)
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
                num = rst.Fields(0).Value
                'er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                'er_min = rst.Fields(3).Value
                dbs.Close
             
             End If
          End If
       Else
          If va = True Then
            If TaFlag Then
              'fillnum = 17
              
              If b = "0" Then
                 Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(ta) as countta ," _
                    & " Avg(val(ta)) " _
                    & "AS Averagta, Max(val(ta)) " _
                    & "AS Maximumta, Min(val(ta)) " _
                    & "As Minta FROM " & MyTableName & " Where  ta = """ & a & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
                num = rst.Fields(0).Value
                er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                er_min = rst.Fields(3).Value
                dbs.Close
              Else
                 Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(ta) as countta ," _
                    & " Avg(val(ta)) " _
                    & "AS Averagta, Max(val(ta)) " _
                    & "AS Maximumta, Min(val(ta)) " _
                    & "As Minta FROM " & MyTableName & " Where  int(ta) > " & a & "" & " and int(ta) <= " & b & "")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount '
                lngFields = rst.Fields.Count
                num = rst.Fields(0).Value
                er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                er_min = rst.Fields(3).Value
                dbs.Close
              End If
            Else
              If C1C2Flag Then
              If b = "0" Then
                    If a = "0" Then
                       Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c1) as countta ," _
                        & " Avg(val(c1)) " _
                        & "AS Averagta, Max(val(c1)) " _
                        & "AS Maximumta, Min(val(c1)) " _
                        & "As Minta FROM " & MyTableName & " Where  int(c1) < """ & a & """")
                        If rst.RecordCount <> 0 Then
                            rst.MoveLast
                        End If
                        lngRecords = rst.RecordCount
                        lngFields = rst.Fields.Count
                        num = rst.Fields(0).Value
                        er_max = rst.Fields(2).Value
                        er_avg = rst.Fields(1).Value
                        er_min = rst.Fields(3).Value
                        dbs.Close
                    Else
                        Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c1) as countta ," _
                        & " Avg(val(c1)) " _
                        & "AS Averagta, Max(val(c1)) " _
                        & "AS Maximumta, Min(val(c1)) " _
                        & "As Minta FROM " & MyTableName & " Where  c1 = """ & a & """")
                        If rst.RecordCount <> 0 Then
                            rst.MoveLast
                        End If
                        lngRecords = rst.RecordCount
                        lngFields = rst.Fields.Count
                        num = rst.Fields(0).Value
                        er_max = rst.Fields(2).Value
                        er_avg = rst.Fields(1).Value
                        er_min = rst.Fields(3).Value
                        dbs.Close
                    
                    End If
              Else
                 If a <> "2" Then
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c1) as countta ," _
                        & " Avg(val(c1)) " _
                        & "AS Averagta, Max(val(c1)) " _
                        & "AS Maximumta, Min(val(c1)) " _
                        & "As Minta FROM " & MyTableName & " Where  int(c1) >= " & a & "" & " and int(c1) < " & b & "")
                    If rst.RecordCount <> 0 Then
                        rst.MoveLast
                    End If
                    lngRecords = rst.RecordCount '
                    lngFields = rst.Fields.Count
                    num = rst.Fields(0).Value
                    er_max = rst.Fields(2).Value
                    er_avg = rst.Fields(1).Value
                    er_min = rst.Fields(3).Value
                    dbs.Close
                Else
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c1) as countta ," _
                        & " Avg(val(c1)) " _
                        & "AS Averagta, Max(val(c1)) " _
                        & "AS Maximumta, Min(val(c1)) " _
                        & "As Minta FROM " & MyTableName & " Where  int(c1) >= """ & b & """")
                    If rst.RecordCount <> 0 Then
                        rst.MoveLast
                    End If
                    lngRecords = rst.RecordCount '
                    lngFields = rst.Fields.Count
                    num = rst.Fields(0).Value
                    er_max = rst.Fields(2).Value
                    er_avg = rst.Fields(1).Value
                    er_min = rst.Fields(3).Value
                    dbs.Close
                
                
                End If
              End If
              
              
              Else
                 If b = "0" Then
                    If a = "0" Then
                       Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c2) as countta ," _
                        & " Avg(val(c2)) " _
                        & "AS Averagta, Max(val(c2)) " _
                        & "AS Maximumta, Min(val(c2)) " _
                        & "As Minta FROM " & MyTableName & " Where  int(c2) < """ & a & """")
                        If rst.RecordCount <> 0 Then
                            rst.MoveLast
                        End If
                        lngRecords = rst.RecordCount
                        lngFields = rst.Fields.Count
                        num = rst.Fields(0).Value
                        er_max = rst.Fields(2).Value
                        er_avg = rst.Fields(1).Value
                        er_min = rst.Fields(3).Value
                        dbs.Close
                    Else
                        Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c2) as countta ," _
                        & " Avg(val(c2)) " _
                        & "AS Averagta, Max(val(c2)) " _
                        & "AS Maximumta, Min(val(c2)) " _
                        & "As Minta FROM " & MyTableName & " Where  c2 = """ & a & """")
                        If rst.RecordCount <> 0 Then
                            rst.MoveLast
                        End If
                        lngRecords = rst.RecordCount
                        lngFields = rst.Fields.Count
                        num = rst.Fields(0).Value
                        er_max = rst.Fields(2).Value
                        er_avg = rst.Fields(1).Value
                        er_min = rst.Fields(3).Value
                        dbs.Close
                    
                    End If
              Else
                 If a <> "2" Then
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c2) as countta ," _
                        & " Avg(val(c2)) " _
                        & "AS Averagta, Max(val(c2)) " _
                        & "AS Maximumta, Min(val(c2)) " _
                        & "As Minta FROM " & MyTableName & " Where  int(c2) >= " & a & "" & " and int(c2) < " & b & "")
                    If rst.RecordCount <> 0 Then
                        rst.MoveLast
                    End If
                    lngRecords = rst.RecordCount '
                    lngFields = rst.Fields.Count
                    num = rst.Fields(0).Value
                    er_max = rst.Fields(2).Value
                    er_avg = rst.Fields(1).Value
                    er_min = rst.Fields(3).Value
                    dbs.Close
                Else
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c2) as countta ," _
                        & " Avg(val(c2)) " _
                        & "AS Averagta, Max(val(c2)) " _
                        & "AS Maximumta, Min(val(c2)) " _
                        & "As Minta FROM " & MyTableName & " Where  int(c2) >= " & b & "")
                    If rst.RecordCount <> 0 Then
                        rst.MoveLast
                    End If
                    lngRecords = rst.RecordCount '
                    lngFields = rst.Fields.Count
                    num = rst.Fields(0).Value
                    er_max = rst.Fields(2).Value
                    er_avg = rst.Fields(1).Value
                    er_min = rst.Fields(3).Value
                    dbs.Close
                
                
                End If
              End If
                
              
              End If
            
            
            
            End If
          Else
             'fillnum = 24
             If RxlevFullFlag Then
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(rxlev_f) as countrxlev ," _
                    & " Avg(rxlev_f) " _
                    & "AS Averagerxlev, Max(rxlev_f) " _
                    & "AS Maximumrxlev,Min(rxlev_f) " _
                    & " as Minrxlev FROM " & MyTableName & " Where  rxlev_f >= " & Format(a) & " and  rxlev_f < " & Format(b))
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
                num = rst.Fields(0).Value
                er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                er_min = rst.Fields(3).Value
                dbs.Close
             Else
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(rxlev_s) as countrxlev ," _
                    & " Avg(rxlev_s) " _
                    & "AS Averagerxlev, Max(rxlev_s) " _
                    & "AS Maximumrxlev,Min(rxlev_s) " _
                    & " as Minrxlev FROM " & MyTableName & " Where  rxlev_s >= " & Format(a) & " and  rxlev_s < " & Format(b))
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
                num = rst.Fields(0).Value
                er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                er_min = rst.Fields(3).Value
                dbs.Close
             End If
          End If
       End If
      
       num_z = num + num_z
       If Val(num) > 0 Then
          If Zero = True Then
             max_z = er_max
             avg_z = Val(er_avg) * num
             min_z = er_min
          Else
             If max_z < er_max Then max_z = er_max
             If min_z > er_min Then min_z = er_min
             avg_z = avg_z + Val(er_avg) * num
          End If
          Zero = False
       End If
   Next
   'fillnum = fillnum - Len(fillhead)
   If Zero = False Then
      If perc = 0 Then
         all_max = max_z
         all_min = min_z
         all_avg = avg_z
      Else
         If all_max < max_z Then all_max = max_z
         If all_min > min_z Then all_min = min_z
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
   End If
      'If Val(num) > 0 Then
      '  num = num + Space$(9 - Len(num))
      '  er_max = er_max + Space$(7 - Len(er))
      '  er_avg = er_avg + Space$(9 - Len(er_avg))
      '  er_min = er_min + Space$(9 - Len(er_min))
      '  er3 = er3 + Space$(11 - Len(er3))
      'End If
      If x9 = True Then
'         Print #FileNumber, fillhead; er3
      Else
         If Zero = True Then
            MyExcel.Rows(Format(RowFirst)).Select
            FontSize
            MyExcel.Range("A" & Format(RowFirst)).Select
            MyExcel.ActiveCell.FormulaR1C1 = fillhead
    
            'Print #FileNumber, fillhead
         Else
            If rxq = True Then
                '************
                MyExcel.Rows(Format(RowFirst)).Select
                FontSize
                MyExcel.Range("A" & Format(RowFirst)).Select
                MyExcel.ActiveCell.FormulaR1C1 = fillhead
                MyExcel.Range(RowNumber & Format(RowFirst)).Select
                LeftSize
                FontSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.Range(RowNumber & Format(RowFirst)).Select
                MyExcel.ActiveCell.FormulaR1C1 = num
                
                Mystring = Asc(RowNumber)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.Range(Mystring & Format(RowFirst)).Select
                LeftSize
                FontSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.ActiveCell.FormulaR1C1 = er_avg
                   
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.Range(Mystring & Format(RowFirst)).Select
                LeftSize
                FontSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.ActiveCell.FormulaR1C1 = er3
                   
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.Range(Mystring & Format(RowFirst)).Select
                LeftSize
                FontSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.ActiveCell.FormulaR1C1 = er4
                  
                  'Print #FileNumber, fillhead; Space$(fillnum); num; er_avg; er3; er4
                
                '*********
            Else
                MyExcel.Rows(Format(RowFirst)).Select
                FontSize
                MyExcel.Range("A" & Format(RowFirst)).Select
                MyExcel.ActiveCell.FormulaR1C1 = fillhead
                MyExcel.Range(RowNumber & Format(RowFirst)).Select
                LeftSize
                FontSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.Range(RowNumber & Format(RowFirst)).Select
                MyExcel.ActiveCell.FormulaR1C1 = num
                
                Mystring = Asc(RowNumber)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.Range(Mystring & Format(RowFirst)).Select
                LeftSize
                FontSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.ActiveCell.FormulaR1C1 = er_max 'er_avg
                   
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.Range(Mystring & Format(RowFirst)).Select
                LeftSize
                FontSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.ActiveCell.FormulaR1C1 = er_avg
                   
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.Range(Mystring & Format(RowFirst)).Select
                LeftSize
                FontSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.ActiveCell.FormulaR1C1 = er_min 'er4
                
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.Range(Mystring & Format(RowFirst)).Select
                LeftSize
                FontSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.ActiveCell.FormulaR1C1 = er3 'er4
                
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.Range(Mystring & Format(RowFirst)).Select
                LeftSize
                FontSize
                MyExcel.Selection.Font.Bold = False
                MyExcel.ActiveCell.FormulaR1C1 = er4
                
                
                'Print #FileNumber, fillhead; Space$(fillnum); num; er_max; er_avg; er_min; er3; er4
            End If
         End If
      End If
errend:
End Sub



Sub Row_Col(ByVal gg As String, ByVal kk As Integer, ByVal RowFirst As Integer, ByVal RowNumber As String)
    Dim i As Integer
    Dim Mystring As String
    On Error Resume Next
    
    MyExcel.Rows(Format(RowFirst)).Select
    FontSize
    MyExcel.Range("A" & Format(RowFirst)).Select
    If kk = 9 Then
       MyExcel.Selection.Font.Bold = True
    Else
       MyExcel.Selection.Font.Bold = False
    End If
    
    MyExcel.ActiveCell.FormulaR1C1 = gg
    If table_s(kk, 1) = 0 Then GoTo tr
      
       '*********
       MyExcel.Range(RowNumber & Format(RowFirst)).Select
       FontSize
       LeftSize
    If kk = 9 Then
       MyExcel.Selection.Font.Bold = True
    Else
       MyExcel.Selection.Font.Bold = False
    End If
    MyExcel.Range(RowNumber & Format(RowFirst)).Select
    MyExcel.ActiveCell.FormulaR1C1 = table_f(kk, 1)
    Mystring = Format(RowNumber)
    If Not AssignmntFlag Then
    
    Mystring = Asc(Mystring)
    i = Val(Mystring) + 1
    Mystring = Chr(i)
    MyExcel.Range(Mystring & Format(RowFirst)).Select
    FontSize
    LeftSize
    If kk = 9 Then
       MyExcel.Selection.Font.Bold = True
    Else
       MyExcel.Selection.Font.Bold = False
    End If
    
    MyExcel.ActiveCell.FormulaR1C1 = table_f(kk, 2)
    
    Mystring = Asc(Mystring)
    i = Val(Mystring) + 1
    Mystring = Chr(i)
    MyExcel.Range(Mystring & Format(RowFirst)).Select
    FontSize
    LeftSize
    If kk = 9 Then
       MyExcel.Selection.Font.Bold = True
    Else
       MyExcel.Selection.Font.Bold = False
    End If
    
    MyExcel.ActiveCell.FormulaR1C1 = table_f(kk, 3)
    
    Mystring = Asc(Mystring)
    i = Val(Mystring) + 1
    Mystring = Chr(i)
    MyExcel.Range(Mystring & Format(RowFirst)).Select
    FontSize
    LeftSize
    If kk = 9 Then
       MyExcel.Selection.Font.Bold = True
       
    Else
       MyExcel.Selection.Font.Bold = False
    End If
    
    MyExcel.ActiveCell.FormulaR1C1 = table_f(kk, 4)
    End If
    If kk = 9 Then
       MyExcel.Rows(Format(RowFirst) & ":" & Format(RowFirst)).Select
       With MyExcel.Selection.Interior
           .ColorIndex = 15
           .Pattern = 1
       End With
       GoTo tr
    Else
        'MyExcel.Selection.Font.Bold = True
        Mystring = Asc(Mystring)
        i = Val(Mystring) + 1
        Mystring = Chr(i)
        MyExcel.Range(Mystring & Format(RowFirst)).Select
        FontSize
        LeftSize
        MyExcel.ActiveCell.FormulaR1C1 = table_f(kk, 5)
        
        Mystring = Asc(Mystring)
        i = Val(Mystring) + 1
        Mystring = Chr(i)
        MyExcel.Range(Mystring & Format(RowFirst)).Select
        FontSize
        LeftSize
        
        MyExcel.ActiveCell.FormulaR1C1 = table_f(kk, 6)
        
    End If
    
tr:
    MyExcel.Range("A" & Format(RowFirst) + 1).Select
End Sub
Function mess_num(ByVal tabna As String) As Integer
    Dim msg As String
    Dim MyTableName As String
    Dim MyDbName As String
    Dim j As Integer
    Dim menum
    Dim dbs As Database, rst As Recordset
    Dim lngRecords As Long, lngFields As Long
    Dim lngRecCount As Long, lngFldCount As Long
    On Error GoTo errend
    DoEvents
    menum = 0
    For j = 1 To stre_num
        MyTableName = convert_filename(j)
        MyDbName = ""
        Do While InStr(MyTableName, "\") > 0
            MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
            MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
        Loop
        MyDbName = Left(MyDbName, Len(MyDbName) - 1)
        Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & " count(*) " _
                    & "AS countmessage FROM " & MyTableName & " Where  message = """ & tabna & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
        menum = menum + rst.Fields(0).Value
        dbs.Close
    Next
    mess_num = menum
errend:
End Function
Function Mark_num(ByVal tabna As String) As Integer
    Dim msg As String
    Dim MyTableName As String
    Dim MyDbName As String
    Dim j As Integer
    Dim menum
    Dim dbs As Database, rst As Recordset
    Dim lngRecords As Long, lngFields As Long
    Dim lngRecCount As Long, lngFldCount As Long
    On Error GoTo errend
    menum = 0
    For j = 1 To stre_num
        MyTableName = convert_filename(j)
        MyDbName = ""
        Do While InStr(MyTableName, "\") > 0
            MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
            MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
        Loop
        MyDbName = Left(MyDbName, Len(MyDbName) - 1)
        Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & " count(*) " _
                    & "AS countmessage FROM " & MyTableName & " Where  mark = """ & tabna & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
        menum = menum + rst.Fields(0).Value
        dbs.Close
    Next
    Mark_num = menum
errend:
End Function

Function Gsm_Dcs(ByVal a As String, ByVal b As String) As Integer
    Dim msg As String
    Dim MyTableName As String
    Dim MyDbName As String
    Dim j As Integer
    Dim menum
    Dim dbs As Database, rst As Recordset
    Dim lngRecords As Long, lngFields As Long
    Dim lngRecCount As Long, lngFldCount As Long
    On Error GoTo errend
    DoEvents
    menum = 0
    For j = 1 To stre_num
        MyTableName = convert_filename(j)
        MyDbName = ""
        Do While InStr(MyTableName, "\") > 0
            MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
            MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
        Loop
        MyDbName = Left(MyDbName, Len(MyDbName) - 1)
        Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
       Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(*) " _
                    & "AS countmessage FROM " & MyTableName & " where BCCH_SERV >= " & Format(a) & " and  BCCH_SERV <" & Format(b))
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
            
        menum = menum + rst.Fields(0).Value
        dbs.Close
    Next
    
    Gsm_Dcs = menum
errend:
End Function


Sub pri_tabe(ByVal a As String, ByVal b As Integer)
    DoEvents
    On Error GoTo errend
    If table_s(b, 1) = 0 Then
       Print #FileNumber, a
       Exit Sub
    End If
    Dim s1 As Integer, s2 As Integer, s3 As Integer, s4 As Integer, s5 As Integer, s6 As Integer
    s1 = 17 - Len(a)
    s2 = 10 - Len(table_f(b, 1))
    s3 = 10 - Len(table_f(b, 2))
    s4 = 10 - Len(table_f(b, 3))
    s5 = 10 - Len(table_f(b, 4))
    s6 = 10 - Len(table_f(b, 5))
    Print #FileNumber, a; space$(s1); table_f(b, 1); space$(s2); table_f(b, 2); space$(s3); table_f(b, 3); space$(s4); table_f(b, 4); space$(s5); table_f(b, 5); space$(s6); table_f(b, 6)
errend:
End Sub

Sub hand_time(getnum1, getnum2, getnum3)
    Dim h_comm() As Single, h_comp() As Single, h_fail() As Single
    Dim tp(1 To 1000) As Single
    Dim menum, tt, all
    Dim dd As String, abc As String
    Dim hav As Boolean
    Dim ggg As Single, sss As Single
    Dim j As Integer, k As Integer, HH As Integer, i As Integer, m As Integer
    Dim pp As Integer, qq As Integer
    Dim mid_dd As Long
    'Dim abc As String
    Dim finds As Integer
    Dim ssi As Integer, bbb As Integer, tal As Integer, tber As Integer
    ReDim h_comm(1 To 1000) As Single
    ReDim h_comp(1 To 1000) As Single
    ReDim h_fail(1 To 1000) As Single
    Dim MyTableName As String
    Dim MyDbName As String
    Dim dbs As Database, rst As Recordset
    Dim lngRecords As Long, lngFields As Long
    Dim lngRecCount As Long, lngFldCount As Long
    On Error Resume Next
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
    
    For HH = 1 To stre_num
        MyTableName = convert_filename(HH)
        MyDbName = ""
        Do While InStr(MyTableName, "\") > 0
            MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
            MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
        Loop
        MyDbName = Left(MyDbName, Len(MyDbName) - 1)
        Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
        For pp = 1 To 1000
            h_comm(pp) = 0
            h_comp(pp) = 0
            h_fail(pp) = 0
        Next
    For qq = 1 To 3
        If qq = 1 Then abc = "HANDOVER COMMAND"
        If qq = 2 Then abc = "HANDOVER COMPLETE"
        If qq = 3 Then abc = "HANDOVER FAILURE"
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & "time " _
                    & "AS countmessage FROM " & MyTableName & " Where  message = """ & abc & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
            menum = lngRecords
        'mapinfo.Do "select * from " + Chr(34) + stre_tab(HH) + Chr(34) + " where col5 = " + Chr(34) + abc + Chr(34) + "into temp order by col1"
        'menum = mapinfo.eval("tableinfo(temp,8)")
        If qq = 1 Then getnum1 = getnum1 + menum
        If qq = 2 Then getnum2 = getnum2 + menum
        If qq = 3 Then getnum3 = getnum3 + menum
        
        If menum > 0 Then
           rst.MoveFirst
           Select Case abc
               Case "HANDOVER COMMAND":
                    For i = 1 To menum
                        dd = rst.Fields(0).Value
                        finds = InStr(dd, ":")
                        mid_dd = Val(Left(dd, finds - 1)) * 3600
                        h_comm(i) = mid_dd
                        dd = Right(dd, Len(dd) - finds)
                        finds = InStr(dd, ":")
                        If finds > 0 Then
                           mid_dd = Val(Left(dd, finds - 1)) * 60
                           h_comm(i) = h_comm(i) + mid_dd
                           dd = Right(dd, Len(dd) - finds)
                        End If
                        finds = InStr(dd, ".")
                        If finds > 0 Then
                           mid_dd = Val(Left(dd, finds - 1))
                           dd = Right(dd, Len(dd) - finds)
                        Else
                           mid_dd = Val(dd)
                           dd = 0
                        End If
                        h_comm(i) = h_comm(i) + mid_dd
                        h_comm(i) = h_comm(i) + Val(dd) / 100
                        If i < menum Then
                            rst.MoveNext
                        End If
                    Next
'                    h_comm(i) = -1
               Case "HANDOVER COMPLETE":
                    For i = 1 To menum
                        dd = rst.Fields(0).Value
                        finds = InStr(dd, ":")
                        mid_dd = Val(Left(dd, finds - 1)) * 3600
                        h_comp(i) = mid_dd
                        dd = Right(dd, Len(dd) - finds)
                        finds = InStr(dd, ":")
                        If finds > 0 Then
                           mid_dd = Val(Left(dd, finds - 1)) * 60
                           h_comp(i) = h_comp(i) + mid_dd
                           dd = Right(dd, Len(dd) - finds)
                        End If
                        finds = InStr(dd, ".")
                        If finds > 0 Then
                           mid_dd = Val(Left(dd, finds - 1))
                           dd = Right(dd, Len(dd) - finds)
                        Else
                           mid_dd = Val(dd)
                           dd = 0
                        End If
                        h_comp(i) = h_comp(i) + mid_dd
                        
                        h_comp(i) = h_comp(i) + Val(dd) / 100
                        If i < menum Then
                           rst.MoveNext
                        End If
                    Next
 '                   h_comp(i) = -1
               Case "HANDOVER FAILURE":
                     For i = 1 To menum
                         dd = rst.Fields(0).Value
                         finds = InStr(dd, ":")
                         mid_dd = Val(Left(dd, finds - 1)) * 3600
                         h_fail(i) = mid_dd
                         dd = Right(dd, Len(dd) - finds)
                         finds = InStr(dd, ":")
                         If finds > 0 Then
                            mid_dd = Val(Left(dd, finds - 1)) * 60
                            h_fail(i) = h_fail(i) + mid_dd
                            dd = Right(dd, Len(dd) - finds)
                         End If
                         finds = InStr(dd, ".")
                         If finds > 0 Then
                            mid_dd = Val(Left(dd, finds - 1))
                            dd = Right(dd, Len(dd) - finds)
                         Else
                            mid_dd = Val(dd)
                            dd = 0
                         End If
                         h_fail(i) = h_fail(i) + mid_dd
                         
                         h_fail(i) = h_fail(i) + Val(dd) / 100
                         If i < menum Then
                            rst.MoveNext
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
       For m = j To getnum3
           tp(k) = h_fail(m)
           k = k + 1
           m = m + 1
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
    
    Next HH
errend:
End Sub


Sub hand_zz()
   Dim i As Integer, j As Integer
   Dim tawri As Boolean
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

Sub xt_time(ByVal hu As Boolean, assnum, ByVal hg As Boolean)
    Dim h_comm() As Single, h_comp() As Single
    Dim menum, tt, all
    Dim dd As String
    Dim hav As Boolean
    Dim ggg As Single
    Dim mid_dd As Long
    Dim i As Integer, j As Integer, assi As Integer, assj As Integer
    Dim k As Integer, HH As Integer, finds As Integer, tal As Integer, pp As Integer
    Dim MyTableName As String
    Dim MyDbName As String
    Dim dbs As Database, rst As Recordset
    Dim lngRecords As Long, lngFields As Long
    Dim lngRecCount As Long, lngFldCount As Long
    ReDim h_comm(1 To 1000) As Single
    ReDim h_comp(1 To 1000) As Single
    On Error GoTo errend
    
    'DoEvents
    menum = 0
    assnum = 0
    For j = 1 To 9
        For k = 1 To 6
            table_s(j, k) = 0
        Next k
    Next j
    For HH = 1 To stre_num
        For pp = 1 To 1000
            h_comm(pp) = 0
            h_comp(pp) = 0
        Next
        
        MyTableName = convert_filename(HH)
        MyDbName = ""
        Do While InStr(MyTableName, "\") > 0
            MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
            MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
        Loop
        MyDbName = Left(MyDbName, Len(MyDbName) - 1)
        Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
        
        DoEvents
        If hu = False Then
            If hg = True Then
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "CONNECT ACKNOWLEDGE" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
            Else
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "CHANNEL REQUEST" & """ or message = """ & "CHANNEL REQUEST REPORT" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
             ' mapinfo.Do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + "CHANNEL REQUEST" + Chr(34) + "into temp order by col1"
            End If
        Else
           Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "SETUP" & """ or message = """ & "EMERGENCY SETUP" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
        End If
                menum = lngRecords
                DoEvents
                If menum > 0 Then
                    rst.MoveFirst
                    For i = 1 To menum
                        dd = rst.Fields(0).Value
                        finds = InStr(dd, ":")
                        mid_dd = Val(Left(dd, finds - 1)) * 3600
                        h_comm(i) = mid_dd
                        dd = Right(dd, Len(dd) - finds)
                        finds = InStr(dd, ":")
                        If finds > 0 Then
                            mid_dd = Val(Left(dd, finds - 1)) * 60
                            h_comm(i) = h_comm(i) + mid_dd
                            dd = Right(dd, Len(dd) - finds)
                        End If
                        finds = InStr(dd, ".")
                        If finds > 0 Then
                            mid_dd = Val(Left(dd, finds - 1))
                            dd = Right(dd, Len(dd) - finds)
                        Else
                            mid_dd = Val(dd)
                            dd = 0
                        End If
                        h_comm(i) = h_comm(i) + mid_dd
                        h_comm(i) = h_comm(i) + Val(dd) / 100
                        If i < menum Then
                            rst.MoveNext
                        End If
                    Next
                    h_comm(i) = -1
                End If
    '***********
      If hg = False Then
        If hu = False Then
          Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "ASSIGNMENT COMMAND" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
          ' mapinfo.Do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + "ASSIGNMENT COMMAND" + Chr(34) + "into temp order by col1"
        Else
          Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "ASSIGNMENT COMPLETE" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
                DoEvents
        End If
      Else
         If hu = False Then
           Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "DISCONNECT" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
               
                DoEvents
         Else
           Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "CONNECT" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
                DoEvents
         End If
      End If
      menum = lngRecords
      DoEvents
      If menum > 0 Then
       '    mapinfo.Do "fetch first from temp"
           rst.MoveFirst
           For i = 1 To menum
               dd = rst.Fields(0).Value
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 3600
               h_comp(i) = mid_dd
               dd = Right(dd, Len(dd) - finds)
               finds = InStr(dd, ":")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1)) * 60
                    h_comp(i) = h_comp(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
               End If
               finds = InStr(dd, ".")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1))
                    h_comp(i) = h_comp(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
                    h_comp(i) = h_comp(i) + Val(dd) / 100
               Else
                    h_comp(i) = h_comp(i) + Val(dd)
                    dd = 0
               End If
               If i < menum Then
                    rst.MoveNext
               End If
           Next
           dbs.Close
           h_comp(i) = -1
        End If
        
        If hu = False And hg = False Then GoTo efid
'        If hu = True Then
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
           GoTo next1
          ' Exit Sub
'        End If
efid:
           i = 1
           j = 1
           tal = 0
           Do While h_comm(i) <> -1 And h_comm(i) <> 0
              hav = False
              If h_comm(i) <= h_comp(j) Then
                 ggg = h_comp(j) - h_comm(i)
                 j = j + 1
                 i = i + 1
                 hav = True
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
next1:
    
    Next HH
errend:
End Sub

Sub xtt_time(ByVal hu As Boolean, assnum, ByVal hg As Boolean)
    Dim h_comm() As Single, h_comp() As Single
    Dim menum, tt, all
    Dim dd As String
    Dim hav As Boolean
    Dim ggg As Single
    Dim mid_dd As Long
    Dim i As Integer, j As Integer, assi As Integer, assj As Integer
    Dim k As Integer, HH As Integer, finds As Integer, tal As Integer, pp As Integer
    Dim MyTableName As String
    Dim MyDbName As String
    Dim dbs As Database, rst As Recordset
    Dim lngRecords As Long, lngFields As Long
    Dim lngRecCount As Long, lngFldCount As Long
    ReDim h_comm(1 To 1000) As Single
    ReDim h_comp(1 To 1000) As Single
    On Error GoTo errend
    
    'DoEvents
    menum = 0
    assnum = 0
    For j = 1 To 9
        For k = 1 To 6
            table_s(j, k) = 0
        Next k
    Next j
    For HH = 1 To stre_num
        For pp = 1 To 1000
            h_comm(pp) = 0
            h_comp(pp) = 0
        Next
        
        MyTableName = convert_filename(HH)
        MyDbName = ""
        Do While InStr(MyTableName, "\") > 0
            MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
            MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
        Loop
        MyDbName = Left(MyDbName, Len(MyDbName) - 1)
        Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
        
        DoEvents
        If hu = False Then
            If hg = False Then
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "ASSIGNMENT COMMAND" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
            End If
        End If
                menum = lngRecords
                DoEvents
                If menum > 0 Then
                    rst.MoveFirst
                    For i = 1 To menum
                        dd = rst.Fields(0).Value
                        finds = InStr(dd, ":")
                        mid_dd = Val(Left(dd, finds - 1)) * 3600
                        h_comm(i) = mid_dd
                        dd = Right(dd, Len(dd) - finds)
                        finds = InStr(dd, ":")
                        If finds > 0 Then
                            mid_dd = Val(Left(dd, finds - 1)) * 60
                            h_comm(i) = h_comm(i) + mid_dd
                            dd = Right(dd, Len(dd) - finds)
                        End If
                        finds = InStr(dd, ".")
                        If finds > 0 Then
                            mid_dd = Val(Left(dd, finds - 1))
                            dd = Right(dd, Len(dd) - finds)
                        Else
                            mid_dd = Val(dd)
                            dd = 0
                        End If
                        h_comm(i) = h_comm(i) + mid_dd
                        h_comm(i) = h_comm(i) + Val(dd) / 100
                        If i < menum Then
                            rst.MoveNext
                        End If
                    Next
                    h_comm(i) = -1
                End If
    '***********
      If hg = False Then
        If hu = False Then
          Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "DISCONNECT" & """")
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                lngRecords = rst.RecordCount
                lngFields = rst.Fields.Count
        End If
      End If
           menum = lngRecords
           If menum > 0 Then
           rst.MoveFirst
           For i = 1 To menum
               dd = rst.Fields(0).Value
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 3600
               h_comp(i) = mid_dd
               dd = Right(dd, Len(dd) - finds)
               finds = InStr(dd, ":")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1)) * 60
                    h_comp(i) = h_comp(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
               End If
               finds = InStr(dd, ".")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1))
                    h_comp(i) = h_comp(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
                    h_comp(i) = h_comp(i) + Val(dd) / 100
               Else
                    h_comp(i) = h_comp(i) + Val(dd)
                    dd = 0
               End If
               If i < menum Then
                    rst.MoveNext
               End If
           Next
           dbs.Close
           h_comp(i) = -1
        End If
        
           i = 1
           j = 1
           tal = 0
           Do While h_comm(i) <> -1 And h_comm(i) <> 0
              hav = False
              If h_comm(i) <= h_comp(j) Then
                 ggg = h_comp(j) - h_comm(i)
                 j = j + 1
                 i = i + 1
                 hav = True
              Else
                 j = j + 1
                 If h_comp(j) = -1 Or h_comp(j) = 0 Then Exit Do
              End If
              If hav = True Then
                 Select Case ggg
                     Case Is >= 4800
                          tal = 8
                     Case Is >= 4200
                          tal = 7
                     Case Is >= 3600
                          tal = 6
                     Case Is >= 300
                          tal = 5
                     Case Is >= 180
                          tal = 4
                     Case Is >= 120
                          tal = 3
                     Case Is >= 60
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
next1:
    
    Next HH
errend:
End Sub


