Attribute VB_Name = "Module1"
Option Explicit

'Public convert_filename(1 To 50) As String '
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
Dim HoASFLFlag As Boolean
Dim CFlag As Boolean
Dim SumNum As Double
Dim dbs As Database
Dim MyFieldType As Byte
Dim rst As Recordset

Sub TEST_REPORT()
    Dim MyC1C1Newflag As Boolean
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
    Dim i As Integer, j As Integer, k As Integer
    Dim MyTempPath As String
    Dim all_0 As String
    Dim MyTableName As String, MyDbName As String
    Dim RowNum As Integer
    Dim MyRecordset As Recordset, MynewRst As Recordset
    
    'Dim MyFirstCall As Integer, MyLastCall As Integer
    Dim MyRowCount As Integer, MyRowTemp As Integer
    Dim MystrTemp As String
    Dim MystrTime As Date
    Dim CauseValue() As Integer
    Dim CVString(0 To 17) As String
    Dim MyQueryTemp As QueryDef
    Dim CallNumber() As Integer
    Dim CallType() As Integer
    Dim CallTime() As Date
    Dim TotalCall As Integer
    Dim CurrentRow As Integer
    
    On Error Resume Next
    MyFieldType = 0
    Frmrepot.Show
    Screen.MousePointer = 11 '����ͳ��
    'DoEvents
    Frmrepot.ProgressBar1.Max = 100
    Frmrepot.ProgressBar1.Value = 1
    CFlag = False
    AssignmntFlag = False
    MyTempPath = Gsm_Path + "\user\"
    If Dir(MyTempPath, 16) <> "" Then
       ChDir MyTempPath
    Else
       MkDir MyTempPath
    End If
    stcname = Gsm_Path + "\user\" + stre_tab(1) + ".xls"
    'cc_all = -2
    cc_all = 0
    
    Frmrepot.Label1.Caption = "���ڴ������ļ� ..."
    Frmrepot.Label1.Refresh
    
        MyTableName = convert_filename(1)
        MyDbName = ""
        Do While InStr(MyTableName, "\") > 0
            MyDbName = MyDbName & Left(MyTableName, InStr(MyTableName, "\"))
            MyTableName = Right(MyTableName, Len(MyTableName) - InStr(MyTableName, "\"))
        Loop
        MyDbName = Left(MyDbName, Len(MyDbName) - 1)
        Set dbs = OpenDatabase(MyDbName, False, False, "Foxpro 3.0;")
    
    For i = 1 To stre_num
        MyTableName = convert_filename(i)
        Set rst = dbs.OpenRecordset("SELECT  " _
        & " count(*) as countrxlev FROM " & MyTableName)
        If rst.RecordCount <> 0 Then
            rst.MoveLast
        End If
        cc_all = cc_all + rst.Fields(0).Value
    
    Next
    '***************
    'Screen.MousePointer = 0
    
    Frmrepot.Label1.Caption = "�������� Excel ..."
    Frmrepot.Label1.Refresh
    Set MyExcel = CreateObject("excel.application")
    MyExcel.Visible = True
    MyExcel.Workbooks.ADD
    MyExcel.Application.DisplayAlerts = False
    'MyExcel.Sheets("Sheet1").Select
    MyExcel.Columns("A:A").ColumnWidth = 29
    MyExcel.Columns("B:B").ColumnWidth = 12
    MyExcel.Sheets("Sheet1").Name = "����ͳ��"
    Frmrepot.ProgressBar1.Value = 3
    Frmrepot.Label1.Caption = "���ڽ��к���ͳ�� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells.Font.Size = 9
    MyExcel.cells.Font.ColorIndex = 0
    MyExcel.cells.HorizontalAlignment = -4131
    
    MyExcel.cells(2, 1).Font.ColorIndex = 11
    MyExcel.cells(2, 1).Value = "...��������:MOC(MS����)/MTC(MS����)..."
    '**********
    
'        MyExcel.cells(1, 1).Value = Timer
'        MyExcel.cells(1, 4).Value = time
    
    CallAttemp '����ͳ�Ʊ���         'Lee
    
 '       MyExcel.cells(1, 2).Value = Timer
 '       MyExcel.cells(1, 3).Value = Format(CLng(MyExcel.cells(1, 2).Value) - CLng(MyExcel.cells(1, 1).Value))
 '       MyExcel.cells(1, 5).Value = time
    
    Frmrepot.ProgressBar1.Value = 25

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&ANT����ͳ���嵥
    
    Frmrepot.Label1.Caption = "�����г�����ͳ����ϸ�嵥 ..."
    Frmrepot.Label1.Refresh
    MyExcel.Sheets("Sheet2").Select
    MyExcel.Sheets("Sheet2").Name = "����ͳ���嵥"
    MyExcel.Columns("A:A").ColumnWidth = 6.25
    MyExcel.Columns("B:B").ColumnWidth = 6.25
    MyExcel.Columns("C:C").ColumnWidth = 6.13
    MyExcel.Columns("D:D").ColumnWidth = 7.75
    MyExcel.Columns("E:E").ColumnWidth = 9
    MyExcel.Columns("F:F").ColumnWidth = 12.63
    MyExcel.Columns("G:G").ColumnWidth = 6.88
    MyExcel.Columns("H:H").ColumnWidth = 6.88
    MyExcel.Columns("I:I").ColumnWidth = 11.25
    MyExcel.Columns("J:J").ColumnWidth = 10.38
    MyExcel.Columns("K:K").ColumnWidth = 11.63
    MyExcel.cells.Font.Size = 9
    MyExcel.cells.Font.ColorIndex = 0
    MyExcel.cells.HorizontalAlignment = -4131
    MyExcel.Columns("B:D").HorizontalAlignment = -4108
    MyExcel.Columns("G:H").HorizontalAlignment = -4108
    MyExcel.Rows("3:4").Font.Bold = False
    MyExcel.Rows("3:4").Font.ColorIndex = 11
    MyExcel.Rows("3:4").HorizontalAlignment = -4108
    MyExcel.cells(3, 1).Value = "�ļ���"
    MyExcel.cells(3, 2).Value = "ͨ�����"
    MyExcel.cells(3, 3).Value = "��������"
    MyExcel.cells(3, 4).Value = "����"
    MyExcel.cells(3, 5).Value = "����С��"
    MyExcel.cells(3, 6).Value = "����״̬"
    MyExcel.cells(3, 7).Value = "��������"
    MyExcel.cells(3, 8).Value = "��������"
    MyExcel.cells(3, 9).Value = "�л��ɹ�"
    MyExcel.cells(3, 10).Value = "�л�ʧ��"
    MyExcel.cells(3, 11).Value = "ͨ���ȼ�"
    MyExcel.cells(4, 5).Value = "(CI)"
    MyExcel.cells(4, 7).Value = "(����)"
    MyExcel.cells(4, 8).Value = "(����)"
    MyExcel.cells(4, 9).Value = "(����+����)"
    MyExcel.cells(4, 10).Value = "(����+ԭ��)"
    MyExcel.Columns("E:E").NumberFormatLocal = "@"
    MyExcel.Columns(4).NumberFormatLocal = "@"
    CVString(0) = "�����ͷ�"
    CVString(1) = "�쳣�ͷţ�δָ��"
    CVString(2) = "�쳣�ͷţ��ŵ�������"
    CVString(3) = "�쳣�ͷţ���ʱ"
    CVString(4) = "����·���޻"
    CVString(5) = "Ԥռ�ͷ�"
    CVString(6) = "�����л���ʱ�䳬��"
    CVString(7) = "���ɽ��ܵ�ͨ��ģʽ"
    CVString(8) = "δ�ṩ��Ƶ��"
    CVString(9) = "���������"
    CVString(10) = "���������"
    CVString(11) = "��Ч��Ϣ��δ�涨"
    CVString(12) = "��Ϣ���Ͳ����ڣ�����ʵ��"
    CVString(13) = "��Ϣ���������״̬�����ݻ򲻴��ڣ�����ʵ��"
    CVString(14) = "��Ч��Ϣ��Ԫ����"
    CVString(15) = "��С������"
    CVString(16) = "δ�涨��Э����"
    CVString(17) = "����ԭ��"
    MyRowCount = 0
    
        'MyExcel.cells(1, 1).Value = Timer
        'MyExcel.cells(1, 4).Value = time
    Frmrepot.ProgressBar1.Value = 28
    MyRowTemp = 0
    For i = 1 To stre_num
        
        Frmrepot.Label2.Caption = convert_filename(i)
    Frmrepot.Label2.Refresh
        MyTableName = convert_filename(i)
        
        'MystrTemp = MyTableName
        'Do While InStr(MystrTemp, "\") > 0
        '    MystrTemp = Right(MystrTemp, Len(MystrTemp) - InStr(MystrTemp, "\"))
        'Loop
        'MyExcel.cells(5 + MyRowCount, 1).Value = MystrTemp
        '******* Set rst = dbs.OpenRecordset("SELECT rxle_same2 FROM " & MyTableName & " where rxle_same2>0 group by rxle_same2 order by rxle_same2 ASC ")
        Set rst = dbs.OpenRecordset("SELECT rxle_same2,bsic_same2 FROM " & MyTableName & " where rxle_same2>0 group by rxle_same2,bsic_same2")
        If rst.RecordCount <> 0 Then
            rst.MoveLast
            rst.MoveFirst
            TotalCall = rst.RecordCount
            ReDim CallNumber(TotalCall) As Integer
            ReDim CallType(TotalCall) As Integer
            ReDim CallTime(TotalCall) As Date
            For j = 1 To TotalCall
                CallNumber(j) = rst.Fields("rxle_same2").Value
                CallType(j) = rst.Fields("bsic_same2").Value
                rst.MoveNext
            Next
            'MyFirstCall = rst.Fields(0).Value
            'rst.MoveLast
            'MyLastCall = rst.Fields(0).Value
            Frmrepot.Label2.Caption = convert_filename(i) & "(�ܹ�" & Format(TotalCall) & "������)"
            Frmrepot.Label2.Refresh
            For j = 1 To TotalCall
                Frmrepot.Label1.Caption = "�����г���" & Format(j) & "��������ϸ�嵥 ..."
                Frmrepot.Label1.Refresh
                '******* Set MynewRst = dbs.OpenRecordset("select ci_serv FROM " & MyTableName & " where ci_serv<>"""" and rxle_same2=" & Format(j) & " group by ci_serv")
                Set MynewRst = dbs.OpenRecordset("select ci_serv FROM " & MyTableName & " where ci_serv<>"""" and rxle_same2=" & Format(CallNumber(j)) & " and bsic_same2=" & Format(CallType(j)) & " group by ci_serv")
                If MynewRst.RecordCount = 0 Then
                    MystrTemp = ""
                Else
                    MynewRst.MoveLast
                    MynewRst.MoveFirst
                    MystrTemp = ""
                    For k = 1 To MynewRst.RecordCount
                        MystrTemp = MystrTemp & MynewRst.Fields(0).Value & ","
                        MynewRst.MoveNext
                    Next
                    If Right(MystrTemp, 1) = "," Then
                        MystrTemp = Left(MystrTemp, Len(MystrTemp) - 1)
                    End If
                End If
                
                '******* Set rst = dbs.OpenRecordset("select * FROM " & MyTableName & " where rxle_same2=" & Format(j))
                Set rst = dbs.OpenRecordset("select * FROM " & MyTableName & " where rxle_same2=" & Format(CallNumber(j)) & " and bsic_same2=" & Format(CallType(j)))
                rst.MoveLast
                If InStr(rst.Fields("time").Value, ".") > 0 Then
                    MystrTime = Left(rst.Fields("time").Value, InStr(rst.Fields("time").Value, ".") - 1)
                Else
                    MystrTime = rst.Fields("time").Value
                End If
                For k = j To 1 Step -1
                    If MystrTime > CallTime(k - 1) Or k = 1 Then
                        CallTime(k) = MystrTime
                        CurrentRow = k
                        Exit For
                    Else
                        CallTime(k) = CallTime(k - 1)
                        'MyExcel.cells(4 + k, 2).Value = Format(k)

                    End If
                Next
                
                MyExcel.Rows(5 + MyRowTemp + CurrentRow - 1).EntireRow.Insert
                MyRowCount = MyRowTemp + CurrentRow - 1
                
                'MyExcel.cells(5 + MyRowCount, 2).Value = Format(j)
                MyExcel.cells(5 + MyRowCount, 5).Value = MystrTemp
                
                'Set MynewRst = dbs.OpenRecordset("select bsic_same2 FROM " & MyTableName & " where rxle_same2=" & Format(j))
                If rst.Fields("bsic_same2").Value = 2 Then
                    MystrTemp = "MTC"
                Else
                    MystrTemp = "MOC"
                End If
                MyExcel.cells(5 + MyRowCount, 3).Value = MystrTemp
                
                rst.Filter = "left(mark2,3)=""tel"""
                Set MynewRst = rst.OpenRecordset
                'Set MynewRst = dbs.OpenRecordset("select mark2 FROM " & MyTableName & " where left(mark2,3)=""tel"" and rxle_same2=" & Format(j))
                If MynewRst.RecordCount > 0 Then
                    MystrTemp = MynewRst.Fields("mark2").Value
                    MystrTemp = Right(MystrTemp, Len(MystrTemp) - 4)
                    MyExcel.cells(5 + MyRowCount, 4).Value = MystrTemp
                End If
                
                rst.Filter = "left(mark1,2)=""CS"""
                Set MynewRst = rst.OpenRecordset
                'Set MynewRst = dbs.OpenRecordset("select mark1 FROM " & MyTableName & " where left(mark1,2)=""CF"" and rxle_same2=" & Format(j))
                If MynewRst.RecordCount > 0 Then
                    MystrTemp = "�ɹ�����"
                    MyExcel.cells(5 + MyRowCount, 6).Value = MystrTemp
                Else
                    rst.Filter = "left(mark1,2)=""CF"""
                    Set MynewRst = rst.OpenRecordset
                    If MynewRst.RecordCount > 0 Then
                        MystrTemp = MynewRst.Fields("MARK1").Value
                        MystrTemp = Left(MystrTemp, InStr(MystrTemp, ",") - 1)
                        MystrTemp = Right(MystrTemp, Len(MystrTemp) - 3)
                        MystrTemp = "����ʧ�ܡ�" & MystrTemp
                        MyExcel.cells(5 + MyRowCount, 6).Value = MystrTemp
                        MyExcel.Range("B" & Format(5 + MyRowCount) & ":K" & Format(5 + MyRowCount)).Select
                        MyExcel.Selection.Interior.ColorIndex = 40
                        GoTo NextCall
                    Else   '�������
                        MystrTemp = "�ɹ�����"
                        MyExcel.cells(5 + MyRowCount, 6).Value = MystrTemp
                    End If
                End If

                'Set MynewRst = dbs.OpenRecordset("select count(*) FROM " & MyTableName & " where left(mark1,5)=""CP UL"" and rxle_same2=" & Format(j))
                'MynewRst.MoveFirst
                rst.Filter = "left(mark1,5)=""CP UL"""
                Set MynewRst = rst.OpenRecordset
                If MynewRst.RecordCount > 0 Then MynewRst.MoveLast
                MyExcel.cells(5 + MyRowCount, 7).Value = Format(MynewRst.RecordCount)
                
                'Set MynewRst = dbs.OpenRecordset("select count(*) FROM " & MyTableName & " where left(mark1,5)=""CP DL"" and rxle_same2=" & Format(j))
                rst.Filter = "left(mark1,5)=""CP DL"""
                Set MynewRst = rst.OpenRecordset
                If MynewRst.RecordCount > 0 Then MynewRst.MoveLast
                MyExcel.cells(5 + MyRowCount, 8).Value = Format(MynewRst.RecordCount)
                
                'Set MynewRst = dbs.OpenRecordset("select mark1 FROM " & MyTableName & " where left(mark1,3)=""HOS"" and rxle_same2=" & Format(j))
                rst.Filter = "left(mark1,3)=""HOS"""
                Set MynewRst = rst.OpenRecordset
                If MynewRst.RecordCount = 0 Then
                    MyExcel.cells(5 + MyRowCount, 9).Value = "0"
                Else
                    MynewRst.MoveLast
                    tmp_1 = 0
                    tmp_2 = 0
                    tmp_3 = 0
                    MynewRst.MoveFirst
                    For k = 1 To MynewRst.RecordCount
                        msg1 = MynewRst.Fields("mark1").Value
                        msg1 = Right(msg1, Len(msg1) - InStr(msg1, ","))
                        If InStr(msg1, ",") > 0 Then
                            msg1 = Right(msg1, Len(msg1) - InStr(msg1, ","))
                            Select Case Val(msg1)
                                Case 1
                                    tmp_1 = tmp_1 + 1
                                Case 2
                                    tmp_2 = tmp_2 + 1
                                Case 3
                                    tmp_3 = tmp_3 + 1
                            End Select
                        Else
                            tmp_2 = tmp_2 + 1
                        End If
                        MynewRst.MoveNext
                    Next
                    If tmp_1 > 0 Then
                        MystrTemp = Format(tmp_1) & ":ʱ϶�л�;  "
                    End If
                    If tmp_2 > 0 Then
                        MystrTemp = Format(tmp_2) & ":С���л�;  "
                    End If
                    If tmp_3 > 0 Then
                        MystrTemp = Format(tmp_3) & ":ϵͳ�л�"
                    End If
                    If Right(MystrTemp, 3) = ";  " Then
                        MystrTemp = Left(MystrTemp, Len(MystrTemp) - 3)
                    End If
                    MyExcel.cells(5 + MyRowCount, 9).Value = MystrTemp
                End If
                
                'Set MynewRst = dbs.OpenRecordset("select mark1 FROM " & MyTableName & " where left(mark1,3)=""HOF"" and rxle_same2=" & Format(j))
                rst.Filter = "left(mark1,3)=""HOF"""
                Set MynewRst = rst.OpenRecordset
                If MynewRst.RecordCount = 0 Then
                    MyExcel.cells(5 + MyRowCount, 10).Value = "0"
                Else
                    MynewRst.MoveLast
                    ReDim CauseValue(0 To 17) As Integer
                    MynewRst.MoveFirst
                    For k = 0 To MynewRst.RecordCount - 1
                        msg1 = MynewRst.Fields("mark1").Value
                        msg1 = Right(msg1, Len(msg1) - InStr(msg1, ","))
                        msg1 = Left(msg1, InStr(msg1, ",") - 1)
                        Select Case Val(msg1)
                            Case 0
                                CauseValue(0) = CauseValue(0) + 1
                            Case 1
                                CauseValue(1) = CauseValue(1) + 1
                            Case 2
                                CauseValue(2) = CauseValue(2) + 1
                            Case 3
                                CauseValue(3) = CauseValue(3) + 1
                            Case 4
                                CauseValue(4) = CauseValue(4) + 1
                            Case 5
                                CauseValue(5) = CauseValue(5) + 1
                            Case 8
                                CauseValue(6) = CauseValue(6) + 1
                            Case 9
                                CauseValue(7) = CauseValue(7) + 1
                            Case 10
                                CauseValue(8) = CauseValue(8) + 1
                            Case 65
                                CauseValue(9) = CauseValue(9) + 1
                            Case 95
                                CauseValue(10) = CauseValue(10) + 1
                            Case 96
                                CauseValue(11) = CauseValue(11) + 1
                            Case 97
                                CauseValue(12) = CauseValue(12) + 1
                            Case 98
                                CauseValue(13) = CauseValue(13) + 1
                            Case 100
                                CauseValue(14) = CauseValue(14) + 1
                            Case 101
                                CauseValue(15) = CauseValue(15) + 1
                            Case 111
                                CauseValue(16) = CauseValue(16) + 1
                            Case Else
                                CauseValue(17) = CauseValue(17) + 1
                        End Select
                        MynewRst.MoveNext
                    Next
                    MystrTemp = ""
                    For k = 0 To 17
                        If CauseValue(k) <> 0 Then
                            If Len(MystrTemp) > 0 Then MystrTemp = MystrTemp + ";  "
                            MystrTemp = MystrTemp + Format(CauseValue(k)) + ":" + CVString(k)
                        End If
                    Next
                    MyExcel.cells(5 + MyRowCount, 10).Value = MystrTemp
                End If
                'Set MynewRst = dbs.OpenRecordset("select mark1 FROM " & MyTableName & " where instr(mark1,""����"")>0 and rxle_same2=" & Format(j))
                rst.Filter = "instr(mark1,""����"")>0"
                Set MynewRst = rst.OpenRecordset
                If MynewRst.RecordCount > 0 Then
                    MystrTemp = MynewRst.Fields("mark1").Value
                    MystrTemp = Left(MystrTemp, InStr(MystrTemp, ",") - 1)
                    MystrTemp = Right(MystrTemp, Len(MystrTemp) - 3) & " & "
                Else
                    MystrTemp = ""
                End If
                'Set rst = dbs.OpenRecordset("select count(*) FROM " & MyTableName & " where rxle_same2=" & Format(j))
                'rst.MoveFirst
                
                'Set MynewRst = dbs.OpenRecordset("select mark FROM " & MyTableName & " where mark=""Noisy Call"" and rxle_same2=" & Format(j))
                rst.Filter = "mark=""Noisy Call"""
                Set MynewRst = rst.OpenRecordset
                If MynewRst.RecordCount > 0 Then
                    MystrTemp = MystrTemp & "����ͨ��" & " & "
                Else
                    'Set MynewRst = dbs.OpenRecordset("select count(*) FROM " & MyTableName & " where int(rxqual_s)>=4 and rxle_same2=" & Format(j))
                    rst.Filter = "int(rxqual_s)>=4"
                    Set MynewRst = rst.OpenRecordset
                    If MynewRst.RecordCount > 0 Then MynewRst.MoveLast
                    'MynewRst.MoveFirst
                    'tmp_1 = MynewRst.Fields(0).Value / rst.Fields(0).Value
                    tmp_1 = MynewRst.RecordCount / rst.RecordCount
                    If tmp_1 > 0.05 Then
                        MystrTemp = MystrTemp & "����ͨ��" & " & "
                    End If
                End If
                'Set MynewRst = dbs.OpenRecordset("select count(*) FROM " & MyTableName & " where rxlev_s<=20 and rxle_same2=" & Format(j))
                rst.Filter = "rxlev_s<=20"
                Set MynewRst = rst.OpenRecordset
                If MynewRst.RecordCount > 0 Then MynewRst.MoveLast
                tmp_1 = MynewRst.RecordCount / rst.RecordCount
                If tmp_1 > 0.05 Then
                    MystrTemp = MystrTemp & "���ź�ͨ��" & " & "
                End If
                If MystrTemp <> "" Then
                    MystrTemp = Left(MystrTemp, Len(MystrTemp) - 3)
                    MyExcel.Range("B" & Format(5 + MyRowCount) & ":K" & Format(5 + MyRowCount)).Select
                    MyExcel.Selection.Interior.ColorIndex = 40
                Else
                    MyExcel.Range("B" & Format(5 + MyRowCount) & ":K" & Format(5 + MyRowCount)).Select
                    MyExcel.Selection.Interior.ColorIndex = 35
                    MystrTemp = "����ͨ��"
                End If
                MyExcel.cells(5 + MyRowCount, 11).Value = MystrTemp
NextCall:
                MyRowCount = MyRowCount + 1
            Next
            For j = 1 To TotalCall
                MyExcel.cells(4 + MyRowTemp + j, 2).Value = Format(j)
            Next
            MystrTemp = MyTableName
            Do While InStr(MystrTemp, "\") > 0
                MystrTemp = Right(MystrTemp, Len(MystrTemp) - InStr(MystrTemp, "\"))
            Loop
            MyExcel.cells(5 + MyRowTemp, 1).Value = MystrTemp
            MyRowTemp = MyRowTemp + TotalCall
        End If
    Next
    Frmrepot.ProgressBar1.Value = 40
    Frmrepot.Label2.Caption = ""
Frmrepot.Label2.Refresh

        'MyExcel.cells(1, 2).Value = Timer
        'MyExcel.cells(1, 3).Value = Format(CLng(MyExcel.cells(1, 2).Value) - CLng(MyExcel.cells(1, 1).Value))
        'MyExcel.cells(1, 5).Value = time

    'MyExcel.Range("A5:K" & Format(5 + MyRowCount - 1)).Select
    MyExcel.Range("A5:K" & Format(5 + MyRowTemp - 1)).Select
    MyExcel.Selection.Borders(5).LineStyle = -4142
    MyExcel.Selection.Borders(6).LineStyle = -4142
    With MyExcel.Selection.Borders(7)
        .LineStyle = 1
        .Weight = 2
        .ColorIndex = -4105
    End With
    With MyExcel.Selection.Borders(8)
        .LineStyle = 1
        .Weight = 2
        .ColorIndex = -4105
    End With
    With MyExcel.Selection.Borders(9)
        .LineStyle = 1
        .Weight = 2
        .ColorIndex = -4105
    End With
    With MyExcel.Selection.Borders(10)
        .LineStyle = 1
        .Weight = 2
        .ColorIndex = -4105
    End With
    With MyExcel.Selection.Borders(11)
        .LineStyle = 1
        .Weight = 2
        .ColorIndex = -4105
    End With
    With MyExcel.Selection.Borders(12)
        .LineStyle = 1
        .Weight = 2
        .ColorIndex = -4105
    End With

'    Exit Sub
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

    '****************************************'ANT����ͳ�Ʊ���
MyTest:
    Frmrepot.ProgressBar1.Value = 45
    Frmrepot.Label1.Caption = "���ڽ��в���ͳ�� ..."
    DoEvents
    Frmrepot.Label1.Caption = "���ڽ���ϵͳ��Ӧʱ��ͳ�Ʊ�ͳ�� ..."
Frmrepot.Label1.Refresh
    Call xt_time(False, tt, False)
    hand_zz
    Frmrepot.ProgressBar1.Value = 48
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
    
    
    MyExcel.Sheets("Sheet3").Select
    MyExcel.ActiveWindow.ScrollColumn = 4
    MyExcel.ActiveWindow.SmallScroll ToRight:=-2
    MyExcel.ActiveWindow.ScrollColumn = 1
    MyExcel.Sheets("Sheet3").Name = "����ͳ�Ʒֱ�"
    
        'MyExcel.cells(1, 1).Value = Timer
        'MyExcel.cells(1, 4).Value = time
    
    
    MyExcel.cells.Select
    With MyExcel.Selection.Font
        .Size = 9
        .ColorIndex = 0
    End With
    MyExcel.Range("A1").Select
    MyExcel.Selection.HorizontalAlignment = -4131
    
    MyExcel.cells(2, 1).Font.Bold = True
    MyExcel.cells(2, 1).Font.Bold = 5
    MyExcel.cells(2, 1).Value = "...ϵͳ��Ӧʱ��ͳ�Ʊ�..."
    
    MyExcel.cells(3, 1).Font.ColorIndex = 10
    MyExcel.cells(3, 1).Value = "��������̣�CHANNEL REQUEST ��ASSIGNMENT COMMAND��"
    If table_s(9, 1) = 0 Then
       MyExcel.cells(4, 1).Font.ColorIndex = 3
       MyExcel.cells(4, 1).Value = "�� CHANNEL REQUEST���� ASSIGNMENT COMMAND"
       RowNum = 6
       GoTo ewi
    End If
    MyExcel.Columns("A:A").ColumnWidth = 24.5 '21
    MyExcel.Columns("B:B").ColumnWidth = 5.63 '8.8
    MyExcel.Columns("C:C").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("D:D").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("E:E").ColumnWidth = 5.63 ' 8.88
    MyExcel.Columns("F:F").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("G:G").ColumnWidth = 5.63 '8.88
    MyExcel.Columns("H:H").ColumnWidth = 5.63 '8.88
    MyExcel.cells(4, 1).Value = ""
    
    MyExcel.Rows("4:4").HorizontalAlignment = -4108
    MyExcel.Rows("4:4").Font.Bold = True
    MyExcel.cells(4, 2).Value = "������"
    MyExcel.cells(4, 3).Value = "�����"
    MyExcel.cells(4, 4).Value = "��ֵ"
    MyExcel.cells(4, 5).Value = "��С��"
    MyExcel.cells(4, 6).Value = "%"
    MyExcel.cells(4, 7).Value = "�ۼ�%"
    Call Row_Col("0s<=x<1s", 1, 5, "B")
    Call Row_Col("0.1s<=x<0.2s", 2, 6, "B")
    Call Row_Col("0.2<=x<0.3s", 3, 7, "B")
    Call Row_Col("0.3s<=x<0.5s", 4, 8, "B")
    Call Row_Col("0.5s<=x<1s", 5, 9, "B")
    Call Row_Col("1s<=x<2s", 6, 10, "B")
    Call Row_Col("2s<=x<5s", 7, 11, "B")
    Call Row_Col("5s<=x<15s", 8, 12, "B")
    Call Row_Col("�ܼ�", 9, 13, "B")
    com_xmax = table_f(9, 2)
    com_xavg = table_f(9, 3)
    com_xmin = table_f(9, 4)
    RowNum = 15
    Frmrepot.ProgressBar1.Value = 52
ewi:
     '***************�����л�ͳ�Ʊ�
     Frmrepot.Label1.Caption = "���ڽ���ÿ�����л�Ƶ��ͳ�� ..."
     Frmrepot.Label1.Refresh
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
    MyExcel.cells(RowNum, 1).Font.Bold = True
    MyExcel.cells(RowNum, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum, 1).Value = "...ÿ�����л�Ƶ��ͳ�Ʊ�..."
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 10
    MyExcel.cells(RowNum + 1, 1).Value = "��������̣�ASSIGNMENT COMMAND��DISCONNECT֮��HANDOVER COMPLETE��Ƶ��"
    If table_s(9, 1) = 0 Then
        MyExcel.cells(RowNum + 2, 1).Font.ColorIndex = 3
        MyExcel.cells(RowNum + 2, 1).Value = "��ASSIGNMENT COMMAND����HANDOVER COMPLETE"
        RowNum = RowNum + 4
        GoTo ei
    End If
    MyExcel.cells(RowNum + 2, 1).Value = ""
    
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "%"
    MyExcel.cells(RowNum + 2, 4).Value = "�ۼ�%"
    AssignmntFlag = True
    Call Row_Col("0min<=x<1min", 1, RowNum + 3, "B")
    Call Row_Col("1min<=x<2min", 2, RowNum + 4, "B")
    Call Row_Col("2min<=x<3min", 3, RowNum + 5, "B")
    Call Row_Col("3min<=x<5min", 4, RowNum + 6, "B")
    Call Row_Col("5min<=x<6min", 5, RowNum + 7, "B")
    Call Row_Col("6min<=x<7min", 6, RowNum + 8, "B")
    Call Row_Col("7min<=x<8min", 7, RowNum + 9, "B")
    Call Row_Col("x>=8min", 8, RowNum + 10, "B")
    Call Row_Col("�ܼ�", 9, RowNum + 11, "B")
    AssignmntFlag = False
    com_hmax = table_f(9, 2)
    com_havg = table_f(9, 3)
    com_hmin = table_f(9, 4)
    RowNum = RowNum + 13
    Frmrepot.ProgressBar1.Value = 56
ei:
     '***************
    
    '*************************�л���������ͳ�Ʊ�
    Frmrepot.Label1.Caption = "���ڽ����л���������ͳ�� ..."
    Frmrepot.Label1.Refresh
    Dim enum1, enum2, enum3
    Call hand_time(enum1, enum2, enum3)
    MyExcel.cells(RowNum, 1).Font.Bold = True
    MyExcel.cells(RowNum, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum, 1).Value = "...�л���������ͳ�Ʊ�..."
    
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 10
    MyExcel.cells(RowNum + 1, 1).Value = "��������̣�HANDOVER COMMAND��HANDOVER COMPLETE��HANDOVER COMMAND FAILUER֮�䣩"
    If enum1 <= 0 Then
        MyExcel.cells(RowNum + 2, 1).Font.ColorIndex = 3
        MyExcel.cells(RowNum + 2, 1).Value = "��HANDOVER COMMAND"
        RowNum = RowNum + 4
        GoTo no_time
    End If
    MyExcel.cells(RowNum + 2, 1).Value = ""
    
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "�����"
    MyExcel.cells(RowNum + 2, 4).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 5).Value = "��С��"
    MyExcel.cells(RowNum + 2, 6).Value = "%"
    MyExcel.cells(RowNum + 2, 7).Value = "�ۼ�%"
    hand_zz
    Call Row_Col("0s<=x<0.1s", 1, RowNum + 3, "B")
    Call Row_Col("0.1s<=x<0.2s", 2, RowNum + 4, "B")
    Call Row_Col("0.2<=x<0.3s", 3, RowNum + 5, "B")
    Call Row_Col("0.3s<=x<0.5s", 4, RowNum + 6, "B")
    Call Row_Col("0.5s<=x<1s", 5, RowNum + 7, "B")
    Call Row_Col("1s<=x<2s", 6, RowNum + 8, "B")
    Call Row_Col("2s<=x<5s", 7, RowNum + 9, "B")
    Call Row_Col("5s<=x<15s", 8, RowNum + 10, "B")
    Call Row_Col("�ܼ�", 9, RowNum + 11, "B")
    com_hmax = table_f(9, 2)
    com_havg = table_f(9, 3)
    com_hmin = table_f(9, 4)
    RowNum = RowNum + 13
Frmrepot.ProgressBar1.Value = 60
    '***********************'�л����ʱ��ͳ�Ʊ�
    Frmrepot.Label1.Caption = "���ڽ����л����ʱ��ͳ�� ..."
    Frmrepot.Label1.Refresh
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
       
    MyExcel.cells(RowNum, 1).Font.Bold = True
    MyExcel.cells(RowNum, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum, 1).Value = "...�л����ʱ��ͳ�Ʊ�..."
    
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 10
    MyExcel.cells(RowNum + 1, 1).Value = "��������̣�HANDOVER COMMAND����һ��HANDOVER COMMAND ֮�䣩"
    If enum1 < 2 Then
        MyExcel.cells(RowNum + 2, 1).Font.ColorIndex = 3
        MyExcel.cells(RowNum + 2, 1).Value = "ֻ��һ��HANDOVER COMMAND"
        RowNum = RowNum + 4
        GoTo no_time
    End If
    MyExcel.cells(RowNum + 2, 1).Value = ""
    
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "�����"
    MyExcel.cells(RowNum + 2, 4).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 5).Value = "��С��"
    MyExcel.cells(RowNum + 2, 6).Value = "%"
    MyExcel.cells(RowNum + 2, 7).Value = "�ۼ�%"
   ' hand_zz
    Call Row_Col("0s<=x<1s", 1, RowNum + 3, "B")
    Call Row_Col("1s<=x<2s", 2, RowNum + 4, "B")
    Call Row_Col("2<=x<4s", 3, RowNum + 5, "B")
    Call Row_Col("4s<=x<10s", 4, RowNum + 6, "B")
    Call Row_Col("10s<=x<120s", 5, RowNum + 7, "B")
    Call Row_Col("2min<=x<20min", 6, RowNum + 8, "B")
    Call Row_Col("�ܼ�", 9, RowNum + 9, "B")
    com_hmax = table_f(9, 2)
    com_havg = table_f(9, 3)
    com_hmin = table_f(9, 4)
    RowNum = RowNum + 11
    Frmrepot.ProgressBar1.Value = 65
no_time:
     '*****************'˫Ƶ����ͳ�Ʊ�
     Frmrepot.Label1.Caption = "���ڽ���˫Ƶ����ͳ�� ..."
     Frmrepot.Label1.Refresh
    MyExcel.cells(RowNum, 1).Font.Bold = True
    MyExcel.cells(RowNum, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum, 1).Value = "...˫Ƶ����ͳ�Ʊ�..."
    MyExcel.Rows(RowNum + 1 & ":" & RowNum + 1).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 1 & ":" & RowNum + 1).Font.Bold = True
    MyExcel.cells(RowNum + 1, 2).Value = "������"
    MyExcel.cells(RowNum + 1, 3).Value = "%"
    Gsm_n = Gsm_Dcs("0", "125")
    Dcs_n = Gsm_Dcs("512", "886")
    GsmDcs_n = Gsm_n + Dcs_n
    Gsm_n1 = Format(Gsm_n / GsmDcs_n, "percent")
    Dcs_n1 = Format(Dcs_n / GsmDcs_n, "percent")
    
    MyExcel.cells(RowNum + 2, 1).Value = "GSM900"
    MyExcel.cells(RowNum + 2, 2).Value = Gsm_n
    MyExcel.cells(RowNum + 2, 3).Value = Gsm_n1
    '***********
    MyExcel.cells(RowNum + 3, 1).Value = "DCS1800"
    MyExcel.cells(RowNum + 3, 2).Value = Dcs_n
    MyExcel.cells(RowNum + 3, 3).Value = Dcs_n1
    MyExcel.cells(RowNum + 4, 1).Font.Bold = True
    MyExcel.cells(RowNum + 4, 1).Value = "�ܼ�"
    
    MyExcel.cells(RowNum + 4, 2).Value = GsmDcs_n
    MyExcel.Rows(Format(RowNum + 4) & ":" & Format(RowNum + 4)).Interior.ColorIndex = 15
    MyExcel.Rows(Format(RowNum + 4) & ":" & Format(RowNum + 4)).Interior.Pattern = 1
    RowNum = RowNum + 6
    If Dcs_n = 0 Then GoTo Dcs
    
  Frmrepot.ProgressBar1.Value = 68
    
  '****************************'Gsm900***�ֻ����͹���ͳ�Ʊ�***
Dcs:
    Frmrepot.Label1.Caption = "���ڽ���Gsm900�ֻ����͹���ͳ�� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(RowNum, 1).Font.Bold = True
    MyExcel.cells(RowNum, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum, 1).Value = "...�ֻ����͹���ͳ�Ʊ�..."
    
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 10
    MyExcel.cells(RowNum + 1, 1).Value = "GSM900"
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 4).Value = "%"
    MyExcel.cells(RowNum + 2, 5).Value = "�ۼ�%"
    Gsm900Dcs1800Flag = False
    SumNum = 0
             
             Set MyRecordset = dbs.OpenRecordset("SELECT rxqual_f FROM " & MyTableName, dbOpenDynaset)
             MyFieldType = MyRecordset.Fields(0).Type
             MyRecordset.Close
             Set MyRecordset = Nothing
  
    Call st_fill("0", "0", "27", True, True, True, "0 (43dBm)", False, RowNum + 3, "B")
    Call st_fill("1", "0", "27", False, True, True, "1 (41dBm)", False, RowNum + 4, "B")
    Call st_fill("2", "0", "27", False, True, True, "2 (39dBm)", False, RowNum + 5, "B")
    Call st_fill("3", "0", "27", False, True, True, "3 (37dBm)", False, RowNum + 6, "B")
    Call st_fill("4", "0", "27", False, True, True, "4 (35dBm)", False, RowNum + 7, "B")
    Call st_fill("5", "0", "27", False, True, True, "5 (33dBm)", False, RowNum + 8, "B")
  Frmrepot.ProgressBar1.Value = 70
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
    If SumNum > 0 Then
        all_avg = all_avg / SumNum
    End If
    putin1 = Format$(all_avg, "fixed")
    
    MyExcel.cells(RowNum + 23, 1).Font.Bold = True
    MyExcel.cells(RowNum + 23, 1).Value = "�ܼ�"
    
    MyExcel.Range("B" & RowNum + 23 & ":" & "C" & RowNum + 23).Font.Bold = True
    MyExcel.cells(RowNum + 23, 2).Value = all_0
    MyExcel.cells(RowNum + 23, 3).Value = putin1
    MyExcel.Rows(Format(RowNum + 23) & ":" & Format(RowNum + 23)).Interior.ColorIndex = 15
    MyExcel.Rows(Format(RowNum + 23) & ":" & Format(RowNum + 23)).Interior.Pattern = 1
    
    RowNum = RowNum + 24
    Frmrepot.ProgressBar1.Value = 71
    '****************************'***Gsm1800�ֻ����͹���ͳ�Ʊ�***
    Frmrepot.Label1.Caption = "���ڽ���Gsm1800�ֻ����͹���ͳ�� ..."
    Frmrepot.Label1.Refresh
    If Dcs_n = 0 Then GoTo Ta
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 10
    MyExcel.cells(RowNum + 1, 1).Value = "DCS1800"
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 4).Value = "%"
    MyExcel.cells(RowNum + 2, 5).Value = "�ۼ�%"
    Gsm900Dcs1800Flag = True
    all_avg = 0
    SumNum = 0
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
        If SumNum > 0 Then
            all_avg = all_avg / SumNum
        End If
        putin1 = Format$(all_avg, "fixed")
    Else
        putin1 = 0
    End If
    MyExcel.cells(RowNum + 19, 1).Value = "�ܼ�"
    MyExcel.Range("B" & RowNum + 19 & ":" & "C" & RowNum + 19).Font.Bold = True
    
    MyExcel.cells(RowNum + 19, 2).Value = all_0
    MyExcel.cells(RowNum + 19, 3).Value = putin1
    MyExcel.Rows(RowNum + 19 & ":" & RowNum + 19).Interior.ColorIndex = 15
    MyExcel.Rows(RowNum + 19 & ":" & RowNum + 19).Interior.Pattern = 1
    
    RowNum = RowNum + 20
   Frmrepot.ProgressBar1.Value = 74
   '*********************"***RXQUAL_FULLͳ�Ʊ�****"
Ta:
    Frmrepot.Label1.Caption = "���ڽ���RXQUAL_Fullͳ�� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(RowNum + 1, 1).Font.Bold = True
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum + 1, 1).Value = "...RXQUAL_Fullͳ�Ʊ�..."
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 4).Value = "%"
    MyExcel.cells(RowNum + 2, 5).Value = "�ۼ�%"
    RxlevFullFlag = True
    Call st_fill("7", "0", "23", True, True, False, "7 (12.8%<BER)", False, RowNum + 3, "B")
    Call st_fill("6", "0", "23", False, True, False, "6 (6.4%<BER<12.8%)", False, RowNum + 4, "B")
    Call st_fill("5", "0", "23", False, True, False, "5 (3.2%<BER<6.4%)", False, RowNum + 5, "B")
    Call st_fill("4", "0", "23", False, True, False, "4 (1.6%<BER<3.2%)", False, RowNum + 6, "B")
    Call st_fill("3", "0", "23", False, True, False, "3 (0.8%<BER<1.6%)", False, RowNum + 7, "B")
    Call st_fill("2", "0", "23", False, True, False, "2 (0.4%<BER<0.8%)", False, RowNum + 8, "B")
    Call st_fill("1", "0", "23", False, True, False, "1 (0.2%<BER<0.4%)", False, RowNum + 9, "B")
    Call st_fill("0", "0", "23", False, True, False, "0 (BER<0.2%)", False, RowNum + 10, "B")
    
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    MyExcel.Rows(RowNum + 11 & ":" & RowNum + 11).Font.Bold = True
    MyExcel.cells(RowNum + 11, 1).Value = "�ܼ�"
    
    MyExcel.cells(RowNum + 11, 2).Value = all_0
    MyExcel.cells(RowNum + 11, 3).Value = putin1
    MyExcel.Rows(RowNum + 11 & ":" & RowNum + 11).Interior.ColorIndex = 15
    MyExcel.Rows(RowNum + 11 & ":" & RowNum + 11).Interior.Pattern = 1
    
    RowNum = RowNum + 12
    
    Frmrepot.ProgressBar1.Value = 78
     '*********************"***RXQUAL_SUBͳ�Ʊ�****"
     Frmrepot.Label1.Caption = "���ڽ���RXQUAL_Subͳ�� ..."
     Frmrepot.Label1.Refresh
    MyExcel.cells(RowNum + 1, 1).Font.Bold = True
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum + 1, 1).Value = "...RXQUAL_SUBͳ�Ʊ�..."
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 4).Value = "%"
    MyExcel.cells(RowNum + 2, 5).Value = "�ۼ�%"
    RxlevFullFlag = False
    Call st_fill("7", "0", "23", True, True, False, "7 (12.8%<BER)", False, RowNum + 3, "B")
    Call st_fill("6", "0", "23", False, True, False, "6 (6.4%<BER<12.8%)", False, RowNum + 4, "B")
    Call st_fill("5", "0", "23", False, True, False, "5 (3.2%<BER<6.4%)", False, RowNum + 5, "B")
    Call st_fill("4", "0", "23", False, True, False, "4 (1.6%<BER<3.2%)", False, RowNum + 6, "B")
    Call st_fill("3", "0", "23", False, True, False, "3 (0.8%<BER<1.6%)", False, RowNum + 7, "B")
    Call st_fill("2", "0", "23", False, True, False, "2 (0.4%<BER<0.8%)", False, RowNum + 8, "B")
    Call st_fill("1", "0", "23", False, True, False, "1 (0.2%<BER<0.4%)", False, RowNum + 9, "B")
    Call st_fill("0", "0", "23", False, True, False, "0 (BER<0.2%)", False, RowNum + 10, "B")
    
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    MyExcel.Rows(RowNum + 11 & ":" & RowNum + 11).Font.Bold = True
    MyExcel.cells(RowNum + 11, 1).Value = "�ܼ�"
    
    MyExcel.cells(RowNum + 11, 2).Value = all_0
    MyExcel.cells(RowNum + 11, 3).Value = putin1
    MyExcel.Rows(RowNum + 11 & ":" & RowNum + 11).Interior.ColorIndex = 15
    MyExcel.Rows(RowNum + 11 & ":" & RowNum + 11).Interior.Pattern = 1
    
    RowNum = RowNum + 12
    
    Frmrepot.ProgressBar1.Value = 82
    '********************************'RXLEV_Fͳ�Ʊ�
    Frmrepot.Label1.Caption = "���ڽ���RXLEV_Fullͳ�� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(RowNum + 1, 1).Font.Bold = True
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum + 1, 1).Value = "...RXLEV_FULLͳ�Ʊ�..."
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.cells(RowNum + 1, 1).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "���ֵ"
    MyExcel.cells(RowNum + 2, 4).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 5).Value = "��Сֵ"
    MyExcel.cells(RowNum + 2, 6).Value = "%"
    MyExcel.cells(RowNum + 2, 7).Value = "�ۼ�%"
    RxlevFullFlag = True
    RangeNum = 3
    RxLevRange(1, 1) = "27"
    RxLevRange(1, 2) = "17"
    RxLevRange(1, 3) = "0"
    RxLevRange(2, 1) = "63"
    'RxLevRange(2, 1) = "150"
    RxLevRange(2, 2) = "27"
    RxLevRange(2, 3) = "17"
    '************
    For i = 1 To RangeNum
        If i = 1 Then
            If RxLevRange(2, i) = "63" Then
                Call st_fill(Format(Val(RxLevRange(1, i))), 150, "22", True, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 3, "B")
            Else
                Call st_fill(Format(Val(RxLevRange(1, i))), RxLevRange(2, i), "22", True, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 3, "B")
            End If
        Else
            If RxLevRange(2, i) = "63" Then
                Call st_fill(Format(Val(RxLevRange(1, i))), 150, "22", False, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 2 + i, "B")
            Else
                Call st_fill(Format(Val(RxLevRange(1, i))), RxLevRange(2, i), "22", False, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 2 + i, "B")
            End If
        End If
    Next
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    MyExcel.Rows(RowNum + 3 + i & ":" & RowNum + 3 + i).Font.Bold = True
    MyExcel.cells(RowNum + 3 + i, 1).Value = "�ܼ�"
    MyExcel.cells(RowNum + 3 + i, 2).Value = all_0
    MyExcel.cells(RowNum + 3 + i, 3).Value = putin
    MyExcel.cells(RowNum + 3 + i, 4).Value = putin1
    MyExcel.cells(RowNum + 3 + i, 5).Value = putin2
    MyExcel.Rows(RowNum + 3 + i & ":" & RowNum + 3 + i).Interior.ColorIndex = 15
    MyExcel.Rows(RowNum + 3 + i & ":" & RowNum + 3 + i).Interior.Pattern = 1
    
    RowNum = RowNum + 4 + i
    '***************************
    Frmrepot.ProgressBar1.Value = 85
    '********************************'RXLEV_Sͳ�Ʊ�
    Frmrepot.Label1.Caption = "���ڽ���RXLEV_Subͳ�� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(RowNum + 1, 1).Font.Bold = True
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum + 1, 1).Value = "...RXLEV_SUBͳ�Ʊ�..."
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "���ֵ"
    MyExcel.cells(RowNum + 2, 4).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 5).Value = "��Сֵ"
    MyExcel.cells(RowNum + 2, 6).Value = "%"
    MyExcel.cells(RowNum + 2, 7).Value = "�ۼ�%"
    RxlevFullFlag = False
    RangeNum = 3
    RxLevRange(1, 1) = "27"
    RxLevRange(1, 2) = "17"
    RxLevRange(1, 3) = "0"
    RxLevRange(2, 1) = "63"
    'RxLevRange(2, 1) = "150"
    RxLevRange(2, 2) = "27"
    RxLevRange(2, 3) = "17"
    '************
    For i = 1 To RangeNum
        If i = 1 Then
            If RxLevRange(2, i) = "63" Then
                Call st_fill(Format(Val(RxLevRange(1, i))), 150, "22", True, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 3, "B")
            Else
                Call st_fill(Format(Val(RxLevRange(1, i))), RxLevRange(2, i), "22", True, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 3, "B")
            End If
        Else
            If RxLevRange(2, i) = "63" Then
                Call st_fill(Format(Val(RxLevRange(1, i))), 150, "22", False, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 2 + i, "B")
            Else
                Call st_fill(Format(Val(RxLevRange(1, i))), RxLevRange(2, i), "22", False, False, False, RxLevRange(1, i) & "-" & RxLevRange(2, i) & " (-" & Format(110 - Val(RxLevRange(1, i))) & "<=dBm<-" & Format(110 - Val(RxLevRange(2, i))) & ")", False, RowNum + 2 + i, "B")
            End If
        End If
    Next
    all_0 = LTrim$(str(cc_all))
    all_avg = all_avg / cc_all
    putin1 = Format$(all_avg, "fixed")
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    MyExcel.Rows(RowNum + 3 + i & ":" & RowNum + 3 + i).Font.Bold = True
    MyExcel.cells(RowNum + 3 + i, 1).Value = "�ܼ�"
    MyExcel.cells(RowNum + 3 + i, 2).Value = all_0
    MyExcel.cells(RowNum + 3 + i, 3).Value = putin
    MyExcel.cells(RowNum + 3 + i, 4).Value = putin1
    MyExcel.cells(RowNum + 3 + i, 5).Value = putin2
    MyExcel.Rows(RowNum + 3 + i & ":" & RowNum + 3 + i).Interior.ColorIndex = 15
    MyExcel.Rows(RowNum + 3 + i & ":" & RowNum + 3 + i).Interior.Pattern = 1
    
    RowNum = RowNum + 4 + i
    
    
    Frmrepot.ProgressBar1.Value = 88
    '***************************
    '********************************'***Timing Advance(TA)ͳ�Ʊ�***
Frmrepot.Label1.Caption = "���ڽ���Timing Advanceͳ�� ..."
Frmrepot.Label1.Refresh
    MyExcel.cells(RowNum + 1, 1).Font.Bold = True
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum + 1, 1).Value = "...Timing Advance(TA)ͳ�Ʊ�..."
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "���ֵ"
    MyExcel.cells(RowNum + 2, 4).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 5).Value = "��Сֵ"
    MyExcel.cells(RowNum + 2, 6).Value = "%"
    MyExcel.cells(RowNum + 2, 7).Value = "�ۼ�%"
    TaFlag = True
    SumNum = 0
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
    If SumNum > 0 Then
        all_avg = all_avg / SumNum
    End If
    putin1 = Format$(all_avg, "fixed")
    putin = LTrim$(str(all_max))
    putin2 = LTrim$(str(all_min))
    MyExcel.Rows(RowNum + 13 & ":" & RowNum + 13).Font.Bold = True
    MyExcel.cells(RowNum + 13, 1).Value = "�ܼ�"
    MyExcel.cells(RowNum + 13, 2).Value = all_0
    MyExcel.cells(RowNum + 13, 3).Value = putin
    MyExcel.cells(RowNum + 13, 4).Value = putin1
    MyExcel.cells(RowNum + 13, 5).Value = putin2
    MyExcel.Rows(RowNum + 13 & ":" & RowNum + 13).Interior.ColorIndex = 15
    MyExcel.Rows(RowNum + 13 & ":" & RowNum + 13).Interior.Pattern = 1

    Frmrepot.ProgressBar1.Value = 92
    
    RowNum = RowNum + 14
    '***************************
    '********************************'***С��ѡ�����C1ͳ�Ʊ�**
    Frmrepot.Label1.Caption = "���ڽ���С��ѡ�����C1ͳ�� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(RowNum + 1, 1).Font.Bold = True
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum + 1, 1).Value = "...С��ѡ�����C1ͳ�Ʊ�... "
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "���ֵ"
    MyExcel.cells(RowNum + 2, 4).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 5).Value = "��Сֵ"
    MyExcel.cells(RowNum + 2, 6).Value = "%"
    MyExcel.cells(RowNum + 2, 7).Value = "�ۼ�%"
    
    MyC1C1Newflag = False
    For i = 1 To stre_num
        Set rst = dbs.OpenRecordset("select c1 from " & convert_filename(i) & " where c1<>""""")
        If rst.RecordCount > 0 Then
           MyC1C1Newflag = True
           Exit For
        End If
    Next
    If MyC1C1Newflag Then
        SumNum = 0
        TaFlag = False
        C1C2Flag = True
        Call st_fill("0", "0", "26", True, False, True, "C1<0", False, RowNum + 3, "B")
        CFlag = True
        Call st_fill("1", "0", "26", True, False, True, "C1=0", False, RowNum + 4, "B")
        CFlag = False
        Call st_fill("1", "10", "26", False, False, True, "1=<C1<10", False, RowNum + 5, "B")
        Call st_fill("10", "20", "26", False, False, True, "10=<C1<20", False, RowNum + 6, "B")
        Call st_fill("20", "30", "26", False, False, True, "20=<C1<30", False, RowNum + 7, "B")
        Call st_fill("30", "40", "26", False, False, True, "30=<C1<40", False, RowNum + 8, "B")
        Call st_fill("40", "50", "26", False, False, True, "40=<C1<50", False, RowNum + 9, "B")
        Call st_fill("50", "60", "26", False, False, True, "50=<C1<60", False, RowNum + 10, "B")
        Call st_fill("2", "60", "26", False, False, True, "C1>=60", False, RowNum + 11, "B")
        all_0 = LTrim$(str(cc_all))
        all_0 = all_0 + space$(9 - Len(all_0))
        If SumNum > 0 Then
            all_avg = all_avg / SumNum
        End If
        putin1 = Format$(all_avg, "fixed")
        putin = LTrim$(str(all_max))
        putin2 = LTrim$(str(all_min))
        MyExcel.Rows(RowNum + 12 & ":" & RowNum + 12).Font.Bold = True
        MyExcel.cells(RowNum + 12, 1).Value = "�ܼ�"
        MyExcel.cells(RowNum + 12, 2).Value = all_0
        MyExcel.cells(RowNum + 12, 3).Value = putin
        MyExcel.cells(RowNum + 12, 4).Value = putin1
        MyExcel.cells(RowNum + 12, 5).Value = putin2
        MyExcel.Rows(RowNum + 12 & ":" & RowNum + 12).Interior.ColorIndex = 15
        MyExcel.Rows(RowNum + 12 & ":" & RowNum + 12).Interior.Pattern = 1
    Else
        
        MyExcel.cells(RowNum + 3, 1).Value = "C1<0"
        MyExcel.cells(RowNum + 4, 1).Value = "C1=0"
        MyExcel.cells(RowNum + 5, 1).Value = "1=<C1<10"
        MyExcel.cells(RowNum + 6, 1).Value = "10=<C1<20"
        MyExcel.cells(RowNum + 7, 1).Value = "20=<C1<30"
        MyExcel.cells(RowNum + 8, 1).Value = "30=<C1<40"
        MyExcel.cells(RowNum + 9, 1).Value = "40=<C1<50"
        MyExcel.cells(RowNum + 10, 1).Value = "50=<C1<60"
        MyExcel.cells(RowNum + 11, 1).Value = "C1>=60"
        
        MyExcel.Rows(RowNum + 12 & ":" & RowNum + 12).Font.Bold = True
        MyExcel.cells(RowNum + 12, 1).Value = "�ܼ�"
        MyExcel.cells(RowNum + 12, 2).Value = all_0
        MyExcel.cells(RowNum + 12, 3).Value = "0"
        MyExcel.cells(RowNum + 12, 4).Value = "0"
        MyExcel.cells(RowNum + 12, 5).Value = "0"
        MyExcel.Rows(RowNum + 12 & ":" & RowNum + 12).Interior.ColorIndex = 15
        MyExcel.Rows(RowNum + 12 & ":" & RowNum + 12).Interior.Pattern = 1
    End If
    
    RowNum = RowNum + 13
    Frmrepot.ProgressBar1.Value = 96
    '********************************'***С��ѡ�����C2ͳ�Ʊ�**
    Frmrepot.Label1.Caption = "���ڽ���С��ѡ�����C2ͳ�� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(RowNum + 1, 1).Font.Bold = True
    MyExcel.cells(RowNum + 1, 1).Font.ColorIndex = 5
    MyExcel.cells(RowNum + 1, 1).Value = "...С��ѡ�����C2ͳ�Ʊ�... "
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).HorizontalAlignment = -4108
    MyExcel.Rows(RowNum + 2 & ":" & RowNum + 2).Font.Bold = True
    MyExcel.cells(RowNum + 2, 2).Value = "������"
    MyExcel.cells(RowNum + 2, 3).Value = "���ֵ"
    MyExcel.cells(RowNum + 2, 4).Value = "��ֵ"
    MyExcel.cells(RowNum + 2, 5).Value = "��Сֵ"
    MyExcel.cells(RowNum + 2, 6).Value = "%"
    MyExcel.cells(RowNum + 2, 7).Value = "�ۼ�%"
    
    If MyC1C1Newflag Then
        TaFlag = False
        SumNum = 0
        C1C2Flag = False
        Call st_fill("0", "0", "26", True, False, True, "C2<0", False, RowNum + 3, "B")
        CFlag = True
        Call st_fill("1", "0", "26", True, False, True, "C2=0", False, RowNum + 4, "B")
        CFlag = False
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
        If SumNum > 0 Then
            all_avg = all_avg / SumNum
        End If
        putin1 = Format$(all_avg, "fixed")
        putin = LTrim$(str(all_max))
        putin2 = LTrim$(str(all_min))
        MyExcel.Rows(RowNum + 16 & ":" & RowNum + 16).Font.Bold = True
        MyExcel.cells(RowNum + 16, 1).Value = "�ܼ�"
        MyExcel.cells(RowNum + 16, 2).Value = all_0
        MyExcel.cells(RowNum + 16, 3).Value = putin
        MyExcel.cells(RowNum + 16, 4).Value = putin1
        MyExcel.cells(RowNum + 16, 5).Value = putin2
        MyExcel.Rows(RowNum + 16 & ":" & RowNum + 16).Interior.ColorIndex = 15
        MyExcel.Rows(RowNum + 16 & ":" & RowNum + 16).Interior.Pattern = 1
    Else
        MyExcel.cells(RowNum + 3, 1).Value = "C2<0"
        MyExcel.cells(RowNum + 4, 1).Value = "C2=0"
        MyExcel.cells(RowNum + 5, 1).Value = "1=<C2<10"
        MyExcel.cells(RowNum + 6, 1).Value = "10=<C2<20"
        MyExcel.cells(RowNum + 7, 1).Value = "20=<C2<30"
        MyExcel.cells(RowNum + 8, 1).Value = "30=<C2<40"
        MyExcel.cells(RowNum + 9, 1).Value = "40=<C2<50"
        MyExcel.cells(RowNum + 10, 1).Value = "50=<C2<60"
        MyExcel.cells(RowNum + 11, 1).Value = "60=<C2<80"
        MyExcel.cells(RowNum + 12, 1).Value = "80=<C2<100"
        MyExcel.cells(RowNum + 13, 1).Value = "100=<C2<150"
        MyExcel.cells(RowNum + 14, 1).Value = "150=<C2<200"
        MyExcel.cells(RowNum + 15, 1).Value = "C2>=200"
        MyExcel.Rows(RowNum + 16 & ":" & RowNum + 16).Font.Bold = True
        MyExcel.cells(RowNum + 16, 1).Value = "�ܼ�"
        MyExcel.cells(RowNum + 16, 2).Value = all_0
        MyExcel.cells(RowNum + 16, 3).Value = "0"
        MyExcel.cells(RowNum + 16, 4).Value = "0"
        MyExcel.cells(RowNum + 16, 5).Value = "0"
        MyExcel.Rows(RowNum + 16 & ":" & RowNum + 16).Interior.ColorIndex = 15
        MyExcel.Rows(RowNum + 16 & ":" & RowNum + 16).Interior.Pattern = 1
    End If
    
'        MyExcel.cells(1, 2).Value = Timer
'        MyExcel.cells(1, 3).Value = Format(CLng(MyExcel.cells(1, 2).Value) - CLng(MyExcel.cells(1, 1).Value))
'        MyExcel.cells(1, 5).Value = time
    
    Frmrepot.Label1.Caption = "ͳ����ϣ����ڱ��汨�� ..."
    Frmrepot.Label1.Refresh
    rst.Close
    MynewRst.Close
    dbs.Close
    Set rst = Nothing
    Set MynewRst = Nothing
    Set dbs = Nothing
    'RowNum = RowNum + 17
    MyExcel.Sheets("����ͳ���嵥").Select
    MyExcel.Range("A1").Select
    Frmrepot.ProgressBar1.Value = 100
    Screen.MousePointer = 0
    'DoEvents
    'MyExcel.ChangeFileOpenDirectory AppPath + "\user\"
    MyExcel.ActiveWorkbook.Saveas filename:=stcname
    Unload Frmrepot
    '**************
    MyExcel.Visible = True
         
End Sub
Sub st_fill(a As String, b As String, col As String, ByVal sta As Boolean, ByVal rxq As Boolean, ByVal va As Boolean, ByVal fillhead As String, ByVal x9 As Boolean, ByVal RowFirst As Integer, ByVal RowNumber As String)
    Dim num, er_max, er_avg, er_min, er3, er4, er
    Dim num_z As Integer, max_z As Integer, avg_z As Single, min_z As Integer
    Dim Msg As String
    Dim fillnum As Integer, j As Integer
    Dim f1 As Integer, f2 As Integer
    Dim Zero As Boolean
    Dim myFilename As String
    Dim MyDbName As String, MyTableName As String
    Dim i As Integer
    Dim Mystring As String
    Static perc
    
    On Error GoTo errend
    DoEvents
    Zero = True
    num_z = 0
    max_z = 0
    If sta = True And Not CFlag Then
       all_max = 0
       all_min = 0
       all_avg = 0
       perc = 0
    End If
    If perc = 1 Then
       MyExcel.cells(RowFirst, 1).Value = fillhead
    
       Exit Sub
    End If
    
    For j = 1 To stre_num
        MyTableName = convert_filename(j)
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
                
                num = rst.Fields(0).Value
                SumNum = SumNum + num
                'er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                'er_min = rst.Fields(3).Value
                
          Else
          
             fillnum = 22 'rxqual
             If RxlevFullFlag Then
                If MyFieldType = 10 Then
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(rxqual_f) as countrxqual ," _
                        & " Avg(val(rxqual_f)) " _
                        & "AS Averagerxqual FROM " & MyTableName & " Where  rxqual_f = """ & a & """")
                Else
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(rxqual_f) as countrxqual ," _
                        & " Avg(rxqual_f) " _
                        & "AS Averagerxqual FROM " & MyTableName & " Where  rxqual_f = " & a & "")
                End If
                num = rst.Fields(0).Value
                'er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                'er_min = rst.Fields(3).Value
                
             Else
                If MyFieldType = 10 Then
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(rxqual_s) as countrxqual ," _
                        & " Avg(val(rxqual_s)) " _
                        & "AS Averagerxqual FROM " & MyTableName & " Where  rxqual_s = """ & a & """")
                Else
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(rxqual_s) as countrxqual ," _
                        & " Avg(rxqual_s) " _
                        & "AS Averagerxqual FROM " & MyTableName & " Where  rxqual_s = " & a & "")
                End If
                num = rst.Fields(0).Value
                'er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                'er_min = rst.Fields(3).Value
                
             
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
                num = rst.Fields(0).Value
                SumNum = SumNum + num
                
                er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                er_min = rst.Fields(3).Value
                
              Else
                 Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(ta) as countta ," _
                    & " Avg(val(ta)) " _
                    & "AS Averagta, Max(val(ta)) " _
                    & "AS Maximumta, Min(val(ta)) " _
                    & "As Minta FROM " & MyTableName & " Where  int(ta) > " & a & "" & " and int(ta) <= " & b & "")
                num = rst.Fields(0).Value
                SumNum = SumNum + num
                
                er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                er_min = rst.Fields(3).Value
                
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
                        num = rst.Fields(0).Value
                        SumNum = SumNum + num
                        er_max = rst.Fields(2).Value
                        er_avg = rst.Fields(1).Value
                        er_min = rst.Fields(3).Value
                        
                    Else
                        Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c1) as countta ," _
                        & " Avg(val(c1)) " _
                        & "AS Averagta, Max(val(c1)) " _
                        & "AS Maximumta, Min(val(c1)) " _
                        & "As Minta FROM " & MyTableName & " Where  c1 = """ & a & """")
                        num = rst.Fields(0).Value
                        SumNum = SumNum + num
                        
                        er_max = rst.Fields(2).Value
                        er_avg = rst.Fields(1).Value
                        er_min = rst.Fields(3).Value
                        
                    
                    End If
              Else
                 If a <> "2" Then
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c1) as countta ," _
                        & " Avg(val(c1)) " _
                        & "AS Averagta, Max(val(c1)) " _
                        & "AS Maximumta, Min(val(c1)) " _
                        & "As Minta FROM " & MyTableName & " Where  int(c1) >= " & a & "" & " and int(c1) < " & b & "")
                    num = rst.Fields(0).Value
                    SumNum = SumNum + num
                    er_max = rst.Fields(2).Value
                    er_avg = rst.Fields(1).Value
                    er_min = rst.Fields(3).Value
                    
                Else
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c1) as countta ," _
                        & " Avg(val(c1)) " _
                        & "AS Averagta, Max(val(c1)) " _
                        & "AS Maximumta, Min(val(c1)) " _
                        & "As Minta FROM " & MyTableName & " Where  int(c1) >= """ & b & """")
                    num = rst.Fields(0).Value
                    SumNum = SumNum + num
                    er_max = rst.Fields(2).Value
                    er_avg = rst.Fields(1).Value
                    er_min = rst.Fields(3).Value
                                    
                
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
                        num = rst.Fields(0).Value
                        SumNum = SumNum + num
                        
                        er_max = rst.Fields(2).Value
                        er_avg = rst.Fields(1).Value
                        er_min = rst.Fields(3).Value
                        
                    Else
                        Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c2) as countta ," _
                        & " Avg(val(c2)) " _
                        & "AS Averagta, Max(val(c2)) " _
                        & "AS Maximumta, Min(val(c2)) " _
                        & "As Minta FROM " & MyTableName & " Where  c2 = """ & a & """")
                        num = rst.Fields(0).Value
                        SumNum = SumNum + num
                        
                        er_max = rst.Fields(2).Value
                        er_avg = rst.Fields(1).Value
                        er_min = rst.Fields(3).Value
                        
                    
                    End If
              Else
                 If a <> "2" Then
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c2) as countta ," _
                        & " Avg(val(c2)) " _
                        & "AS Averagta, Max(val(c2)) " _
                        & "AS Maximumta, Min(val(c2)) " _
                        & "As Minta FROM " & MyTableName & " Where  int(c2) >= " & a & "" & " and int(c2) < " & b & "")
                    num = rst.Fields(0).Value
                    SumNum = SumNum + num
                        
                    er_max = rst.Fields(2).Value
                    er_avg = rst.Fields(1).Value
                    er_min = rst.Fields(3).Value
                    
                Else
                    Set rst = dbs.OpenRecordset("SELECT  " _
                        & " Count(c2) as countta ," _
                        & " Avg(val(c2)) " _
                        & "AS Averagta, Max(val(c2)) " _
                        & "AS Maximumta, Min(val(c2)) " _
                        & "As Minta FROM " & MyTableName & " Where  int(c2) >= " & b & "")
                    num = rst.Fields(0).Value
                    SumNum = SumNum + num
                         
                    er_max = rst.Fields(2).Value
                    er_avg = rst.Fields(1).Value
                    er_min = rst.Fields(3).Value
                
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
                num = rst.Fields(0).Value
                er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                er_min = rst.Fields(3).Value
                
             Else
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(rxlev_s) as countrxlev ," _
                    & " Avg(rxlev_s) " _
                    & "AS Averagerxlev, Max(rxlev_s) " _
                    & "AS Maximumrxlev,Min(rxlev_s) " _
                    & " as Minrxlev FROM " & MyTableName & " Where  rxlev_s >= " & Format(a) & " and  rxlev_s < " & Format(b))
                num = rst.Fields(0).Value
                er_max = rst.Fields(2).Value
                er_avg = rst.Fields(1).Value
                er_min = rst.Fields(3).Value
                
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
            MyExcel.cells(RowFirst, 1).Value = fillhead
    
            'Print #FileNumber, fillhead
         Else
            If rxq = True Then
                '************
                MyExcel.cells(RowFirst, 1).Value = fillhead
                MyExcel.cells(RowFirst, RowNumber).Value = num
                
                Mystring = Asc(RowNumber)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.cells(RowFirst, Mystring).Value = er_avg
                   
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.cells(RowFirst, Mystring).Value = er3
                   
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.cells(RowFirst, Mystring).Value = er4
                  
                  'Print #FileNumber, fillhead; Space$(fillnum); num; er_avg; er3; er4
                
                '*********
            Else
                MyExcel.cells(RowFirst, 1).Value = fillhead
                MyExcel.cells(RowFirst, RowNumber).Value = num
                
                Mystring = Asc(RowNumber)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.cells(RowFirst, Mystring).Value = er_max 'er_avg
                   
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.cells(RowFirst, Mystring).Value = er_avg
                   
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.cells(RowFirst, Mystring).Value = er_min 'er4
                
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.cells(RowFirst, Mystring).Value = er3 'er4
                
                Mystring = Asc(Mystring)
                i = Val(Mystring) + 1
                Mystring = Chr(i)
                MyExcel.cells(RowFirst, Mystring).Value = er4
                
                
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
    
    If kk = 9 Then
       MyExcel.cells(RowFirst, 1).Font.Bold = True
    End If
    
    MyExcel.cells(RowFirst, 1).Value = gg
    If table_s(kk, 1) = 0 Then GoTo tr
      
       '*********

    
    If kk = 9 Then
       MyExcel.cells(RowFirst, RowNumber).Font.Bold = True
    End If
    MyExcel.cells(RowFirst, RowNumber).Value = table_f(kk, 1)
    Mystring = Format(RowNumber)
    If Not AssignmntFlag Then
    
    Mystring = Asc(Mystring)
    i = Val(Mystring) + 1
    Mystring = Chr(i)
    
    If kk = 9 Then
       MyExcel.cells(RowFirst, Mystring).Font.Bold = True
    End If
    
    MyExcel.cells(RowFirst, Mystring).Value = table_f(kk, 2)
    
    Mystring = Asc(Mystring)
    i = Val(Mystring) + 1
    Mystring = Chr(i)
    
    If kk = 9 Then
       MyExcel.cells(RowFirst, Mystring).Font.Bold = True
    End If
    
    MyExcel.cells(RowFirst, Mystring).Value = table_f(kk, 3)
    
    Mystring = Asc(Mystring)
    i = Val(Mystring) + 1
    Mystring = Chr(i)
    
    If kk = 9 Then
       MyExcel.cells(RowFirst, Mystring).Font.Bold = True
    End If
    
    MyExcel.cells(RowFirst, Mystring).Value = table_f(kk, 4)
    End If
    If kk = 9 Then
       MyExcel.Rows(Format(RowFirst) & ":" & Format(RowFirst)).Interior.ColorIndex = 15
        MyExcel.Rows(Format(RowFirst) & ":" & Format(RowFirst)).Interior.Pattern = 1
       GoTo tr
    Else
        'MyExcel.Selection.Font.Bold = True
        Mystring = Asc(Mystring)
        i = Val(Mystring) + 1
        Mystring = Chr(i)
        MyExcel.cells(RowFirst, Mystring).Value = table_f(kk, 5)
        
        Mystring = Asc(Mystring)
        i = Val(Mystring) + 1
        Mystring = Chr(i)
        MyExcel.cells(RowFirst, Mystring).Value = table_f(kk, 6)
        
    End If
    
tr:
    'MyExcel.Range("A" & Format(RowFirst) + 1).Select
End Sub

Function mess_num(ByVal tabna As String) As Long
    Dim j As Integer
    Dim menum As Long
    
    On Error Resume Next
    DoEvents
    menum = 0
    For j = 1 To stre_num
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & " count(*) " _
                    & "AS countmessage FROM " & convert_filename(j) & " Where  message = """ & tabna & """")
        menum = menum + rst.Fields(0).Value
    Next
    mess_num = menum
    
End Function

Function Mark_num(ByVal tabna As String) As Long
    Dim j As Integer
    Dim menum As Long
    
    On Error Resume Next
    menum = 0
    For j = 1 To stre_num
        Set rst = dbs.OpenRecordset("SELECT count(*) " _
                    & "AS countmessage FROM " & convert_filename(j) & " Where  mark = """ & tabna & """")
        menum = menum + rst.Fields(0).Value
    Next
    Mark_num = menum

End Function

Function Gsm_Dcs(ByVal a As String, ByVal b As String) As Long
    Dim j As Integer
    Dim menum As Long
    
    On Error Resume Next
    DoEvents
    menum = 0
    For j = 1 To stre_num
       If a = "0" Then
            Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(*) " _
                    & "AS countmessage FROM " & convert_filename(j) & " where BCCH_SERV=null or BCCH_SERV >= " & Format(a) & " and  BCCH_SERV <" & Format(b))
       Else
            Set rst = dbs.OpenRecordset("SELECT  " _
                    & " Count(*) " _
                    & "AS countmessage FROM " & convert_filename(j) & " where BCCH_SERV >= " & Format(a) & " and  BCCH_SERV <" & Format(b))
       End If
        menum = menum + rst.Fields(0).Value
    Next
    Gsm_Dcs = menum

End Function

Sub hand_time(getnum1, getnum2, getnum3)
    Dim h_comm() As Single, h_comp() As Single, h_fail() As Single
    Dim tp(1 To 3000) As Single
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
    ReDim h_comm(1 To 3000) As Single
    ReDim h_comp(1 To 3000) As Single
    ReDim h_fail(1 To 3000) As Single
    Dim MyTableName As String
    Dim MyDbName As String
    Dim MyNewtmp As Integer
    
    On Error Resume Next
    menum = 0
    getnum1 = 0
    getnum2 = 0
    getnum3 = 0
    For j = 1 To 9
        For k = 1 To 6
            table_s(j, k) = 0
            table_f(j, k) = ""
            tibeh(j, k) = 0
        Next k
    Next j
    
    For HH = 1 To stre_num
        MyTableName = convert_filename(HH)
        For pp = 1 To 3000
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

                menum = rst.RecordCount
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
                        h_comm(i) = h_comm(i) + Val("0." & dd)
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
                        
                        h_comp(i) = h_comp(i) + Val("0." & dd)
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
                         
                         h_fail(i) = h_fail(i) + Val("0." & dd)
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
    'For bbb = 1 To getnum1
    
    If getnum2 + getnum3 > getnum1 Then
        MyNewtmp = getnum2 + getnum3
    Else
        MyNewtmp = getnum1
    End If
    For bbb = 1 To MyNewtmp
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
    ReDim h_comm(1 To 3000) As Single
    ReDim h_comp(1 To 3000) As Single
    
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
        For pp = 1 To 3000
            h_comm(pp) = 0
            h_comp(pp) = 0
        Next
        
        MyTableName = convert_filename(HH)
        
        DoEvents
        If hu = False Then
            If hg = True Then
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "CONNECT ACKNOWLEDGE" & """")
            Else
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "CHANNEL REQUEST" & """ or message = """ & "CHANNEL REQUEST REPORT" & """")

            End If
        Else
           Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "SETUP" & """ or message = """ & "EMERGENCY SETUP" & """")
        End If
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
                menum = rst.RecordCount
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

                         h_comm(i) = h_comm(i) + Val("0." & dd)
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
                
          ' mapinfo.Do "select * from " + Chr(34) + stre_tab(hh) + Chr(34) + " where col5 = " + Chr(34) + "ASSIGNMENT COMMAND" + Chr(34) + "into temp order by col1"
        Else
          Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "ASSIGNMENT COMPLETE" & """")
                DoEvents
        End If
      Else
         If hu = False Then
           Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "DISCONNECT" & """")
               
                DoEvents
         Else
           Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "CONNECT" & """")
                DoEvents
         End If
      End If
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
      
      menum = rst.RecordCount
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
                     h_comp(i) = h_comp(i) + Val("0." & dd)
               Else
                    h_comp(i) = h_comp(i) + Val(dd)
                    dd = 0
               End If
               If i < menum Then
                    rst.MoveNext
               End If
           Next
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
                 Do While h_comm(i + 1) <> -1 And h_comm(i + 1) <> 0
                    If h_comm(i + 1) > h_comp(j) Then
                        Exit Do
                    End If
                    i = i + 1
                 Loop
                 If h_comm(i) = -1 Or h_comm(i) = 0 Then
                    Exit Do
                 End If
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
    ReDim h_comm(1 To 3000) As Single
    ReDim h_comp(1 To 3000) As Single
    ReDim h_compt(1 To 3000) As Single
    'Dim k As Integer
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
        For pp = 1 To 3000
            h_comm(pp) = 0
            h_comp(pp) = 0
        Next
        
        MyTableName = convert_filename(HH)
        
        DoEvents
        If hu = False Then
            If hg = False Then
                Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "ASSIGNMENT COMMAND" & """")
            End If
        End If
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
        
                menum = rst.RecordCount
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
                        h_comm(i) = h_comm(i) + Val("0." & dd)
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
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "HANDOVER COMPLETE" & """")
        End If
      End If
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
      
           menum = rst.RecordCount
           If menum > 0 Then
           rst.MoveFirst
           For i = 1 To menum
               dd = rst.Fields(0).Value
               finds = InStr(dd, ":")
               mid_dd = Val(Left(dd, finds - 1)) * 3600
               h_compt(i) = mid_dd
               dd = Right(dd, Len(dd) - finds)
               finds = InStr(dd, ":")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1)) * 60
                    h_compt(i) = h_compt(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
               End If
               finds = InStr(dd, ".")
               If finds > 0 Then
                    mid_dd = Val(Left(dd, finds - 1))
                    h_compt(i) = h_compt(i) + mid_dd
                    dd = Right(dd, Len(dd) - finds)
                    h_compt(i) = h_compt(i) + Val("0." & dd)
               Else
                    h_compt(i) = h_compt(i) + Val(dd)
                    dd = 0
               End If
               If i < menum Then
                    rst.MoveNext
               End If
           Next
           h_compt(i) = -1
        End If
      
    '***********
      If hg = False Then
        If hu = False Then
          Set rst = dbs.OpenRecordset("SELECT  " _
                    & " time " _
                    & "AS timemessage FROM " & MyTableName & " Where  message = """ & "DISCONNECT" & """")
        End If
      End If
                If rst.RecordCount <> 0 Then
                    rst.MoveLast
                End If
      
           menum = rst.RecordCount
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
                     h_comp(i) = h_comp(i) + Val("0." & dd)
               Else
                    h_comp(i) = h_comp(i) + Val(dd)
                    dd = 0
               End If
               If i < menum Then
                    rst.MoveNext
               End If
           Next
           h_comp(i) = -1
        End If
   '****************************
           i = 1
           j = 1
           k = 1
           tal = 0
           Do While h_comm(i) <> -1 And h_comm(i) <> 0
              hav = False
              If h_comm(i) <= h_comp(j) Then
                 For assi = 1 To 100
                 If h_compt(k) < h_comp(j) And h_compt(k) > h_comm(i) Then
                    If assi = 1 Then
                       ggg = h_compt(k) - h_comm(i)
                       k = k + 1
                    ElseIf assi > 1 Then
                       ggg = h_compt(k) - h_compt(k - 1)
                       k = k + 1
                    End If
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
                Else
                    For assj = 1 To 100
                        If h_compt(k) < h_comm(i) Then
                           k = k + 1
                        Else
                           Exit For
                        End If
                    Next
                    'If assi = 1 Then k = k + 1
                    j = j + 1
                    i = i + 1
                    Exit For
                 End If
                Next
                 'ggg = h_comp(j) - h_comm(i)
                 'j = j + 1
                 'i = i + 1
                 'hav = True
              Else
                 j = j + 1
                 If h_comp(j) = -1 Or h_comp(j) = 0 Then Exit Do
              End If
              
           Loop
next1:
    
    Next HH
errend:

End Sub

Function Mark1_num(ByVal tabna As String) As Long
    Dim j As Integer
    Dim menum As Long
    
    On Error Resume Next
    menum = 0
    For j = 1 To stre_num
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & " count(*) " _
                    & "AS countmessage FROM " & convert_filename(j) & " Where  left( trim(mark1),2) = """ & tabna & """")
        menum = menum + rst.Fields(0).Value
    Next
    Mark1_num = menum

End Function

Function Mark1_cause(ByVal tabna As String) As Long
    Dim j As Integer
    Dim menum As Long
    
    On Error Resume Next
    menum = 0
    For j = 1 To stre_num
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & " count(*) " _
                    & "AS countmessage FROM " & convert_filename(j) & " Where  left( trim(mark1),7) = """ & tabna & """")
        menum = menum + rst.Fields(0).Value
    Next
    Mark1_cause = menum
End Function

Function Mark1_Talk(ByVal tabna As String) As Long
    Dim j As Integer
    Dim menum As Long
    
    On Error Resume Next
    menum = 0
    For j = 1 To stre_num
      If HoASFLFlag Then
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & " count(*) " _
                    & "AS countmessage FROM " & convert_filename(j) & " Where  left( trim(mark1),3) = """ & tabna & """")
                menum = menum + rst.Fields(0).Value
      Else
        Set rst = dbs.OpenRecordset("SELECT  " _
                    & " count(*) " _
                    & "AS countmessage FROM " & convert_filename(j) & " Where  left( trim(mark1),5) = """ & tabna & """")
                menum = menum + rst.Fields(0).Value
      End If
    Next
    Mark1_Talk = menum

End Function

Sub CallAttemp()
    Dim CA As Integer, setup_n As Integer, setup_n1 As Integer
    Dim CS As Integer
    Dim CF As Integer
    Dim CFTime As Integer, CFrelease As Integer, CFNorse As Integer, CFOther As Integer
    Dim CFBroked As Integer, CFNo As Integer, CFSys As Integer
    Dim CFputin, HofFail, LuaFail, CdDropped
    Dim CpUl As Integer, CpDl As Integer, Hoa As Integer, HoaIn As Integer, HoaCell As Integer, HoaSys As Integer
    Dim Hos As Integer, Hof As Integer, Lur As Integer, Lua As Integer, Luf As Integer, Luf1 As Integer, Luf2 As Integer
    Dim Hand_1 As Integer
    Dim CdRelease As Integer, CdDr As Integer, CdHandDr As Integer, CDNoDr As Integer
    Dim MySetupTmp As Integer
    Dim i As Integer, j As Integer
    Dim MOCCall As Integer, MTCCall As Integer
    
    On Error Resume Next
    
    'setup_n = mess_num("SETUP") '�Ժ�����
    'setup_n1 = mess_num("EMERGENCY SETUP")
    'setup_n = setup_n + setup_n1
    
    'setup_n = Mark_num("Start Call")

Frmrepot.ProgressBar1.Value = 7
    Frmrepot.Label1.Caption = "����ͳ�ƺ��н������� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(3, 1).Value = "1.���н�������"
    
    MOCCall = 0
    For j = 1 To stre_num
        Set rst = dbs.OpenRecordset("SELECT count(*) " _
                    & "AS countmessage FROM " & convert_filename(j) & " Where  mark = ""Start Call"" and bsic_same2 =1 ")
        MOCCall = MOCCall + rst.Fields(0).Value
    Next
    MTCCall = 0
    For j = 1 To stre_num
        Set rst = dbs.OpenRecordset("SELECT count(*) " _
                    & "AS countmessage FROM " & convert_filename(j) & " Where  mark = ""Start Call"" and bsic_same2 =2 ")
        MTCCall = MTCCall + rst.Fields(0).Value
    Next
    MyExcel.cells(5, 1).Value = "���в������"
    MyExcel.cells(5, 2).Value = MOCCall
    MyExcel.cells(6, 1).Value = "���в������"
    MyExcel.cells(6, 2).Value = MTCCall
    setup_n = MOCCall + MTCCall
    'MyExcel.Range("A4").Select
    'MyExcel.ActiveCell.FormulaR1C1 = "���н������Դ�����"
    
    Frmrepot.Label1.Caption = "����ͳ�ƽ���ͨ������ ..."
Frmrepot.Label1.Refresh
    'CA = Mark1_num("CA")
    'If CA < setup_n Then
    '    CA = setup_n
    'End If
    'MyExcel.Range("B4").Select
    'MyExcel.ActiveCell.FormulaR1C1 = CA
    MyExcel.cells(7, 1).Value = "���н����ɹ�������"
    CS = Mark1_num("CS")
    
    Frmrepot.Label1.Caption = "����ͳ�ƽ���ʧ�ܴ��� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(7, 2).Value = CS
    MyExcel.cells(8, 1).Value = "���н���ʧ�ܴ�����"
    
                        CF = 0
                        CFTime = 0
                        CFrelease = 0
                        CFNorse = 0
                        CFOther = 0
                        CFBroked = 0
                        CFNo = 0
                        CFSys = 0
                
                For i = 1 To stre_num
                    Set rst = dbs.OpenRecordset("select mark1 from " & convert_filename(i) & " where left(mark1,2)=""CF""")
                    If rst.RecordCount > 0 Then
                        rst.MoveLast
                        rst.MoveFirst
                        CF = CF + rst.RecordCount
                        For j = 1 To rst.RecordCount
                            Select Case Left(Trim(rst.Fields("mark1").Value), 6)
                                'Case "CF ��ͨǰ��"
                                '    CFTime = CFTime + 1
                                'Case "CF ��ͨǰ��"
                                '    CFrelease = CFrelease + 1
                                Case "CF �޷���"
                                    CFNorse = CFNorse + 1
                                'Case "CF �Է�ռ��"
                                '    CFOther = CFOther + 1
                                Case "CF ����ӵ"
                                    CFBroked = CFBroked + 1
                                'Case "CF ����δ��"
                                '    CFNo = CFNo + 1
                                Case "CF ��Ӧ��"
                                    CFSys = CFSys + 1
                            End Select
                            rst.MoveNext
                        Next
                    End If
                Next
                    
    MyExcel.cells(8, 2).Value = CF
    MyExcel.cells(9, 1).Value = "               ʧ��ԭ�����ͳ�ƣ�"
    MyExcel.cells(10, 1).Value = "                     1.����ӵ��"
    MyExcel.cells(10, 2).Value = CFBroked
    MyExcel.cells(11, 1).Value = "                     2.�Ƿ�����"
    MyExcel.cells(11, 2).Value = CFNorse
    MyExcel.cells(12, 1).Value = "                     3.��Ӧ��"
    MyExcel.cells(12, 2).Value = CFSys
    MyExcel.cells(13, 1).Value = "���н���ʧ���ʣ�"
    If setup_n > 0 Then
        CFputin = Format((CFBroked + CFNorse + CFSys) / setup_n, "percent") '���н���ʧ����
    End If
    MyExcel.cells(13, 2).Value = CFputin
    HoASFLFlag = False
    Frmrepot.ProgressBar1.Value = 12
    Frmrepot.Label1.Caption = "����ͳ��ͨ������ ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(17 - 2, 1).Value = "2.ͨ������"  '15
    
    Frmrepot.Label1.Caption = "����ͳ������������� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(18 - 2, 1).Value = "�������������"
    CpUl = Mark1_Talk("CP UL")
    MyExcel.cells(18 - 2, 2).Value = CpUl
    
    Frmrepot.Label1.Caption = "����ͳ������������� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(19 - 2, 1).Value = "�������������"
    CpDl = Mark1_Talk("CP DL")
    MyExcel.cells(19 - 2, 2).Value = CpDl
    
    
    Frmrepot.Label1.Caption = "����ͳ���л����Դ��� ..."
    Frmrepot.Label1.Refresh
    
                HoaIn = 0
                HoaCell = 0
                HoaSys = 0
                Hos = 0
                Hof = 0
                For i = 1 To stre_num
                    Set rst = dbs.OpenRecordset("select mark1 from " & convert_filename(i) & " where left(mark1,2)=""HO""")
                    If rst.RecordCount > 0 Then
                        rst.MoveLast
                        rst.MoveFirst
                        For j = 1 To rst.RecordCount
                            Select Case Left(Trim(rst.Fields("mark1").Value), 3)
                                Case "HOA"
                                    Select Case Left(Trim(rst.Fields("mark1").Value), 5)
                                        Case "HOA 1"
                                            HoaIn = HoaIn + 1
                                        Case "HOA 2"
                                            HoaCell = HoaCell + 1
                                        Case "HOA 3"
                                            HoaSys = HoaSys + 1
                                    End Select
                                Case "HOS"
                                    Hos = Hos + 1
                                Case "HOF"
                                    Hof = Hof + 1
                            End Select
                            rst.MoveNext
                        Next
                    End If
                Next
    
    
    MyExcel.cells(20 - 2, 1).Value = "�л����Դ�����"
    MyExcel.cells(21 - 2, 1).Value = "              1.ʱ϶�л�����"
    MyExcel.cells(21 - 2, 2).Value = HoaIn
    
    MyExcel.cells(22 - 2, 1).Value = "              2.С���л�����"
    MyExcel.cells(22 - 2, 2).Value = HoaCell
    
    MyExcel.cells(23 - 2, 1).Value = "              3.ϵͳ�л�����"
    MyExcel.cells(23 - 2, 2).Value = HoaSys
    Hoa = HoaIn + HoaCell + HoaSys
    MyExcel.cells(20 - 2, 2).Value = Hoa  '�л����Դ���
    Frmrepot.ProgressBar1.Value = 16
    HoASFLFlag = True
    MyExcel.cells(24 - 2, 1).Value = "�л��ɹ�������"
    
    'Hand_1 = mess_num("HANDOVER COMMAND") '�л�����
    Frmrepot.Label1.Caption = "����ͳ���л��ɹ����� ..."
    Frmrepot.Label1.Refresh
    'Hos = Mark1_Talk("HOS")
    MyExcel.cells(24 - 2, 2).Value = Hos
    Frmrepot.Label1.Caption = "����ͳ���л�ʧ�ܴ��� ..."
    Frmrepot.Label1.Refresh
    MyExcel.cells(25 - 2, 1).Value = "�л�ʧ�ܴ�����"
    'Hof = Mark1_Talk("HOF")
    
    If Hoa < Hof + Hos Then
        MyExcel.cells(22 - 2, 2).Value = HoaCell + (Hof + Hos - Hoa)
       Hoa = Hof + Hos
     '  Hand_1 = Hof + Hos
        MyExcel.cells(20 - 2, 2).Value = Hoa  '�л����Դ���
    End If
    
    'If Hof > Hand_1 Then Hof = Hand_1
    If Hoa > 0 Then
        HofFail = Format(Hof / Hoa, "percent")  'ʧ���л���
    End If
    
    MyExcel.cells(25 - 2, 2).Value = Hof
    MyExcel.cells(26 - 2, 1).Value = "             ʧ��ԭ�����ͳ�ƣ�"
    MyExcel.cells(27 - 2, 1).Value = "                        1.RRԭ��"
    MyExcel.cells(27 - 2, 2).Value = Hof
    MyExcel.cells(28 - 2, 1).Value = "                        2.����ԭ��"
    MyExcel.cells(29 - 2, 1).Value = "�л�ʧ���ʣ�"
    MyExcel.cells(29 - 2, 2).Value = HofFail
    
    Frmrepot.Label1.Caption = "����ͳ��λ�ø��³��Դ��� ..."
    Frmrepot.Label1.Refresh
    Frmrepot.ProgressBar1.Value = 19
    
                Lur = 0
                Lua = 0
                Luf = 0
                Luf1 = 0
                Luf2 = 0
                For i = 1 To stre_num
                    Set rst = dbs.OpenRecordset("select mark1 from " & convert_filename(i) & " where left(mark1,2)=""LU""")
                    If rst.RecordCount > 0 Then
                        rst.MoveLast
                        rst.MoveFirst
                        For j = 1 To rst.RecordCount
                            Select Case Left(Trim(rst.Fields("mark1").Value), 3)
                                Case "LUR"
                                    Lur = Lur + 1
                                Case "LUA"
                                    Lua = Lua + 1
                                Case "LUF"
                                    If Left(Trim(rst.Fields("mark1").Value), 5) = "LUF 1" Then
                                       Luf1 = Luf1 + 1
                                    ElseIf Left(Trim(rst.Fields("mark1").Value), 5) = "LUF 2" Then
                                       Luf2 = Luf2 + 1
                                    End If
                                    Luf = Luf + 1
                            End Select
                            rst.MoveNext
                        Next
                    End If
                Next
    
    MyExcel.cells(30 - 2, 1).Value = "λ�ø��³��Դ�����"
    'Lur = Mark1_Talk("LUR")
    MyExcel.cells(30 - 2, 2).Value = Lur
    
    Frmrepot.Label1.Caption = "����ͳ��λ�ø��³ɹ����� ..."
    Frmrepot.Label1.Refresh
    
    MyExcel.cells(31 - 2, 1).Value = "λ�ø��³ɹ�������"
    'Lua = Mark1_Talk("LUA")
    MyExcel.cells(31 - 2, 2).Value = Lua
    
    Frmrepot.Label1.Caption = "����ͳ��λ�ø���ʧ�ܴ��� ..."
    Frmrepot.Label1.Refresh
    
    MyExcel.cells(32 - 2, 1).Value = "λ�ø���ʧ�ܴ�����"
    'Luf = Mark1_Talk("LUF")
    MyExcel.cells(32 - 2, 2).Value = Luf
    
    If Lur < Lua + Luf Then
        Lur = Lua + Luf
        MyExcel.cells(30 - 2, 2).Value = Lur
    End If
    
    HoASFLFlag = False
    
    MyExcel.cells(33 - 2, 1).Value = "             ʧ��ԭ�����ͳ�ƣ�"
    MyExcel.cells(34 - 2, 1).Value = "                        1.��ʱ"
    'Luf1 = Mark1_Talk("LUF 1")
    MyExcel.cells(34 - 2, 2).Value = Luf1
    
    
    MyExcel.cells(35 - 2, 1).Value = "                        2.�ܾ�"
    'Luf2 = Mark1_Talk("LUF 2")
    MyExcel.cells(35 - 2, 2).Value = Luf2
    
    
    MyExcel.cells(36 - 2, 1).Value = "λ�ø���ʧ���ʣ�"
    If Lur > 0 Then
       LuaFail = Format(Luf / Lur, "percent")  'λ�ø���ʧ����
    End If
    MyExcel.cells(36 - 2, 2).Value = LuaFail
    Frmrepot.ProgressBar1.Value = 22
    Frmrepot.Label1.Caption = "����ͳ�ƺ����ͷŹ��� ..."
    Frmrepot.Label1.Refresh
    
                CdRelease = 0
                CdDr = 0    '���е���
                CDNoDr = 0  '���е���
                For i = 1 To stre_num
                    Set rst = dbs.OpenRecordset("select mark1 from " & convert_filename(i) & " where left(mark1,2)=""CD""")
                    If rst.RecordCount > 0 Then
                        rst.MoveLast
                        rst.MoveFirst
                        For j = 1 To rst.RecordCount
                            Select Case Left(Trim(rst.Fields("mark1").Value), 5)
                                Case "CD ����"
                                    CdRelease = CdRelease + 1
                                Case "CD ����"
                                    CdDr = CdDr + 1
                                Case "CD ����"
                                    CDNoDr = CDNoDr + 1
                            End Select
                            rst.MoveNext
                        Next
                    End If
                Next
    
    MyExcel.cells(38 - 2, 1).Value = "3.�����ͷŹ���"
    MyExcel.cells(39 - 2, 1).Value = "�����ͷŴ�����"
    Frmrepot.Label1.Caption = "����ͳ�������ͷŴ��� ..."
    Frmrepot.Label1.Refresh
    'CdRelease = Mark1_Talk("CD ����")
    MyExcel.cells(39 - 2, 2).Value = CdRelease
    Frmrepot.Label1.Caption = "����ͳ�Ƶ������� ..."
    Frmrepot.Label1.Refresh
   ' CdDr = Mark1_Talk("CD ����")
   ' CdHandDr = Mark1_Talk("CD �л�")
   ' CDNoDr = Mark1_Talk("CD �޷�")
    
    MyExcel.cells(40 - 2, 1).Value = "����������"
    MyExcel.cells(40 - 2, 2).Value = CdDr + CDNoDr
    
    MyExcel.cells(41 - 2, 1).Value = "             ����ԭ�����ͳ�ƣ�"
    MyExcel.cells(42 - 2, 1).Value = "                        1.���е���"
    MyExcel.cells(42 - 2, 2).Value = CdDr
    MyExcel.cells(43 - 2, 1).Value = "                        2.���е���"
    MyExcel.cells(43 - 2, 2).Value = CDNoDr
    
    MyExcel.cells(44 - 2, 1).Value = "�����ʣ�"
    If setup_n > 0 Then
        CdDropped = Format((CdDr + CDNoDr) / setup_n, "percent")  '���н���ʧ����
    End If
    
    MyExcel.cells(44 - 2, 2).Value = CdDropped
    
    '@@@@@@@@@@@@@@@@@@@@
    
End Sub
