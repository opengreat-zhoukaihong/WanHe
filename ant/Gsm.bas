Attribute VB_Name = "Module3"
Public Back_Sel, Replay_flag, Replay_Time As Integer
Public rmsg1, rmsg2 As String
Public Legend_Tog As Integer
Public Can As Integer
Public tran_f(1 To 50) As String
Public tran_fn, CELL_CCH As Integer
Public tran_del As Integer
Public convert_filename(1 To 50) As String
Public hDbfFile As Integer

Type street
     a As String * 1
     time As String * 12
     frame As String * 10
     lon As String * 12
     lat As String * 12
     message As String * 30
     hex_string As String * 90
     col(1 To 52) As String * 5
     ncell_num As String * 1
End Type
    
Type oldtypeNormal
     a As String * 1
     time As String * 12
     frame As String * 10
     lon As String * 12
     lat As String * 12
     message As String * 30
     hex_string As String * 90
     FieldCol1(1 To 15) As String * 5
     FieldCol2(1 To 10) As String * 3
     FER As String * 5
     SQI As String * 5
     MARK As String * 16
     Cell_2 As String * 5
     FieldCol3(1 To 40) As String * 3
     ncell_num As String * 1
End Type
    
Type typeNormal
     a As String * 1
     time As String * 12
     frame As String * 10
     lon As String * 12
     lat As String * 12
     message As String * 30
     hex_string As String * 90
     FieldCol1(1 To 15) As String * 5
     FieldCol2(1 To 10) As String * 3
     FER As String * 5
     SQI As String * 5
     MARK As String * 16
     Cell_2 As String * 5
     FieldCol3(1 To 40) As String * 3
     ncell_num As String * 1
     NewField(1 To 12) As String * 3
End Type

Type ScanFieldType
    Scan_AR As String * 3
    Scan_RX As String * 2
    Scan_BS As String * 2
End Type

Type typeMarkNormal
     a As String * 1
     time As String * 12
     frame As String * 10
     lon As String * 12
     lat As String * 12
     message As String * 30
     hex_string As String * 90
     FieldCol1(1 To 15) As String * 5
     FieldCol2(1 To 10) As String * 3
     FER As String * 5
     SQI As String * 5
     MARK As String * 30
     MARK1 As String * 60
     MARK2 As String * 20
     Cell_2 As String * 5
     FieldCol3(1 To 40) As String * 3
     ncell_num As String * 1
     NewField(1 To 12) As String * 3
     ScanField(1 To 20) As ScanFieldType
End Type

Type fieldhead
     Name As String * 11
     Type As String * 1
     off As Long
     len As Byte
     Dec As Byte
     filler2 As String * 13
     has_tag As String * 1
End Type

Type ScanField
     start As String * 1
     londbf As String * 12
     latdbf As String * 12
     timedbf As String * 11
End Type

Type dbfhead
     ver As Byte
     year As Byte
     month As Byte
     day As Byte
     recordno As Long
     header_len As Integer
     record_len As Integer
     Zero As String * 20
 End Type

Type np
     Name As String * 30
     path As String * 150
End Type

Function FindArfcn(ByVal ci As String, ARFCN As String) As String
    On Error Resume Next
    i = 0
    ci = Trim(ci)
    row = Val(mapinfo.eval("tableinfo(cell,8)"))
    mapinfo.do "fetch first from cell"
    msg = mapinfo.eval("cell.ci")
    While i < row And msg <> ci
         mapinfo.do "fetch next from cell"
         msg = mapinfo.eval("cell.ci")
         i = i + 1
    Wend
    FindArfcn = mapinfo.eval("cell.cell_name")
    ARFCN = mapinfo.eval("cell.arfcn")
End Function


Function Findcell(ByVal ci As String) As String
    Dim i As Integer
    Dim CellRow As Integer
    Dim TempRow As Integer
    
    On Error Resume Next
    mapinfo.do "fetch first from cell"
    mapinfo.do "select * from cell where ci = " + Chr(34) + Trim(ci) + Chr(34) + " into temp"
    TempRow = Val(mapinfo.eval("tableinfo(temp,8)"))
    If TempRow > 0 Then
       Findcell = mapinfo.eval("temp.cell_name")
    Else
       Findcell = ""
    End If
    mapinfo.do "close table temp"
End Function

Function Find_name(ByVal AA As String, ByVal bb As String) As String
    Dim bs, arf As Integer
    On Error Resume Next
    i = 0
    row = Val(mapinfo.eval("tableinfo(cell,8)"))
    mapinfo.do "fetch first from cell"
    msg1 = mapinfo.eval("cell.bsic")
    msg2 = mapinfo.eval("cell.arfcn")
    While (i < row) And (msg1 <> AA Or msg2 <> bb)
         msg1 = mapinfo.eval("cell.bsic")
         msg2 = mapinfo.eval("cell.arfcn")
         mapinfo.do "fetch next from cell"
         i = i + 1
    Wend
    Find_name = mapinfo.eval("cell.cell_name")

End Function

Function Find_id(ByVal AA As String, ByVal bb As String) As String
    Dim bs, arf As Integer
    On Error Resume Next
    i = 0
    row = Val(mapinfo.eval("tableinfo(cell,8)"))
    mapinfo.do "fetch first from cell"
    msg = mapinfo.eval("cell.bsic")
    msg1 = mapinfo.eval("cell.arfcn")
    While (i < row) And (msg <> AA) And (msg1 <> bb)
         msg = mapinfo.eval("cell.bsic")
         msg1 = mapinfo.eval("cell.arfcn")
         mapinfo.do "fetch next from cell"
         i = i + 1
    Wend
    Find_id = mapinfo.eval("cell.ci")

End Function

Function FindNArfcn(ByVal ci As String, arf As String, bs_no As String) As String
    Dim SelectName As String
    
    On Error Resume Next
    ci = Trim(ci)
    mapinfo.do "select * from cell where ci = " + Chr(34) + ci + Chr(34) + " into seltemp"
    If Val(mapinfo.eval("tableinfo(seltemp,8)")) > 0 Then
    
    'row = Val(mapinfo.eval("tableinfo(cell,8)"))
    'mapinfo.do "fetch first from cell"
    'msg = mapinfo.eval("cell.ci")
    'While i < row And msg <> ci
    '     mapinfo.do "fetch next from cell"
    '     msg = mapinfo.eval("cell.ci")
    '     i = i + 1
    'Wend
       SelectName = mapinfo.eval("seltemp.cell_name")
       If InStr(SelectName, Chr(0)) > 0 Then
          FindNArfcn = Left(SelectName, InStr(SelectName, Chr(0)) - 1)
       Else
          FindNArfcn = SelectName
       End If
       arf = mapinfo.eval("seltemp.ARFCN")
       bs_no = mapinfo.eval("seltemp.bs_no")
    Else
       FindNArfcn = ""
       arf = ""
       bs_no = ""
    End If
End Function

Sub MakeNormalFile()
    Dim NormalHeadData As ScanHead
    Dim MyField As WriteField
    Dim i As Integer
    Dim ReturnStr As String * 1
    Dim MyPos As Long
    
    On Error Resume Next
    NormalHeadData.ver = 3
    NormalHeadData.year = Val(Right(Format(year(Now)), 2))
    NormalHeadData.month = month(Now)
    NormalHeadData.day = day(Now)
    NormalHeadData.recordno = 0
    NormalHeadData.HeaderLen = 2849
    NormalHeadData.RecordLen = 460
    NormalHeadData.Zero = String(20, Chr(0))
    Put #hDbfFile, , NormalHeadData
      
    Call PutFieldDef("TIME", "C", 1, 12, 0)
    Call PutFieldDef("NUM_FRAME", "C", 13, 10, 0)
    Call PutFieldDef("LON", "N", 23, 12, 6)
    Call PutFieldDef("LAT", "N", 35, 12, 6)
    Call PutFieldDef("MESSAGE", "C", 47, 30, 0)
    Call PutFieldDef("HEX_STRING", "C", 77, 90, 0)
    Call PutFieldDef("NUM_DCH", "C", 167, 5, 0)
    Call PutFieldDef("TN_DCH", "C", 172, 5, 0)
    Call PutFieldDef("TYPE_DCH", "C", 177, 5, 0)
    Call PutFieldDef("MODE_DCH", "C", 182, 5, 0)
    Call PutFieldDef("NUM_S_DCH", "C", 187, 5, 0)
    Call PutFieldDef("HOPPING", "C", 192, 5, 0)
    Call PutFieldDef("MAIO_DCH", "C", 197, 5, 0)
    Call PutFieldDef("HSN_DCH_", "C", 202, 5, 0)
    Call PutFieldDef("CELL_SERV", "C", 207, 5, 0)
    Call PutFieldDef("CI_SERV", "C", 212, 5, 0)
    Call PutFieldDef("BSIC_SERV", "N", 217, 5, 0)
    Call PutFieldDef("BCCH_SERV", "N", 222, 5, 0)
    Call PutFieldDef("MCC_SERV", "C", 227, 5, 0)
    Call PutFieldDef("MNC_SERV", "C", 232, 5, 0)
    Call PutFieldDef("LAC_SERV", "C", 237, 5, 0)
    Call PutFieldDef("RXLEV_F", "N", 242, 3, 0)
    Call PutFieldDef("RXQUAL_F", "C", 245, 3, 0)
    Call PutFieldDef("RXLEV_S", "N", 248, 3, 0)
    Call PutFieldDef("RXQUAL_S", "C", 251, 3, 0)
    Call PutFieldDef("TA", "C", 254, 3, 0)
    Call PutFieldDef("TX_POWER", "C", 257, 3, 0)
    Call PutFieldDef("ACT_RLINK", "C", 260, 3, 0)
    Call PutFieldDef("MAX_RLINK", "C", 263, 3, 0)
    Call PutFieldDef("C1", "C", 266, 3, 0)
    Call PutFieldDef("C2", "C", 269, 3, 0)
    Call PutFieldDef("FER", "C", 272, 5, 0)
    Call PutFieldDef("SQI", "C", 277, 5, 0)
    Call PutFieldDef("MARK", "C", 282, 16, 0)
    Call PutFieldDef("CELL_2", "C", 298, 5, 0)
    Call PutFieldDef("BSIC_2", "C", 303, 3, 0)
    Call PutFieldDef("ARFCN_2", "C", 306, 3, 0)
    Call PutFieldDef("RXLEV_F_2", "N", 309, 3, 0)
    Call PutFieldDef("RXQUQL_F_2", "C", 312, 3, 0)
    Call PutFieldDef("RXLEV_S_2", "N", 315, 3, 0)
    Call PutFieldDef("RXQUQL_S_2", "C", 318, 3, 0)
    Call PutFieldDef("TA_2", "C", 321, 3, 0)
    Call PutFieldDef("TX_POWER_2", "C", 324, 3, 0)
    Call PutFieldDef("ACT_RLINK2", "C", 327, 3, 0)
    Call PutFieldDef("DTX", "C", 330, 3, 0)
    MyPos = 333
    For i = 1 To 6
        Call PutFieldDef("BCCH_N" & Format(i), "N", MyPos, 3, 0)
        Call PutFieldDef("RXLEV_N" & Format(i), "N", MyPos + 3, 3, 0)
        Call PutFieldDef("BSIC_N" & Format(i), "N", MyPos + 6, 3, 0)
        Call PutFieldDef("C1_N" & Format(i), "C", MyPos + 9, 3, 0)
        Call PutFieldDef("C2_N" & Format(i), "C", MyPos + 12, 3, 0)
        MyPos = MyPos + 15
    Next
    Call PutFieldDef("NCELL_NUM", "N", MyPos, 1, 0)
    Call PutFieldDef("RXLE_SAME1", "N", MyPos + 1, 3, 0)
    Call PutFieldDef("BSIC_SAME1", "N", MyPos + 3, 3, 0)
    Call PutFieldDef("RXLE_SAME2", "N", MyPos + 3, 3, 0)
    Call PutFieldDef("BSIC_SAME2", "N", MyPos + 3, 3, 0)
    Call PutFieldDef("RXLE_NEIG1", "N", MyPos + 3, 3, 0)
    Call PutFieldDef("BSIC_NEIG1", "N", MyPos + 3, 3, 0)
    Call PutFieldDef("RXLE_NEIG2", "N", MyPos + 3, 3, 0)
    Call PutFieldDef("BSIC_NEIG2", "N", MyPos + 3, 3, 0)
    Call PutFieldDef("RXLE_NEIG3", "N", MyPos + 3, 3, 0)
    Call PutFieldDef("BSIC_NEIG3", "N", MyPos + 3, 3, 0)
    Call PutFieldDef("RXLE_NEIG4", "N", MyPos + 3, 3, 0)
    Call PutFieldDef("BSIC_NEIG4", "N", MyPos + 3, 3, 0)
    ReturnStr = Chr(13)
    Put #hDbfFile, , ReturnStr
    
End Sub

Sub PutFieldDef(FieldName As String, FieldType As String, FieldPos As Long, FieldLength As Byte, FieldDec As Byte)
    Dim MyField As WriteField
    
    On Error Resume Next
    MyField.Name = FieldName + String(11 - Len(FieldName), Chr(0))
    MyField.Type = FieldType
    MyField.Pos = FieldPos
    MyField.length = FieldLength
    MyField.Dec = FieldDec
    MyField.Zero = String(14, Chr(0))
    Put #hDbfFile, , MyField
    
End Sub

Sub MakeCellFile()
    Dim NormalHeadData As ScanHead
    Dim MyField As WriteField
    Dim i As Integer
    Dim ReturnStr As String * 1
    Dim MyPos As Long
    
    On Error Resume Next
    NormalHeadData.ver = 3
    NormalHeadData.year = Val(Right(Format(year(Now)), 2))
    NormalHeadData.month = month(Now)
    NormalHeadData.day = day(Now)
    NormalHeadData.recordno = 0
    NormalHeadData.HeaderLen = 1025
    NormalHeadData.RecordLen = 309
    NormalHeadData.Zero = String(20, Chr(0))
    Put #hDbfFile, , NormalHeadData
      
    Call PutFieldDef("CELL_NAME", "C", 1, 15, 0)
    Call PutFieldDef("BS_NO", "C", 16, 10, 0)
    Call PutFieldDef("CI", "C", 26, 5, 0)
    Call PutFieldDef("ARFCN", "N", 31, 3, 0)
    Call PutFieldDef("BSIC", "N", 34, 3, 0)
    Call PutFieldDef("BEARING", "N", 37, 3, 0)
    Call PutFieldDef("LAC", "N", 40, 5, 0)
    Call PutFieldDef("NON_BCCH", "C", 45, 64, 0)
    Call PutFieldDef("DOWNTILT", "N", 109, 3, 0)
    Call PutFieldDef("MAX_TX_BTS", "N", 112, 2, 0)
    Call PutFieldDef("MAX_TX_MS", "N", 114, 2, 0)
    Call PutFieldDef("TIME", "C", 116, 8, 0)
    Call PutFieldDef("LON", "N", 124, 12, 6)
    Call PutFieldDef("LAT", "N", 136, 12, 6)
    Call PutFieldDef("MICROCELL", "C", 148, 1, 0)
    
    MyPos = 149
    For i = 1 To 16
        Call PutFieldDef("NCELL" & Format(i), "C", MyPos, 10, 0)
        MyPos = MyPos + 10
    Next
    
    ReturnStr = Chr(13)
    Put #hDbfFile, , ReturnStr
    
End Sub

Sub MakeCell1800File()
    Dim NormalHeadData As ScanHead
    Dim MyField As WriteField
    Dim i As Integer
    Dim ReturnStr As String * 1
    Dim MyPos As Long
    
    On Error Resume Next
    NormalHeadData.ver = 3
    NormalHeadData.year = Val(Right(Format(year(Now)), 2))
    NormalHeadData.month = month(Now)
    NormalHeadData.day = day(Now)
    NormalHeadData.recordno = 0
    NormalHeadData.HeaderLen = (35 + 1) * 32 + 1
    NormalHeadData.RecordLen = 336 + 5
    NormalHeadData.Zero = String(20, Chr(0))
    Put #hDbfFile, , NormalHeadData
    
    Call PutFieldDef("CELL_NAME", "C", 1, 21, 0)
    Call PutFieldDef("BS_NO", "C", 22, 10, 0)
    Call PutFieldDef("CI", "C", 32, 5, 0)
    Call PutFieldDef("ARFCN", "N", 37, 3, 0)
    Call PutFieldDef("BSIC", "N", 40, 3, 0)
    Call PutFieldDef("BEARING", "N", 43, 3, 0)
    Call PutFieldDef("LAC", "N", 46, 5, 0)
    Call PutFieldDef("NON_BCCH", "C", 51, 64, 0)
    Call PutFieldDef("DOWNTILT", "N", 115, 3, 0)
    'Call PutFieldDef("MAX_TX_BTS", "N", 118, 2, 0)
    Call PutFieldDef("MAX_TX_BTS", "C", 118, 2, 0)
    Call PutFieldDef("ANT_HEIGH", "C", 120, 3, 0)
    Call PutFieldDef("MAX_TX_MS", "C", 123, 2, 0)
    Call PutFieldDef("ANT_GAIN", "C", 125, 3, 0)
    Call PutFieldDef("ANT_TYPE", "C", 128, 15, 0)
    Call PutFieldDef("TIME", "C", 143, 8, 0)
    Call PutFieldDef("LON", "N", 151, 12, 6)
    Call PutFieldDef("LAT", "N", 163, 12, 6)
    'Call PutFieldDef("MICROCELL", "C", 175, 1, 0)
    'Call PutFieldDef("DCSBASE", "C", 176, 1, 0)
    Call PutFieldDef("BASETYPE", "C", 175, 1, 0)
    
    Call PutFieldDef("LENGTH", "C", 176, 5, 0)
    
    MyPos = 149 + 7 + 20 + 5
    For i = 1 To 16
        Call PutFieldDef("NCELL" & Format(i), "C", MyPos, 10, 0)
        MyPos = MyPos + 10
    Next
    
    ReturnStr = Chr(13)
    Put #hDbfFile, , ReturnStr
    
End Sub

Function funcCreateCell(filename As String) As Integer
    Dim hfile As Integer
    
    On Error Resume Next
    If Dir(filename, 0) <> "" Then
       hfile = 0
    Else
       hfile = FreeFile
       Open filename For Binary As #hfile
       hDbfFile = hfile
       MakeCell1800File
    End If
    funcCreateCell = hfile

End Function

Function GetBaseName(MyCellName As String)
    Dim mychar As String
    Dim mycode As Integer, finds As Integer
    
    On Error Resume Next
    finds = InStr(MyCellName, Chr(0))
    If finds > 0 Then
       MyCellName = Left(MyCellName, finds - 1)
    End If
    MyCellName = Trim(MyCellName)
    If Len(MyCellName) > 0 Then
       mychar = Right(MyCellName, 1)
       mycode = Asc(mychar)
       'If mycode >= 65 And mycode <= 90 Or mycode >= 97 And mycode <= 122 Or mycode >= 48 And mycode <= 57 Then
       If mycode >= 48 And mycode <= 57 Then
          MyCellName = Left(MyCellName, Len(MyCellName) - 1)
          MyCellName = Trim(MyCellName)
       End If
    End If
    GetBaseName = MyCellName
End Function


Sub MakeBase1800File()
    Dim NormalHeadData As ScanHead
    Dim MyField As WriteField
    Dim i As Integer
    Dim ReturnStr As String * 1
    Dim MyPos As Long
    
    On Error Resume Next
    NormalHeadData.ver = 3
    NormalHeadData.year = Val(Right(Format(year(Now)), 2))
    NormalHeadData.month = month(Now)
    NormalHeadData.day = day(Now)
    NormalHeadData.recordno = 0
    NormalHeadData.HeaderLen = (21 + 1) * 32 + 1
    NormalHeadData.RecordLen = 126
    NormalHeadData.Zero = String(20, Chr(0))
    Put #hDbfFile, , NormalHeadData
    
    Call PutFieldDef("BS_NAME", "C", 1, 20, 0)
    Call PutFieldDef("BS_NO", "C", 21, 10, 0)
    Call PutFieldDef("BCCH_1", "N", 31, 3, 0)
    Call PutFieldDef("BCCH_2", "N", 34, 3, 0)
    Call PutFieldDef("BCCH_3", "N", 37, 3, 0)
    Call PutFieldDef("CI_1", "C", 40, 5, 0)
    Call PutFieldDef("CI_2", "C", 45, 5, 0)
    Call PutFieldDef("CI_3", "C", 50, 5, 0)
    Call PutFieldDef("BSIC_1", "N", 55, 3, 0)
    Call PutFieldDef("BSIC_2", "N", 58, 3, 0)
    Call PutFieldDef("BSIC_3", "N", 61, 3, 0)
    Call PutFieldDef("BEARING_1", "N", 64, 3, 0)
    Call PutFieldDef("BEARING_2", "N", 67, 3, 0)
    Call PutFieldDef("BEARING_3", "N", 70, 3, 0)
    Call PutFieldDef("LAC", "N", 73, 5, 0)
    Call PutFieldDef("BSC__SYSGE", "C", 78, 7, 0)
    Call PutFieldDef("BASE_TYPE", "C", 85, 5, 0)
    Call PutFieldDef("BTS_TYPE", "C", 90, 9, 0)
    Call PutFieldDef("POWER_TYPE", "C", 99, 3, 0)
    Call PutFieldDef("LON", "N", 102, 12, 6)
    Call PutFieldDef("LAT", "N", 114, 12, 6)
    
    ReturnStr = Chr(13)
    Put #hDbfFile, , ReturnStr
    
End Sub

Sub MakeNormalMarkFile()
    Dim NormalHeadData As ScanHead
    Dim MyField As WriteField
    Dim i As Integer
    Dim ReturnStr As String * 1
    Dim MyPos As Long
    
    On Error Resume Next
    NormalHeadData.ver = 3
    NormalHeadData.year = Val(Right(Format(year(Now)), 2))
    NormalHeadData.month = month(Now)
    NormalHeadData.day = day(Now)
    NormalHeadData.recordno = 0
    NormalHeadData.HeaderLen = 151 * 32 + 1 '2849
    NormalHeadData.RecordLen = 460 + 204 + 30
    NormalHeadData.Zero = String(20, Chr(0))
    Put #hDbfFile, , NormalHeadData
      
    Call PutFieldDef("TIME", "C", 1, 12, 0)
    Call PutFieldDef("NUM_FRAME", "C", 13, 10, 0)
    Call PutFieldDef("LON", "N", 23, 12, 6)
    Call PutFieldDef("LAT", "N", 35, 12, 6)
    Call PutFieldDef("MESSAGE", "C", 47, 30, 0)
    Call PutFieldDef("HEX_STRING", "C", 77, 90, 0)
    Call PutFieldDef("NUM_DCH", "C", 167, 5, 0)
    Call PutFieldDef("TN_DCH", "C", 172, 5, 0)
    Call PutFieldDef("TYPE_DCH", "C", 177, 5, 0)
    Call PutFieldDef("MODE_DCH", "C", 182, 5, 0)
    Call PutFieldDef("NUM_S_DCH", "C", 187, 5, 0)
    Call PutFieldDef("HOPPING", "C", 192, 5, 0)
    Call PutFieldDef("MAIO_DCH", "C", 197, 5, 0)
    Call PutFieldDef("HSN_DCH_", "C", 202, 5, 0)
    Call PutFieldDef("CELL_SERV", "C", 207, 5, 0)
    Call PutFieldDef("CI_SERV", "C", 212, 5, 0)
    Call PutFieldDef("BSIC_SERV", "N", 217, 5, 0)
    Call PutFieldDef("BCCH_SERV", "N", 222, 5, 0)
    Call PutFieldDef("MCC_SERV", "C", 227, 5, 0)
    Call PutFieldDef("MNC_SERV", "C", 232, 5, 0)
    Call PutFieldDef("LAC_SERV", "C", 237, 5, 0)
    Call PutFieldDef("RXLEV_F", "N", 242, 3, 0)
    Call PutFieldDef("RXQUAL_F", "C", 245, 3, 0)
    Call PutFieldDef("RXLEV_S", "N", 248, 3, 0)
    Call PutFieldDef("RXQUAL_S", "C", 251, 3, 0)
    Call PutFieldDef("TA", "C", 254, 3, 0)
    Call PutFieldDef("TX_POWER", "C", 257, 3, 0)
    Call PutFieldDef("ACT_RLINK", "C", 260, 3, 0)
    Call PutFieldDef("MAX_RLINK", "C", 263, 3, 0)
    Call PutFieldDef("C1", "C", 266, 3, 0)
    Call PutFieldDef("C2", "C", 269, 3, 0)
    Call PutFieldDef("FER", "C", 272, 5, 0)
    Call PutFieldDef("SQI", "C", 277, 5, 0)
    Call PutFieldDef("MARK", "C", 282, 30, 0)
    Call PutFieldDef("MARK1", "C", 312, 60, 0)
    Call PutFieldDef("MARK2", "C", 372, 20, 0)
    Call PutFieldDef("CELL_2", "C", 392, 5, 0)
    Call PutFieldDef("BSIC_2", "C", 397, 3, 0)
    Call PutFieldDef("ARFCN_2", "C", 400, 3, 0)
    Call PutFieldDef("RXLEV_F_2", "N", 403, 3, 0)
    Call PutFieldDef("RXQUQL_F_2", "C", 406, 3, 0)
    Call PutFieldDef("RXLEV_S_2", "N", 409, 3, 0)
    Call PutFieldDef("RXQUQL_S_2", "C", 412, 3, 0)
    Call PutFieldDef("TA_2", "C", 415, 3, 0)
    Call PutFieldDef("TX_POWER_2", "C", 418, 3, 0)
    Call PutFieldDef("ACT_RLINK2", "C", 421, 3, 0)
    Call PutFieldDef("DTX", "C", 424, 3, 0)
    MyPos = 427
    For i = 1 To 6
        Call PutFieldDef("BCCH_N" & Format(i), "N", MyPos, 3, 0)
        Call PutFieldDef("RXLEV_N" & Format(i), "N", MyPos + 3, 3, 0)
        Call PutFieldDef("BSIC_N" & Format(i), "N", MyPos + 6, 3, 0)
        Call PutFieldDef("C1_N" & Format(i), "C", MyPos + 9, 3, 0)
        Call PutFieldDef("C2_N" & Format(i), "C", MyPos + 12, 3, 0)
        MyPos = MyPos + 15
    Next
    Call PutFieldDef("NCELL_NUM", "N", MyPos, 1, 0)    '900 ncell
    Call PutFieldDef("RXLE_SAME1", "N", MyPos + 1, 3, 0)
    Call PutFieldDef("BSIC_SAME1", "N", MyPos + 4, 3, 0)
    Call PutFieldDef("RXLE_SAME2", "N", MyPos + 7, 3, 0)
    Call PutFieldDef("BSIC_SAME2", "N", MyPos + 10, 3, 0)
    Call PutFieldDef("RXLE_NEIG1", "N", MyPos + 13, 3, 0)
    Call PutFieldDef("BSIC_NEIG1", "N", MyPos + 16, 3, 0)
    Call PutFieldDef("RXLE_NEIG2", "N", MyPos + 19, 3, 0)
    Call PutFieldDef("BSIC_NEIG2", "N", MyPos + 22, 3, 0)
    Call PutFieldDef("RXLE_NEIG3", "N", MyPos + 25, 3, 0)
    Call PutFieldDef("BSIC_NEIG3", "N", MyPos + 28, 3, 0)
    Call PutFieldDef("RXLE_NEIG4", "N", MyPos + 31, 3, 0)
    Call PutFieldDef("BSIC_NEIG4", "N", MyPos + 34, 3, 0)
    
    Call PutFieldDef("SCAN_AR1", "N", MyPos + 37, 3, 0)
    Call PutFieldDef("Scan_RX1", "N", MyPos + 40, 2, 0)
    Call PutFieldDef("Scan_BS1", "N", MyPos + 42, 2, 0)
    Call PutFieldDef("Scan_AR2", "N", MyPos + 44, 3, 0)
    Call PutFieldDef("Scan_RX2", "N", MyPos + 47, 2, 0)
    Call PutFieldDef("Scan_BS2", "N", MyPos + 49, 2, 0)
    Call PutFieldDef("Scan_AR3", "N", MyPos + 51, 3, 0)
    Call PutFieldDef("Scan_RX3", "N", MyPos + 54, 2, 0)
    Call PutFieldDef("Scan_BS3", "N", MyPos + 56, 2, 0)
    Call PutFieldDef("Scan_AR4", "N", MyPos + 58, 3, 0)
    Call PutFieldDef("Scan_RX4", "N", MyPos + 61, 2, 0)
    Call PutFieldDef("Scan_BS4", "N", MyPos + 63, 2, 0)
    Call PutFieldDef("Scan_AR5", "N", MyPos + 65, 3, 0)
    Call PutFieldDef("Scan_RX5", "N", MyPos + 68, 2, 0)
    Call PutFieldDef("Scan_BS5", "N", MyPos + 70, 2, 0)
    Call PutFieldDef("Scan_AR6", "N", MyPos + 73, 3, 0)
    Call PutFieldDef("Scan_RX6", "N", MyPos + 75, 2, 0)
    Call PutFieldDef("Scan_BS6", "N", MyPos + 77, 2, 0)
    Call PutFieldDef("Scan_AR7", "N", MyPos + 80, 3, 0)
    Call PutFieldDef("Scan_RX7", "N", MyPos + 82, 2, 0)
    Call PutFieldDef("Scan_BS7", "N", MyPos + 84, 2, 0)
    Call PutFieldDef("Scan_AR8", "N", MyPos + 86, 3, 0)
    Call PutFieldDef("Scan_RX8", "N", MyPos + 89, 2, 0)
    Call PutFieldDef("Scan_BS8", "N", MyPos + 91, 2, 0)
    Call PutFieldDef("Scan_AR9", "N", MyPos + 93, 3, 0)
    Call PutFieldDef("Scan_RX9", "N", MyPos + 96, 2, 0)
    Call PutFieldDef("Scan_BS9", "N", MyPos + 98, 2, 0)
    Call PutFieldDef("Scan_AR10", "N", MyPos + 100, 3, 0)
    Call PutFieldDef("Scan_RX10", "N", MyPos + 103, 2, 0)
    Call PutFieldDef("Scan_BS10", "N", MyPos + 105, 2, 0)
    Call PutFieldDef("Scan_AR11", "N", MyPos + 107, 3, 0)
    Call PutFieldDef("Scan_RX11", "N", MyPos + 110, 2, 0)
    Call PutFieldDef("Scan_BS11", "N", MyPos + 112, 2, 0)
    Call PutFieldDef("Scan_AR12", "N", MyPos + 114, 3, 0)
    Call PutFieldDef("Scan_RX12", "N", MyPos + 117, 2, 0)
    Call PutFieldDef("Scan_BS12", "N", MyPos + 119, 2, 0)
    Call PutFieldDef("Scan_AR13", "N", MyPos + 121, 3, 0)
    Call PutFieldDef("Scan_RX13", "N", MyPos + 124, 2, 0)
    Call PutFieldDef("Scan_BS13", "N", MyPos + 126, 2, 0)
    Call PutFieldDef("Scan_AR14", "N", MyPos + 128, 3, 0)
    Call PutFieldDef("Scan_RX14", "N", MyPos + 131, 2, 0)
    Call PutFieldDef("Scan_BS14", "N", MyPos + 133, 2, 0)
    Call PutFieldDef("Scan_AR15", "N", MyPos + 135, 3, 0)
    Call PutFieldDef("Scan_RX15", "N", MyPos + 138, 2, 0)
    Call PutFieldDef("Scan_BS15", "N", MyPos + 140, 2, 0)
    Call PutFieldDef("Scan_AR16", "N", MyPos + 142, 3, 0)
    Call PutFieldDef("Scan_RX16", "N", MyPos + 145, 2, 0)
    Call PutFieldDef("Scan_BS16", "N", MyPos + 147, 2, 0)
    Call PutFieldDef("Scan_AR17", "N", MyPos + 149, 3, 0)
    Call PutFieldDef("Scan_RX17", "N", MyPos + 152, 2, 0)
    Call PutFieldDef("Scan_BS17", "N", MyPos + 154, 2, 0)
    Call PutFieldDef("Scan_AR18", "N", MyPos + 156, 3, 0)
    Call PutFieldDef("Scan_RX18", "N", MyPos + 159, 2, 0)
    Call PutFieldDef("Scan_BS18", "N", MyPos + 161, 2, 0)
    Call PutFieldDef("Scan_AR19", "N", MyPos + 163, 3, 0)
    Call PutFieldDef("Scan_RX19", "N", MyPos + 166, 2, 0)
    Call PutFieldDef("Scan_BS19", "N", MyPos + 168, 2, 0)
    Call PutFieldDef("Scan_AR20", "N", MyPos + 170, 3, 0)
    Call PutFieldDef("Scan_RX20", "N", MyPos + 173, 2, 0)
    Call PutFieldDef("Scan_BS20", "N", MyPos + 175, 2, 0)
    
    ReturnStr = Chr(13)
    Put #hDbfFile, , ReturnStr
End Sub

Sub SearchCellName(MyBsic, Mybcch, MyLon, MyLat, MyCellName, MyCi, MyServingCell)
    Dim MyRows As Integer
    Dim i As Integer, j As Integer
    Dim Distance As Variant
    Dim CurrentLon As Variant, CurrentLat As Variant
    Dim Mytemp As String
    
    On Error Resume Next
    If MyCi <> "" Then
       mapinfo.do "select * from cell where ci = " & Chr(34) & MyCi & Chr(34) & " into Mytemp"
       MyRows = mapinfo.eval("tableinfo(mytemp,8)")
       If MyRows = 0 Then
          MyCellName = ""
          mapinfo.do "x1=0"
          mapinfo.do "y1=0"
          mapinfo.do "x3=0"
       Else
          MyCellName = mapinfo.eval("mytemp.cell_name")
          mapinfo.do "x1=mytemp.lon"
          mapinfo.do "y1=mytemp.lat"
          mapinfo.do "x3=mytemp.bearing"
       End If
       If InStr(MyCellName, Chr(0)) > 0 Then
          MyCellName = Trim(Left(MyCellName, InStr(MyCellName, Chr(0)) - 1))
       End If
    Else
       mapinfo.do "select * from cell where bsic = " & MyBsic & " and arfcn = " & Mybcch & " into Mytemp"
       MyRows = mapinfo.eval("tableinfo(mytemp,8)")
       If MyRows = 0 Then
          MyCellName = ""
          mapinfo.do "x1=0"
          mapinfo.do "y1=0"
          mapinfo.do "x3=0"
       Else
          If MyRows > 1 Then
             mapinfo.do "fetch first from mytemp"
             CurrentLon = mapinfo.eval("mytemp.lon")
             CurrentLat = mapinfo.eval("mytemp.lat")
             Distance = (MyLon - CurrentLon) ^ 2 + (MyLat - CurrentLat) ^ 2
             MyCellName = mapinfo.eval("mytemp.cell_name")
            mapinfo.do "x1=mytemp.lon"
            mapinfo.do "y1=mytemp.lat"
            mapinfo.do "x3=mytemp.bearing"
             j = 2
                If InStr(MyCellName, Chr(0)) > 0 Then
                   MyCellName = Trim(Left(MyCellName, InStr(MyCellName, Chr(0)) - 1))
                End If
                If MyCellName = MyServingCell Then
                   
                    mapinfo.do "fetch next from mytemp"
                    j = j + 1
                    CurrentLon = mapinfo.eval("mytemp.lon")
                    CurrentLat = mapinfo.eval("mytemp.lat")
                    Distance = (MyLon - CurrentLon) ^ 2 + (MyLat - CurrentLat) ^ 2
                    MyCellName = mapinfo.eval("mytemp.cell_name")
                    mapinfo.do "x1=mytemp.lon"
                    mapinfo.do "y1=mytemp.lat"
                    mapinfo.do "x3=mytemp.bearing"
                End If

             mapinfo.do "fetch next from mytemp"

             For i = j To MyRows
                CurrentLon = mapinfo.eval("mytemp.lon")
                CurrentLat = mapinfo.eval("mytemp.lat")
                 If (MyLon - CurrentLon) ^ 2 + (MyLat - CurrentLat) ^ 2 < Distance Then
                    Mytemp = mapinfo.eval("mytemp.cell_name")
                    If InStr(MyCellName, Chr(0)) > 0 Then
                       Mytemp = Trim(Left(MyCellName, InStr(MyCellName, Chr(0)) - 1))
                    End If
                    If Mytemp <> MyServingCell Then
                        Distance = (MyLon - CurrentLon) ^ 2 + (MyLat - CurrentLat) ^ 2
                        MyCellName = mapinfo.eval("mytemp.cell_name")
                        mapinfo.do "x1=mytemp.lon"
                        mapinfo.do "y1=mytemp.lat"
                        mapinfo.do "x3=mytemp.bearing"
                    End If
                 End If
                 mapinfo.do "fetch next from mytemp"
             Next
          Else
             MyCellName = mapinfo.eval("mytemp.cell_name")
            mapinfo.do "x1=mytemp.lon"
            mapinfo.do "y1=mytemp.lat"
            mapinfo.do "x3=mytemp.bearing"
          End If
          If InStr(MyCellName, Chr(0)) > 0 Then
             MyCellName = Trim(Left(MyCellName, InStr(MyCellName, Chr(0)) - 1))
          End If
       End If
    End If
    mapinfo.do "close table mytemp"
End Sub

