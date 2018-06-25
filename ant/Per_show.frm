VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Per_Show 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "数据转换"
   ClientHeight    =   1515
   ClientLeft      =   765
   ClientTop       =   6975
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Per_show.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1515
   ScaleWidth      =   4845
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   320
      Left            =   1935
      TabIndex        =   2
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3795
      Top             =   90
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   315
      TabIndex        =   1
      Top             =   750
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "正在转换"
      Height          =   180
      Left            =   315
      TabIndex        =   0
      Top             =   315
      Width           =   720
   End
End
Attribute VB_Name = "Per_Show"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As field
Dim field_1 As ScanField
Dim dbf_field As fieldhead
Dim dbf_head As dbfhead
Dim s_data As typeMarkNormal
Dim Tems_Data As typeNormal
Dim Convert_Flag As Integer

Function conv_gsm(txtname, dbfname) As Integer     'Tems数据不采用新数据格式。
    Dim fillword As String * 77
    Dim fill As Integer, recordno As Integer
    Dim dbfhead As String * 1919
    Dim sword As String * 2
    Dim linetxt As String, wtab As String
    Dim findtab, findN, lenline, l As Integer
    Dim col As String, lat As String, lon As String
    Dim spac As String * 1, endd As String * 1
    Dim res As Double
    Dim dtx As String
    Dim val_dtx
    Dim old_lon As String * 12, old_lat As String * 12
    Dim DelFlag As Boolean
    Dim MyValueTemp As String
    
    On Error Resume Next
    ncell_num = 0
    old_lon = space$(12)
    old_lat = space$(12)
    Dim lon_no, lat_no As Integer
    Select Case GPS_NO
           Case 1
                lon_no = 8
                lat_no = 7
           Case 2
                lon_no = 9
                lat_no = 8
           Case 3
                lon_no = 10
                lat_no = 9
           Case 4
                lon_no = 11
                lat_no = 10
           Case 5
                lon_no = 12
                lat_no = 11
           Case 6
                lon_no = 13
                lat_no = 12
    End Select
    
    'Gsm_FileName = Gsm_Path + "\data_new.dbf"
    'FileCopy Gsm_FileName, dbfname
    'Lixuhua:
    'Open Gsm_FileName For Binary As #1

    If Dir(dbfname, 0) <> "" Then
       Kill dbfname
    End If
    Open dbfname For Binary As #2
    hDbfFile = 2
    MakeNormalFile
    'Seek #2, 2466 + 384
    'll = FileLen(Gsm_FileName)
    'Get #1, ll, endd
    recordno = 0
    spac = space$(1)
    wtab = Chr$(9)
    eofword = Chr$(26)
    fillword = space$(77)
    i = 0
    j = 1
    Label1.Caption = "正在转换 " + txtname
    Label1.Refresh
    Open txtname For Input As #3
    nline = FileLen(txtname) / 260
    bline = Fix(nline / 100)
    bs = 1
    ProgressBar1.Value = 0
    longline = 0
    'Seek #3, 1061
    For i = 1 To 3
        Line Input #3, linetxt
    Next
    Tems_Data.a = " "
    Tems_Data.FieldCol3(10) = space$(3)
    Do While Not EOF(3)
    
       Tems_Data.time = space$(12)
       Tems_Data.frame = space$(10)
       Tems_Data.lon = space$(12)
       Tems_Data.lat = space$(12)
       Tems_Data.message = space$(30)
       Tems_Data.hex_string = space$(90)
       For i = 1 To 15
           Tems_Data.FieldCol1(i) = space$(5)
       Next
       For i = 1 To 10
           Tems_Data.FieldCol2(i) = space$(3)
       Next
       Tems_Data.FER = space$(5)
       Tems_Data.SQI = space$(5)
       Tems_Data.MARK = space$(16)
       Tems_Data.Cell_2 = space$(5)
       For i = 1 To 40
           If i <> 10 Then
              Tems_Data.FieldCol3(i) = space$(3)
           End If
       Next
       Tems_Data.ncell_num = "0"
       For i = 1 To 12
           Tems_Data.NewField(i) = space$(3)
       Next
       DelFlag = False
       
       Line Input #3, linetxt
       longline = longline + 1
       If longline = bs * bline And ProgressBar1.Value < 99 Then
          ProgressBar1.Value = ProgressBar1.Value + 1
          bs = bs + 1
       End If

       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       Tems_Data.time = Left(linetxt, Abs(findtab - 1))
       linetxt = Right(linetxt, lenline - findtab)
       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       Tems_Data.frame = Left(linetxt, Abs(findtab - 1))
       linetxt = Right(linetxt, lenline - findtab)
    
       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       Tems_Data.lat = Left(linetxt, Abs(findtab - 1))
       l = Len(Tems_Data.lat)
       findN = InStr(Tems_Data.lat, "N")
       Tems_Data.lat = Right(Tems_Data.lat, l - findN)
       Tems_Data.lat = LTrim(Tems_Data.lat)
       findspace = InStr(Tems_Data.lat, " ")
    
       If findspace = lat_no Then
          res = Fix(Val(Tems_Data.lat) / 100) + (Val(Tems_Data.lat) - Fix(Val(Tems_Data.lat) / 100) * 100) / 60
          Tems_Data.lat = Left(res, 9)
          lat = LTrim$(Tems_Data.lat)
          If Len(lat) < 9 Then
             lat = lat + String(9 - Len(lat), "0")
          End If
          Tems_Data.lat = lat
       Else
          If Not Data_Report Then
             GoTo delrecord
          End If
       End If
       linetxt = Right(linetxt, lenline - findtab)
       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       Tems_Data.lon = Left(linetxt, Abs(findtab - 1))
       l = Len(Tems_Data.lon)
       findN = InStr(Tems_Data.lon, "E")
       Tems_Data.lon = Right(Tems_Data.lon, l - findN)
       Tems_Data.lon = LTrim(Tems_Data.lon)
       findspace = InStr(Tems_Data.lon, " ")
       If findspace = lon_no Then
          res = Fix(Val(Tems_Data.lon) / 100) + (Val(Tems_Data.lon) - Fix(Val(Tems_Data.lon) / 100) * 100) / 60
          Tems_Data.lon = Left(res, 10)
          lon = LTrim$(Tems_Data.lon)
          If Len(lon) < 10 Then
             lon = lon + String(10 - Len(lon), "0")
          End If
          Tems_Data.lon = lon
       End If
       If tran_del = 2 Then
          If Tems_Data.lat = old_lat And Tems_Data.lon = old_lon Then
             If Not Data_Report Then
                DelFlag = True
             End If
             'GoTo delrecord
          Else
             old_lat = Tems_Data.lat
             old_lon = Tems_Data.lon
          End If
       Else
          If tran_del = 3 Then
             If (Val(Tems_Data.lat) - Val(old_lat)) * (Val(Tems_Data.lat) - Val(old_lat)) + (Val(Tems_Data.lon) - Val(old_lon)) * (Val(Tems_Data.lon) - Val(old_lon)) < 0.000001 Then
                If Not Data_Report Then
                   GoTo delrecord
                End If
             Else
                old_lat = Tems_Data.lat
                old_lon = Tems_Data.lon
             End If
          End If
       End If
             
       linetxt = Right(linetxt, lenline - findtab)
       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       If findtab > 1 Then
          Tems_Data.message = Left(linetxt, Abs(findtab - 1))
       End If
       If DelFlag Then
          DelFlag = False
          Select Case UCase(Trim(Tems_Data.message))
             Case "ASSIGNMENT COMPLETE", "CONGESTION CONTROL", "NO SERVICE REPORT", "HANDOVER COMPLETE", "HANDOVER FAILURE", "RELEASE"
             Case "HANDOVER COMMAND", "CONNECT", "CHANNEL REQUEST", "SETUP", "ASSIGNMENT COMMAND", "ASSIGNMENT COMPLETE", "DISCONNECT"
             Case "ASSIGNMENT FAILURES", "ALERTING", "LOCATION UPDATING REQUEST", "LOCATION UPDATING ACCEPT", "LOCATION UPDATING REJECT", "CHANNEL RELEASE"
             Case Else
                  GoTo delrecord
          End Select
       End If
       
       linetxt = Right(linetxt, lenline - findtab)
       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       If findtab > 1 Then
          Tems_Data.hex_string = Left(linetxt, Abs(findtab - 1))
       End If
       recordno = recordno + 1
   
       linetxt = Right(linetxt, lenline - findtab)
       lenline = Len(linetxt)
       For i = 1 To 34
          findtab = InStr(linetxt, wtab)
          If findtab = 1 Then
             linetxt = Right(linetxt, lenline - 1)
             lenline = lenline - 1
          Else
             If findtab = 0 Then
                fill = 1
                'Exit Do
                Exit For
             Else
             If i <> 34 Then
                If i >= 1 And i <= 15 Then
                   Tems_Data.FieldCol1(i) = Trim(Left(linetxt, Abs(findtab - 1)))
                ElseIf i >= 16 And i <= 23 Then
                   Tems_Data.FieldCol2(i - 15) = Trim(Left(linetxt, Abs(findtab - 1)))
                ElseIf i = 24 Then
                   Tems_Data.Cell_2 = Trim(Left(linetxt, Abs(findtab - 1)))
                Else
                   Tems_Data.FieldCol3(i - 24) = Trim(Left(linetxt, Abs(findtab - 1)))
                End If
             End If
             linetxt = Right(linetxt, lenline - findtab)
             lenline = Len(linetxt)
             End If
          End If
       Next
       fill = 0
       If Val(Tems_Data.FieldCol2(3)) = 0 Then
          Tems_Data.FieldCol2(3) = Tems_Data.FieldCol2(1)
          Tems_Data.FieldCol2(4) = Tems_Data.FieldCol2(2)
       End If
       Tems_Data.hex_string = LTrim$(Tems_Data.hex_string)
       If Mid(Tems_Data.hex_string, 1, 5) = "06 15" Then
          dtx = Mid(Tems_Data.hex_string, 7, 2)
          dtx = "&h" + dtx
          val_dtx = Val(dtx)
          val_dtx = (val_dtx And &H7F) \ 64
          If val_dtx = 1 Then
             Tems_Data.FieldCol3(10) = "YES"
          Else
             Tems_Data.FieldCol3(10) = "NO"
          End If
       End If
       If Len(linetxt) > 0 Then
          Tems_Data.ncell_num = "0"
          For j = 0 To 5
              Do While 1
                 findtab = InStr(linetxt, wtab)
                 If findtab = 1 Then
                    linetxt = Right(linetxt, Len(linetxt) - 1)
                    GoTo again
                 End If
                 If findtab = 0 Then
                    GoTo noncell
                 End If
                 'Tems_data.col(35 + j * 3) = Trim(Left(linetxt, Abs(findtab - 1)))
                 Tems_Data.FieldCol3(11 + j * 5) = Trim(Left(linetxt, Abs(findtab - 1)))
                 linetxt = Right(linetxt, Len(linetxt) - findtab)
                 findtab = InStr(linetxt, wtab)
                 Tems_Data.FieldCol3(12 + j * 5) = Trim(Left(linetxt, Abs(findtab - 1)))
                 linetxt = Right(linetxt, Len(linetxt) - findtab)
                 findtab = InStr(linetxt, wtab)
                 If findtab = 0 Then
                    Tems_Data.FieldCol3(13 + j * 5) = Trim(linetxt)
                 Else
                    Tems_Data.FieldCol3(13 + j * 5) = Trim(Left(linetxt, Abs(findtab - 1)))
                 End If
                 linetxt = Right(linetxt, Len(linetxt) - findtab)
                 findstar = InStr(Tems_Data.FieldCol3(13 + j * 5), "*")
                 If findstar = 0 Then
                    Tems_Data.ncell_num = Format(Val(Tems_Data.ncell_num) + 1)
                 '   Exit Do
                 Else
                    Tems_Data.FieldCol3(13 + j * 5) = "99"
                 End If
                 Exit Do
again:
              Loop
          Next
       End If
noncell:
       If Val(Tems_Data.FieldCol2(8)) > 0 Then
          Tems_Data.FER = Format(Int(((Val(Tems_Data.FieldCol2(8)) - Val(Tems_Data.FieldCol2(7))) / Val(Tems_Data.FieldCol2(8))) * 100))
       End If
       'MyValueTemp = Tems_data.FieldCol2(2)
       'If Trim(MyValueTemp) = "" Then
       '   Tems_data.FieldCol2(2) = "9"
       'End If
       'MyValueTemp = Tems_data.FieldCol2(4)
       'If Trim(MyValueTemp) = "" Then
       '   Tems_data.FieldCol2(4) = "9"
       'End If
       Put #2, , Tems_Data
delrecord:
       DoEvents
       If Convert_Stop = True Then
          Put #2, , endd
          Seek #2, 5
          Put #2, , recordno
          Close
          Exit Function
       End If
    Loop
    
    'If fill = 1 Then
    '   Put #2, , fillword
    'End If
    Put #2, , endd
    Seek #2, 5
    Put #2, , recordno

    If ProgressBar1.Value < 100 Then
       ProgressBar1.Value = 100
    End If
    Close

    changedbf = 1
End Function

Sub get_dbf(lines, db)
    wtab = Chr$(9)
    findtab = InStr(lines, wtab)
    If findtab = 0 Then
       db = lines
    Else
       db = Left(lines, findtab - 1)
       lines = Right(lines, Len(lines) - findtab)
    End If
End Sub


'***************************************
'Input: scnname
'Output: dbfname
'***************************************

Function conv_scan(dbfname, scnname) As Integer
Dim colno As String * 2
Dim dbfhead As String * 32
Dim field_2 As String * 4
Dim lines As String
Dim wtab As String * 1
Dim wwend As String * 1
Dim wspace As String * 1
Dim recordno As Integer
Dim tt As String * 1
Dim fill0 As String * 1
Dim kk As String * 1
Dim res As Double

On Error Resume Next

Dim lon_no, lat_no As Integer
  Select Case GPS_NO
   Case 1
        lon_no = 8
        lat_no = 7
   Case 2
        lon_no = 9
        lat_no = 8
   Case 3
        lon_no = 10
        lat_no = 9
   Case 4
        lon_no = 11
        lat_no = 10
   Case 5
        lon_no = 12
        lat_no = 11
   Case 6
        lon_no = 13
        lat_no = 12
  End Select

    Label1.Caption = "正在转换 " + scnname
    Label1.Refresh

recordno = 0
begined = 0
kk = Chr$(13)
fill0 = Chr$(0)
tt = Chr$(33)
wtab = Chr$(9)
wwend = Chr$(26)
wspace = Chr$(20)
field_1.start = space$(1)
Open scnname For Binary As #1

'Seek #1, 139       li_del
'Get #1, , colno    li_del

Do While Not EOF(1)
   Line Input #1, lines
   lines = UCase(lines)
   findtt = InStr(lines, "MEASUREMENT")
   If findtt > 0 Then
      findtt = InStr(lines, wtab)
      If findtt > 0 Then
         lines = Right(lines, Len(lines) - findtt)
         colno = Val(lines)
      Else
         findtt = InStr(lines, ":")
         If findtt > 0 Then
            lines = Right(lines, Len(lines) - findtt)
            colno = Val(lines)
         End If
      End If
      Exit Do
   End If
Loop

Close #1
Open dbfname For Binary As #3

dbf_head.ver = 3
dbf_head.year = Mid(year(Now), 3, 2)
dbf_head.month = month(Now)
dbf_head.day = day(Now)
dbf_head.recordno = 0
dbf_head.header_len = (3 + colno * 2 + 1) * 32 + 1
dbf_head.record_len = 35 + colno * 2 * 4 + 1
'dbf_head.zero = String(fill0, 20)
dbf_head.Zero = String(20, fill0)
Put #3, , dbf_head

Open scnname For Input As #4
findtt = 1
Do While findtt <> 0
   Line Input #4, lines
   findtt = InStr(lines, tt)
Loop
For i = 1 To 3
    findtab = InStr(lines, wtab)
    lines = Right(lines, Len(lines) - findtab)
Next
dbf_field.Name = "LON" + String(8, fill0)
dbf_field.Type = "N"
dbf_field.len = 12
dbf_field.Dec = 6
dbf_field.off = 1
Put #3, , dbf_field
dbf_field.Name = "LAT" + String(8, fill0)
dbf_field.Type = "N"
dbf_field.len = 12
dbf_field.Dec = 6
dbf_field.off = dbf_field.off + 10
Put #3, , dbf_field
dbf_field.Name = "TIME" + String(7, fill0)
dbf_field.Type = "C"
dbf_field.len = 11
dbf_field.Dec = 0
dbf_field.off = dbf_field.off + 10
Put #3, , dbf_field
Call get_dbf(lines, name0)
dbf_field.Name = name0 + String(11 - Len(name0), fill0)
dbf_field.Type = "N"
dbf_field.len = 4
dbf_field.Dec = 0
dbf_field.off = dbf_field.off + 11
Put #3, , dbf_field

For i = 2 To colno * 2
   Call get_dbf(lines, name0)
   dbf_field.Name = name0 + String(11 - Len(name0), fill0)
   dbf_field.Type = "N"
   dbf_field.len = 4
   dbf_field.Dec = 0
   dbf_field.off = dbf_field.off + 4
   Put #3, , dbf_field
Next
Put #3, , kk

lenth = FileLen(scnname)
nline = lenth / 150
bline = Fix(nline / 100)
percent_step = 1
If bline = 0 Then
   bline = 1
   percent_step = 100 / nline
End If
bs = 1
scnline = 0
ProgressBar1.Value = 0

Do While 1
Line Input #4, lines

iss = Mid(lines, 1, 1)
If iss = "N" Or iss = "n" Then
   begined = 1
   Exit Do
End If
Loop

Do While Not EOF(4)
   If begined = 0 Then
      Line Input #4, lines
   End If
'If recordno > 10 Then
'Exit Do
'Else
   scnline = scnline + 1
   If scnline = bs * bline And ProgressBar1.Value < 99 Then
      ProgressBar1.Value = ProgressBar1.Value + percent_step
      bs = bs + 1
   End If

   lend = Len(lines)
   findtab = InStr(lines, wtab)
   field_1.latdbf = Left(lines, Abs(findtab - 1))
'   MsgBox lines
'   MsgBox field_1.latdbf
   l = Len(field_1.latdbf)
   findN = InStr(field_1.latdbf, "N")
   field_1.latdbf = Right(field_1.latdbf, l - findN)
   field_1.latdbf = RTrim$(field_1.latdbf)
   field_1.latdbf = LTrim$(field_1.latdbf)
   findspace = InStr(field_1.latdbf, " ")

   If findspace = lat_no Then
      field_1.latdbf = CDbl(field_1.latdbf)
      res = Fix(field_1.latdbf / 100) + (field_1.latdbf - Fix(field_1.latdbf / 100) * 100) / 60
      field_1.latdbf = Left(res, 12)
      lat = RTrim$(field_1.latdbf)
      If Len(lat) < 9 Then
         lat = lat + String(9 - Len(lat), "0")
      End If
        field_1.latdbf = lat
     
   Else
      GoTo delrecord
   End If
   lines = Right(lines, lend - findtab)
   lend = Len(lines)
   findtab = InStr(lines, wtab)
   field_1.londbf = Left(lines, findtab - 1)
   l = Len(field_1.londbf)
   findN = InStr(field_1.londbf, "E")
   field_1.londbf = Right(field_1.londbf, l - findN)
   l = Len(field_1.londbf)
   field_1.londbf = RTrim(field_1.londbf)
   field_1.londbf = LTrim(field_1.londbf)
   findspace = InStr(field_1.londbf, " ")
   If findspace = lon_no Then
      field_1.londbf = CDbl(field_1.londbf)
      res = Fix(field_1.londbf / 100) + (field_1.londbf - Fix(field_1.londbf / 100) * 100) / 60
      field_1.londbf = Left(res, 10)
      lon = RTrim$(field_1.londbf)
      If Len(lon) < 10 Then
         lon = lon + String(10 - Len(lon), "0")
      End If
        field_1.londbf = lon
   End If
   lines = Right(lines, lend - findtab)
   lend = Len(lines)
   findtab = InStr(lines, wtab)
   field_1.timedbf = Left(lines, findtab - 1)
   
   Put #3, , field_1
   
   lines = Right(lines, lend - findtab)
   lend = Len(lines)
   For i = 1 To colno * 2
      findtab = InStr(lines, wtab)
      If findtab = 1 Or findtab = 0 And Len(lines) = 0 Then
         field_2 = space$(4)
         lines = Right(lines, lend - 1)
         lend = Len(lines)
      Else
         If findtab = 0 Then
            col = lines
             field_2 = col
         Else
            field_2 = Left(lines, findtab - 1)
            find2d = InStr(field_2, "-")
            If find2d <> 0 Then
               field_2 = Right(field_2, Len(field_2) - find2d)
            End If
            col = RTrim$(field_2)
              field_2 = col
                    lines = Right(lines, lend - findtab)
            lend = Len(lines)
         End If
      End If
      Put #3, , field_2
   Next

recordno = recordno + 1
delrecord:
'End If
    DoEvents
    If Convert_Stop = True Then
       Put #3, , wwend
       Seek #3, 5
       Put #3, , recordno
       Close
       Exit Function
    End If
begined = 0
Loop
Put #3, , wwend
Seek #3, 5
Put #3, , recordno
If ProgressBar1.Value < 100 Then
   ProgressBar1.Value = 100
End If
Close
End Function


Private Sub Command1_Click()
    On Error Resume Next
    
    If Convert_Flag = 2 Then
       If (MsgBox("确实要中止平滑处理吗?", 33, "提示")) = 1 Then
          Convert_Stop = True
       End If
    Else
       If (MsgBox("确实要中止转换吗?", 33, "提示")) = 1 Then
          Convert_Stop = True
       End If
    End If
End Sub

Private Sub Timer1_Timer()
    Dim temp_file As String
    On Error Resume Next
    Convert_Stop = False
        pot = Len(sinput) - 4
        soutput = Left$(sinput, pot)

        If Data_Tran_Flag = 1 Then
           If tran_del = 2 Then
              If Mid(soutput, Len(soutput) - 8, 1) = "\" Then
                  soutput = Left(soutput, Len(soutput) - 1) + "F"
              Else
                  soutput = soutput + "F"
              End If
           Else
              If tran_del = 3 Then
                 If Mid(soutput, Len(soutput) - 8, 1) = "\" Then
                     soutput = Left(soutput, Len(soutput) - 1) + "e"
                 Else
                     soutput = soutput + "e"
                 End If
              End If
           End If
           soutput = soutput + ".DBF"
        End If

        If Data_Tran_Flag = 2 Then
           If tran_del = 2 Then
              If Mid(soutput, Len(soutput) - 8, 1) = "\" Then
                  soutput = Left(soutput, Len(soutput) - 1) + "F"
              Else
                  soutput = soutput + "F"
              End If
           End If
           If tran_del = 3 Then
              If Mid(soutput, Len(soutput) - 8, 1) = "\" Then
                  soutput = Left(soutput, Len(soutput) - 1) + "e"
              Else
                  soutput = soutput + "e"
              End If
           End If
           temp_file = Left(soutput, Len(soutput) - 2) + "t.DBF"
           soutput = soutput + ".DBF"
        End If

        If sys = 0 Then
           If Data_Tran_Flag = 1 Then
              Convert_Flag = 1
              If Menu_Flag = 2222 Or Menu_Flag = 4444 Then
                 'If tran_del = 3 Or tran_del = 2 Then
                    Call ConvertANT(sinput, soutput)
                 'Else
                 '   FileCopy sinput, soutput
                 'End If
              Else
                 If Menu_Flag = 121 Then
                    i = conv_gsm(sinput, soutput)
                 Else
                    i = conv_tems98(sinput, soutput)
                 End If
              End If
           End If

           If Data_Tran_Flag = 2 Then
              FileCopy soutput, temp_file
              Convert_Flag = 2
              Call DRAG(temp_file, soutput)
              Kill temp_file
           End If
        Else
           Convert_Flag = 1
           i = conv_scan(soutput, sinput)
        End If
ExitSub:
  Unload Me
End Sub

Sub ConvertANT(txtname, dbfname)
    Dim fillword As String * 77
    Dim fill As Integer, recordno As Integer
    Dim dbfhead As String * 1919
    Dim sword As String * 2
    Dim linetxt As String, wtab As String
    Dim findtab, findN, lenline, l As Integer
    Dim col As String, lat As String, lon As String
    Dim spac As String * 1, endd As String * 1
    Dim res As Double
    Dim dtx As String
    Dim val_dtx
    Dim old_lon As String * 12, old_lat As String * 12
    Dim GetHead As dbfhead
    Dim FileFlag As Integer
    Dim old_data1 As street
    Dim old_data2 As oldtypeNormal
    Dim old_data3 As typeNormal
    Dim DelFlag As Boolean
    Dim MyValueTemp As String
    Dim MyStrTmp1 As String, MyStrTmp2 As String, MyStrTmp3 As String
    Dim MyHexString As String, MyMark1 As String
    
    On Error Resume Next
    old_lon = space$(12)
    old_lat = space$(12)
    endd = Chr(26)
    FileFlag = 0
    
    'Gsm_FileName = Gsm_Path + "\data_new.dbf"
    'FileCopy Gsm_FileName, dbfname
    'Lixuhua:
    'Open Gsm_FileName For Binary As #1
    If Dir(dbfname, 0) <> "" Then
       Kill dbfname
    End If
    Open dbfname For Binary As #2
    hDbfFile = 2
    'MakeNormalFile
    MakeNormalMarkFile
    
    'Seek #2, 2466 + 384
    Seek #2, 151 * 32 + 1 + 1
    recordno = 0
    spac = space$(1)
    wtab = Chr$(9)
    eofword = Chr$(26)
    fillword = space$(77)
    i = 0
    j = 1
    MyHexString = ""
    MyMark1 = ""
    Label1.Caption = "正在转换 " + txtname
    Label1.Refresh
    Open txtname For Binary As #3
    Get #3, , GetHead
    If GetHead.header_len = 1921 Then     '1 --- 1921
       FileFlag = 1                       '2 --- 2465
       Seek #3, 1922                      '3 --- 2465 + 384
    ElseIf GetHead.header_len = 2465 Then '0 --- 151 * 32 + 1
       FileFlag = 2
       Seek #3, 2466
    ElseIf GetHead.header_len = 2465 + 384 Then
       FileFlag = 3
       Seek #3, 2465 + 384 + 1
    Else
       If Not (tran_del = 3 Or tran_del = 2) Then
          Close
          FileCopy txtname, dbfname
          Open dbfname For Binary As #2
          Seek #2, 151 * 32 + 1 + 1
          For i = 1 To 30
              Get #2, , s_data
              'MyValueTemp = s_data.message
              'If Trim(MyValueTemp) = "HEADER" Then
              '   MyHexString = s_data.hex_string
              '   MyMark1 = s_data.MARK1
              'End If
              If Val(s_data.lon) > 0 And Val(s_data.lat) > 0 Then
                 If i = 1 Then
                    Close
                    Exit Sub
                 End If
                 old_lon = s_data.lon
                 old_lat = s_data.lat
                 For j = 0 To i - 1
                     'Seek #2, 151 * 32 + 1 + 1 + j * (694 + 1)
                     'Get #2, , s_data
                     Get #2, 151 * 32 + 1 + 1 + j * 694, s_data
                     'If j = 0 And MyHexString <> "" Then
                     '   s_data.message = "HEADER"
                     '   s_data.hex_string = MyHexString
                     '   s_data.MARK1 = MyMark1
                     'End If
                     s_data.lon = old_lon
                     s_data.lat = old_lat
                     'Seek #2, 151 * 32 + 1 + 1 + j * (694 + 1)
                     'Put #2, , s_data
                     Put #2, 151 * 32 + 1 + 1 + j * 694, s_data
                 Next
                 Exit For
              End If
          Next
          Close
          Exit Sub
       End If
       FileFlag = 0
       Seek #3, 151 * 32 + 1 + 1
    End If
    nline = FileLen(txtname) / 428
    bline = Fix(nline / 100)
    bs = 1
    ProgressBar1.Value = 0
    longline = 0
    old_lon = ""
    old_lat = ""
    
       s_data.time = space$(12)
       s_data.frame = space$(10)
       s_data.lon = space$(12)
       s_data.lat = space$(12)
       s_data.message = space$(30)
       s_data.hex_string = space$(90)
       For i = 1 To 15
           s_data.FieldCol1(i) = space$(5)
       Next
       For i = 1 To 10
           s_data.FieldCol2(i) = space$(3)
       Next
       s_data.FER = space$(5)
       s_data.SQI = space$(5)
       s_data.MARK = space$(30)
       s_data.MARK1 = space$(60)
       s_data.MARK2 = space$(20)
       s_data.Cell_2 = space$(5)
       For i = 1 To 40
           If i <> 10 Then
              s_data.FieldCol3(i) = space$(3)
           End If
       Next
       s_data.ncell_num = "0"
       For i = 1 To 12
           s_data.NewField(i) = space$(3)
       Next
       For i = 1 To 20
           s_data.ScanField(i).Scan_AR = space$(3)
           s_data.ScanField(i).Scan_RX = space$(2)
           s_data.ScanField(i).Scan_BS = space$(2)
       Next
    Do While Not EOF(3)
       DelFlag = False
       If FileFlag = 1 Then
          Get #3, , old_data1
          s_data.time = old_data1.time
          s_data.frame = old_data1.frame
          s_data.lon = old_data1.lon
          s_data.lat = old_data1.lat
          s_data.message = old_data1.message
          s_data.hex_string = old_data1.hex_string
          For i = 1 To 15
              s_data.FieldCol1(i) = old_data1.col(i)
          Next
          For i = 1 To 8
              s_data.FieldCol2(i) = old_data1.col(i + 15)
          Next
          s_data.Cell_2 = old_data1.col(24)
          j = 1
          For i = 1 To 40
              If i = 14 Or i = 19 Or i = 24 Or i = 29 Or i = 34 Or i = 39 Then
                 i = i + 1
              Else
                 s_data.FieldCol3(i) = old_data1.col(j + 24)
                 j = j + 1
              End If
          Next
          s_data.ncell_num = old_data1.ncell_num
       ElseIf FileFlag = 2 Then
          Get #3, , old_data2
          s_data.time = old_data2.time
          s_data.frame = old_data2.frame
          s_data.lon = old_data2.lon
          s_data.lat = old_data2.lat
          s_data.message = old_data2.message
          s_data.hex_string = old_data2.hex_string
          For i = 1 To 15
              s_data.FieldCol1(i) = old_data2.FieldCol1(i)
          Next
          For i = 1 To 10
              s_data.FieldCol2(i) = old_data2.FieldCol2(i)
          Next
          s_data.FER = old_data2.FER
          s_data.SQI = old_data2.SQI
          s_data.MARK = old_data2.MARK
          s_data.Cell_2 = old_data2.Cell_2
          For i = 1 To 40
              s_data.FieldCol3(i) = old_data2.FieldCol3(i)
          Next
          s_data.ncell_num = old_data2.ncell_num
       ElseIf FileFlag = 3 Then
          Get #3, , old_data3
          s_data.time = old_data3.time
          s_data.frame = old_data3.frame
          s_data.lon = old_data3.lon
          s_data.lat = old_data3.lat
          s_data.message = old_data3.message
          s_data.hex_string = old_data3.hex_string
          For i = 1 To 15
              s_data.FieldCol1(i) = old_data3.FieldCol1(i)
          Next
          For i = 1 To 10
              s_data.FieldCol2(i) = old_data3.FieldCol2(i)
          Next
          s_data.FER = old_data3.FER
          s_data.SQI = old_data3.SQI
          s_data.MARK = old_data3.MARK
          s_data.Cell_2 = old_data3.Cell_2
          For i = 1 To 40
              s_data.FieldCol3(i) = old_data3.FieldCol3(i)
          Next
          s_data.ncell_num = old_data3.ncell_num
          For i = 1 To 12
              s_data.NewField(i) = old_data3.NewField(i)
          Next
       Else
          Get #3, , s_data
       End If
       longline = longline + 1
       If longline = bs * bline And ProgressBar1.Value < 99 Then
          ProgressBar1.Value = ProgressBar1.Value + 1
          bs = bs + 1
       End If
    If tran_del = 3 Or tran_del = 2 Then
       If InStr(s_data.lat, Chr(0)) > 0 Then
          s_data.lat = Left(s_data.lat, InStr(s_data.lat, Chr(0)) - 1)
       End If
       If InStr(s_data.lon, Chr(0)) > 0 Then
          s_data.lon = Left(s_data.lon, InStr(s_data.lon, Chr(0)) - 1)
       End If
       If Val(s_data.lon) = 0 And Val(s_data.lat) = 0 Then
          MyValueTemp = s_data.message
          If Trim(MyValueTemp) = "HEADER" Then
             MyHexString = s_data.hex_string
             MyMark1 = s_data.MARK1
          End If
          GoTo delrecord
       End If
       If tran_del = 2 Then
          If ((Trim(old_lon) = Trim(s_data.lon) And Trim(old_lat) = Trim(s_data.lat)) And (Trim(s_data.lon) <> "" And Trim(s_data.lat) <> "")) Or ((Trim(old_lon) <> Trim(s_data.lon) And Trim(old_lat) <> Trim(s_data.lat)) And (Trim(s_data.lon) = "" And Trim(s_data.lat) = "")) Then
             DelFlag = True
             'GoTo delrecord
          End If
       Else
          If (Val(s_data.lat) - Val(old_lat)) * (Val(s_data.lat) - Val(old_lat)) + (Val(s_data.lon) - Val(old_lon)) * (Val(s_data.lon) - Val(old_lon)) < 0.000001 Then
             MyValueTemp = s_data.message
             If Trim(MyValueTemp) = "END-OF-FILE MARKER" Then
                recordno = recordno + 1
                Put #2, , s_data
             End If
             GoTo delrecord
          End If
       End If
       old_lon = Trim(s_data.lon)
       old_lat = Trim(s_data.lat)
    End If
       If DelFlag Then
          DelFlag = False
          MyValueTemp = s_data.message
          Select Case UCase(Trim(MyValueTemp))
             Case "ASSIGNMENT COMPLETE", "CONGESTION CONTROL", "NO SERVICE REPORT", "HANDOVER COMPLETE", "HANDOVER FAILURE", "RELEASE"
             Case "HANDOVER COMMAND", "CONNECT", "CHANNEL REQUEST", "SETUP", "ASSIGNMENT COMMAND", "ASSIGNMENT COMPLETE", "DISCONNECT", "END-OF-FILE MARKER", "SYSTEM INFORMATION TYPE 5"
             Case "ASSIGNMENT FAILURES", "ALERTING", "LOCATION UPDATING REQUEST", "LOCATION UPDATING ACCEPT", "LOCATION UPDATING REJECT", "CHANNEL RELEASE"
             
             Case Else
                  MyStrTmp1 = s_data.MARK
                  MyStrTmp2 = s_data.MARK1
                  MyStrTmp3 = s_data.MARK2
                  If Trim(MyStrTmp1) = "" And Trim(MyStrTmp2) = "" And Trim(MyStrTmp3) = "" Then
                     GoTo delrecord
                  End If
          End Select
       End If
       If Val(s_data.FieldCol2(3)) = 0 Then
          s_data.FieldCol2(3) = s_data.FieldCol2(1)
          s_data.FieldCol2(4) = s_data.FieldCol2(2)
       End If
       'MyValueTemp = s_data.FieldCol2(2)
       'If Trim(MyValueTemp) = "" Then
       '   s_data.FieldCol2(2) = "9"
       'End If
       'MyValueTemp = s_data.FieldCol2(4)
       'If Trim(MyValueTemp) = "" Then
       '   s_data.FieldCol2(4) = "9"
       'End If
       If MyHexString <> "" Then
          MyValueTemp = s_data.message
          If Trim(MyValueTemp) <> "HEADER" Then
             MyStrTmp1 = s_data.message
             MyStrTmp2 = s_data.hex_string
             MyStrTmp3 = s_data.MARK1
             s_data.message = "HEADER"
             s_data.hex_string = MyHexString
             s_data.MARK1 = MyMark1
             recordno = recordno + 1
             Put #2, , s_data
             s_data.message = MyStrTmp1
             s_data.hex_string = MyStrTmp2
             s_data.MARK1 = MyStrTmp3
          End If
          MyHexString = ""
          MyMark1 = ""
       End If
       If Val(s_data.lon) > 0 And Val(s_data.lat) > 0 Then
          recordno = recordno + 1
          Put #2, , s_data
       End If
delrecord:
       DoEvents
       If Convert_Stop = True Then
          Put #2, , endd
          Seek #2, 5
          Put #2, , recordno
          Close
          Exit Sub
       End If
    Loop
    
    Put #2, , endd
    Seek #2, 5
    Put #2, , recordno

    If ProgressBar1.Value < 100 Then
       ProgressBar1.Value = 100
    End If
    Close

    changedbf = 1
End Sub

Function conv_tems98(txtname, dbfname) As Integer
    Dim fillword As String * 77
    Dim fill As Integer, recordno As Integer
    Dim dbfhead As String * 1919
    Dim sword As String * 2
    Dim linetxt As String, wtab As String
    Dim findtab, findN, lenline, l As Integer
    Dim col As String, lat As String, lon As String
    Dim spac As String * 1, endd As String * 1
    Dim res As Double
    Dim dtx As String
    Dim val_dtx
    Dim old_lon As String * 12, old_lat As String * 12
    Dim MyValueTemp As String
    
    On Error Resume Next
    ncell_num = 0
    old_lon = space$(12)
    old_lat = space$(12)
    Dim lon_no, lat_no As Integer
    Select Case GPS_NO
           Case 1
                lon_no = 8
                lat_no = 7
           Case 2
                lon_no = 9
                lat_no = 8
           Case 3
                lon_no = 10
                lat_no = 9
           Case 4
                lon_no = 11
                lat_no = 10
           Case 5
                lon_no = 12
                lat_no = 11
           Case 6
                lon_no = 13
                lat_no = 12
    End Select
    
    'Gsm_FileName = Gsm_Path + "\data_new.dbf"
    'FileCopy Gsm_FileName, dbfname
    'Lixuhua:
    'Open Gsm_FileName For Binary As #1
    If Dir(dbfname, 0) <> "" Then
        Kill dbfname
    End If
    hDbfFile = 2
    Open dbfname For Binary As #2
    MakeNormalFile
    Seek #2, 2466 + 384
    'll = FileLen(Gsm_FileName)
    'Get #1, ll, endd
    recordno = 0
    spac = space$(1)
    wtab = Chr$(9)
    eofword = Chr$(26)
    fillword = space$(77)
    i = 0
    j = 1
    Label1.Caption = "正在转换 " + txtname
    Label1.Refresh
    Open txtname For Input As #3
    nline = FileLen(txtname) / 260
    bline = Fix(nline / 100)
    bs = 1
    ProgressBar1.Value = 0
    longline = 0
    'Seek #3, 1061
    For i = 1 To 3
        Line Input #3, linetxt
    Next
    Tems_Data.a = " "
    Tems_Data.FieldCol3(10) = space$(3)
    Do While Not EOF(3)
    
       Tems_Data.time = space$(12)
       Tems_Data.frame = space$(10)
       Tems_Data.lon = space$(12)
       Tems_Data.lat = space$(12)
       Tems_Data.message = space$(30)
       Tems_Data.hex_string = space$(90)
       For i = 1 To 15
           Tems_Data.FieldCol1(i) = space$(5)
       Next
       For i = 1 To 10
           Tems_Data.FieldCol2(i) = space$(3)
       Next
       Tems_Data.FER = space$(5)
       Tems_Data.SQI = space$(5)
       Tems_Data.MARK = space$(16)
       Tems_Data.Cell_2 = space$(5)
       For i = 1 To 40
           If i <> 10 Then
              Tems_Data.FieldCol3(i) = space$(3)
           End If
       Next
       Tems_Data.ncell_num = "0"
       For i = 1 To 12
           Tems_Data.NewField(i) = space$(3)
       Next
       
       Line Input #3, linetxt
       longline = longline + 1
       If longline = bs * bline And ProgressBar1.Value < 99 Then
          ProgressBar1.Value = ProgressBar1.Value + 1
          bs = bs + 1
       End If

       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       Tems_Data.time = Left(linetxt, Abs(findtab - 1))
       linetxt = Right(linetxt, lenline - findtab)
       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       Tems_Data.frame = Left(linetxt, Abs(findtab - 1))
       linetxt = Right(linetxt, lenline - findtab)
    
       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       Tems_Data.lat = Left(linetxt, Abs(findtab - 1))
       l = Len(Tems_Data.lat)
       findN = InStr(Tems_Data.lat, "N")
       Tems_Data.lat = Right(Tems_Data.lat, l - findN)
       Tems_Data.lat = LTrim(Tems_Data.lat)
       findspace = InStr(Tems_Data.lat, " ")
    
       If findspace = lat_no Then
          res = Fix(Val(Tems_Data.lat) / 100) + (Val(Tems_Data.lat) - Fix(Val(Tems_Data.lat) / 100) * 100) / 60
          Tems_Data.lat = Left(res, 9)
          lat = LTrim$(Tems_Data.lat)
          If Len(lat) < 9 Then
             lat = lat + String(9 - Len(lat), "0")
          End If
          Tems_Data.lat = lat
       Else
          If Not Data_Report Then
             GoTo delrecord
          End If
       End If
       linetxt = Right(linetxt, lenline - findtab)
       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       Tems_Data.lon = Left(linetxt, Abs(findtab - 1))
       l = Len(Tems_Data.lon)
       findN = InStr(Tems_Data.lon, "E")
       Tems_Data.lon = Right(Tems_Data.lon, l - findN)
       Tems_Data.lon = LTrim(Tems_Data.lon)
       findspace = InStr(Tems_Data.lon, " ")
       If findspace = lon_no Then
          res = Fix(Val(Tems_Data.lon) / 100) + (Val(Tems_Data.lon) - Fix(Val(Tems_Data.lon) / 100) * 100) / 60
          Tems_Data.lon = Left(res, 10)
          lon = LTrim$(Tems_Data.lon)
          If Len(lon) < 10 Then
             lon = lon + String(10 - Len(lon), "0")
          End If
          Tems_Data.lon = lon
       End If
       If tran_del = 2 Then
          If Tems_Data.lat = old_lat And Tems_Data.lon = old_lon Then
             If Not Data_Report Then
                GoTo delrecord
             End If
          Else
             old_lat = Tems_Data.lat
             old_lon = Tems_Data.lon
          End If
       Else
          If tran_del = 3 Then
             If (Val(Tems_Data.lat) - Val(old_lat)) * (Val(Tems_Data.lat) - Val(old_lat)) + (Val(Tems_Data.lon) - Val(old_lon)) * (Val(Tems_Data.lon) - Val(old_lon)) < 0.000001 Then
                If Not Data_Report Then
                   GoTo delrecord
                End If
             Else
                old_lat = Tems_Data.lat
                old_lon = Tems_Data.lon
             End If
          End If
       End If
             
       linetxt = Right(linetxt, lenline - findtab)
       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       If findtab > 1 Then
          Tems_Data.message = Left(linetxt, Abs(findtab - 1))
       End If
       linetxt = Right(linetxt, lenline - findtab)
       lenline = Len(linetxt)
       findtab = InStr(linetxt, wtab)
       If findtab > 1 Then
          Tems_Data.hex_string = Left(linetxt, Abs(findtab - 1))
       End If
       recordno = recordno + 1
   
       linetxt = Right(linetxt, lenline - findtab)
       lenline = Len(linetxt)
       For i = 1 To 34 + 4
          findtab = InStr(linetxt, wtab)
          If findtab = 1 Then
             linetxt = Right(linetxt, lenline - 1)
             lenline = lenline - 1
          Else
             If findtab = 0 Then
                fill = 1
                Exit For
             Else
             If i <> 34 Then
                If i >= 1 And i <= 15 Then
                   If i = 10 Then
                      Tems_Data.FieldCol1(i) = Format(CDbl("&H" & Trim(Left(linetxt, Abs(findtab - 1)))))
                   Else
                      Tems_Data.FieldCol1(i) = Trim(Left(linetxt, Abs(findtab - 1)))
                   End If
                ElseIf i >= 16 And i <= 25 Then
                   Tems_Data.FieldCol2(i - 15) = Trim(Left(linetxt, Abs(findtab - 1)))
                ElseIf i = 26 Then
                   Tems_Data.FER = Trim(Left(linetxt, Abs(findtab - 1)))
                ElseIf i = 27 Then
                   Tems_Data.SQI = Trim(Left(linetxt, Abs(findtab - 1)))
                ElseIf i = 28 Then
                   Tems_Data.Cell_2 = Trim(Left(linetxt, Abs(findtab - 1)))
                Else
                   Tems_Data.FieldCol3(i - 24) = Trim(Left(linetxt, Abs(findtab - 1)))
                End If
             End If
             linetxt = Right(linetxt, lenline - findtab)
             lenline = Len(linetxt)
             End If
          End If
       Next
       fill = 0
       If Val(Tems_Data.FieldCol2(3)) = 0 Then
          Tems_Data.FieldCol2(3) = Tems_Data.FieldCol2(1)
          Tems_Data.FieldCol2(4) = Tems_Data.FieldCol2(2)
       End If
       Tems_Data.hex_string = LTrim$(Tems_Data.hex_string)
       If Mid(Tems_Data.hex_string, 1, 5) = "06 15" Then
          dtx = Mid(Tems_Data.hex_string, 7, 2)
          dtx = "&h" + dtx
          val_dtx = Val(dtx)
          val_dtx = (val_dtx And &H7F) \ 64
          If val_dtx = 1 Then
             Tems_Data.FieldCol3(10) = "YES"
          Else
             Tems_Data.FieldCol3(10) = "NO"
          End If
       End If
       If Len(linetxt) > 0 Then
          Tems_Data.ncell_num = "0"
          For j = 0 To 5
              Do While 1
                 findtab = InStr(linetxt, wtab)
                 If findtab = 1 Then
                    linetxt = Right(linetxt, Len(linetxt) - 1)
                    GoTo again
                 End If
                 If findtab = 0 Then
                    GoTo noncell
                 End If
                 Tems_Data.FieldCol3(11 + j * 5) = Trim(Left(linetxt, Abs(findtab - 1)))
                 linetxt = Right(linetxt, Len(linetxt) - findtab)
                 findtab = InStr(linetxt, wtab)
                 Tems_Data.FieldCol3(12 + j * 5) = Trim(Left(linetxt, Abs(findtab - 1)))
                 linetxt = Right(linetxt, Len(linetxt) - findtab)
                 findtab = InStr(linetxt, wtab)
                 Tems_Data.FieldCol3(13 + j * 5) = Trim(Left(linetxt, Abs(findtab - 1)))
                 linetxt = Right(linetxt, Len(linetxt) - findtab)
                 findtab = InStr(linetxt, wtab)
                 findstar = InStr(Tems_Data.FieldCol3(13 + j * 5), "*")
                 If findstar = 0 Then
                    Tems_Data.ncell_num = Format(Val(Tems_Data.ncell_num) + 1)
                 Else
                    Tems_Data.FieldCol3(13 + j * 5) = "99"
                 End If
                 Tems_Data.FieldCol3(14 + j * 5) = Trim(Left(linetxt, Abs(findtab - 1)))
                 linetxt = Right(linetxt, Len(linetxt) - findtab)
                 findtab = InStr(linetxt, wtab)
                 If findtab = 0 Then
                    Tems_Data.FieldCol3(15 + j * 5) = Trim(linetxt)
                 Else
                    Tems_Data.FieldCol3(15 + j * 5) = Trim(Left(linetxt, Abs(findtab - 1)))
                 End If
                 linetxt = Right(linetxt, Len(linetxt) - findtab)
                 Exit Do
again:
              Loop
          Next
       End If
noncell:
       'MyValueTemp = Tems_data.FieldCol2(2)
       'If Trim(MyValueTemp) = "" Then
       '   Tems_data.FieldCol2(2) = "9"
       'End If
       'MyValueTemp = Tems_data.FieldCol2(4)
       'If Trim(MyValueTemp) = "" Then
       '   Tems_data.FieldCol2(4) = "9"
       'End If

       Put #2, , Tems_Data
delrecord:
       DoEvents
       If Convert_Stop = True Then
          Put #2, , endd
          Seek #2, 5
          Put #2, , recordno
          Close
          Exit Function
       End If
    Loop
    
    'If fill = 1 Then
    '   Put #2, , fillword
    'End If
    Put #2, , endd
    Seek #2, 5
    Put #2, , recordno

    If ProgressBar1.Value < 100 Then
       ProgressBar1.Value = 100
    End If
    Close
    
    changedbf = 1
End Function
