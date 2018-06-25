VERSION 5.00
Begin VB.Form Data_Convert 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数据转换"
   ClientHeight    =   1425
   ClientLeft      =   825
   ClientTop       =   5595
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Convert.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1425
   ScaleWidth      =   4815
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   320
      Left            =   1875
      TabIndex        =   1
      Top             =   1035
      Width           =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   180
      Left            =   345
      TabIndex        =   0
      Top             =   225
      Width           =   540
   End
End
Attribute VB_Name = "Data_Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bc(1 To 16) As String * 3
Dim mbcch(1 To 16) As String * 3
Dim end_frag As Boolean
Dim mb As Integer
Dim New_Cell As NewCellStru
Dim New_Ncell As NewCellStru
Dim Italtel_Nonbcch As String
Dim Prev_Bs_no As String * 10
Dim Get_Bs_no As Boolean
Dim Public_Line As String
Dim Have_Read As Boolean
'*****************************
'StsEricssonHex
Dim FileNumber3 As Integer, FileNumber1 As Integer, FileNumber2 As Integer
Dim CounterType() As Integer
Dim CounterData(1 To 20) As String
Dim TchFlag As Boolean
Dim CounterNum As Integer, ObjectRecNum As Integer, ObjectTypeNum As Integer
Dim StsRecordno As Long
Dim StsPercentStep As Integer, IncreasePer As Integer, LinesPer As Integer, IncreaseLine As Integer
'*****************************
Dim Nortel_Line As String

Sub Ncell_Ericsson()
    Dim s_data As NewCellStru
    Dim gline As String
    Dim d_end As String * 1
    Dim recordno As Long
    Dim f_spa As Integer
    Dim f_cell As String
    Dim bs_no As String * 10
    Dim File_num As Integer
    Dim CheckFlag As Boolean
    Dim BaseNoSave() As String
    Dim NcellCol(1 To 16) As String
    
    On Error Resume Next
    load_new = 0
    load_sam = 0
    File_num = 0
    Do While Trim(convert_filename(File_num + 1)) <> ""
       File_num = File_num + 1
    Loop
    d_end = Chr(26)
    recordno = 0
    Gsm_FileName = Gsm_Path + "\map\cell.dbf"
    ProgressBar1.Value = 0
    Open Gsm_FileName For Binary As #2
        ReDim BaseNoSave(1 To File_num) As String

  For j = 1 To File_num
    CheckFlag = False
    Open convert_filename(j) For Input As #1
        File_Lenth = FileLen(convert_filename(j))
        nline = File_Lenth / 800
        bline = Int(nline / 100)
        percent_step = 2
        If bline = 0 Then
           bline = 1
           percent_step = Int(100 / nline + 0.5) + 1
        End If
        bss = 1
        scnline = 0
        ProgressBar1.Value = 0
        Label1.Caption = "正在转换 " + convert_filename(j)
    Do While Not EOF(1)
none:
       If EOF(1) = True Then GoTo end_p
       Line Input #1, gline
       If Len(gline) = 0 Then GoTo another
       gline = Trim(gline)
       f_spa = InStr(gline, " ")
       If f_spa = 0 Then
          f_cell = gline
       Else
          f_cell = Left(gline, f_spa - 1)
       End If
       If UCase(f_cell) <> "CELL" Then GoTo another
hasfind:
       
       scnline = scnline + 1
       If scnline = bss * bline And ProgressBar1.Value < 98 Then
          ProgressBar1.Value = ProgressBar1.Value + percent_step
          bss = bss + 1
       End If
       bs_no = space$(10)

       For i = 1 To 16
           NcellCol(i) = ""
       Next

       If EOF(1) = True Then GoTo end_p
       Line Input #1, gline
       gline = Trim(gline)
       f_spa = InStr(gline, " ")
       If f_spa > 0 Then
          f_cell = UCase(Left(gline, f_spa - 1))
          If f_cell = "WO" Then
             Line Input #1, gline
             gline = Trim(gline)
          End If
       End If
       bs_no = gline
       'Call bs_lac(gline, a_put.bs_name, a_put.ci, a_put.Lac)
       For k = 1 To 16
           Do While Not EOF(1)
              Line Input #1, gline
              If Len(gline) = 0 Then GoTo nc
              gline = Trim(gline)
              f_spa = InStr(gline, " ")
              If f_spa = 0 Then
                 f_cell = gline
              Else
                 f_cell = Left(gline, f_spa - 1)
              End If
              If UCase(f_cell) = "CELLR" Or UCase(f_cell) = "CELL" Then
                 Exit Do
              Else
                 GoTo nc
              End If
nc:
           Loop
           If UCase(f_cell) = "CELL" Then Exit For
wo:
           If EOF(1) = True Then GoTo end_p
           Line Input #1, gline
           gline = Trim(gline)
           f_spa = InStr(gline, " ")
           If f_spa = 0 Then
              f_cell = gline
           Else
              f_cell = Left(gline, f_spa - 1)
           End If
           If UCase(f_cell) = "NONE" Then
              If CheckFlag = False Then
                 If j > 1 Then
                    For i = 1 To j - 1
                        If Trim(bs_no) = BaseNoSave(File_num) Then
                           GoTo end_p
                        End If
                    Next
                 Else
                    BaseNoSave(File_num) = Trim(bs_no)
                 End If
                 BaseNoSave(File_num) = Trim(bs_no)
                 CheckFlag = True
              End If
              recordno = recordno + 1
                     Pos = 0
       Do While Not EOF(2)
          'Seek #2, 962 + pos * 153
          Seek #2, 1026 + Pos * 309
          Get #2, , s_data
          If UCase(Trim(s_data.bs_no)) = UCase(Trim(bs_no)) Then
             load_sam = load_sam + 1
             For i = 1 To 16
                 s_data.NCELL(i) = NcellCol(i)
             Next
             s_data.time = DATE
             Seek #2, 1026 + Pos * 309
             Put #2, , s_data
             load_new = load_new + 1
             GoTo s2
          End If
          Pos = Pos + 1
       Loop
       Close #2
       Open Gsm_FileName For Binary As #2
              GoTo none
           Else
              If UCase(f_cell) = "WO" Then GoTo wo
           End If
           
           NcellCol(k) = f_cell
           'Call bs_ali(f_cell, a_put.col(k).arfcn_c, a_put.col(k).ci_c, a_put.col(k).bsic_c)
       Next
              If CheckFlag = False Then
                 If j > 1 Then
                    For i = 1 To j - 1
                        If Trim(bs_no) = BaseNoSave(i) Then
                           GoTo end_p
                        End If
                    Next
                 End If
                 BaseNoSave(j) = Trim(bs_no)
                 CheckFlag = True
              End If
       
       recordno = recordno + 1
                     Pos = 0
       Do While Not EOF(2)
          'Seek #2, 962 + pos * 153
          Seek #2, 1026 + Pos * 309
          Get #2, , s_data
          If UCase(Trim(s_data.bs_no)) = UCase(Trim(bs_no)) Then
             load_sam = load_sam + 1
             For i = 1 To 16
                 s_data.NCELL(i) = NcellCol(i)
             Next
             s_data.time = DATE
             Seek #2, 1026 + Pos * 309
             Put #2, , s_data
             load_new = load_new + 1
             GoTo s2
          End If
          Pos = Pos + 1
       Loop
       Close #2
       Open Gsm_FileName For Binary As #2
s2:
       DoEvents
       If Convert_Stop = True Then
          Close
          Exit Sub
       End If
       GoTo hasfind
another:
    Loop
end_p:
    
    Close #1
  Next
    
    Close
    If ProgressBar1.Value < 100 Then
       ProgressBar1.Value = 100
    End If
End Sub

Sub getmbcch()
    Dim li As String, xu As String
    Dim fh As Integer, i As Integer

    On Error Resume Next
    Do While Not EOF(1)
       Line Input #1, li
       li = Trim(li)
       If Len(li) > 0 Then
          fh = InStr(li, " ")
          If fh > 0 Then
             xu = UCase(Left(li, fh - 1))
          Else
             xu = UCase(li)
          End If
          If xu = "MBCCHNO" Then Exit Do
       End If
    Loop
    If EOF(1) = True Then
       end_frag = True
       Exit Sub
    End If
    Line Input #1, li
    li = Trim(li)
    If Len(li) > 0 Then
       fh = InStr(li, " ")
       If fh > 0 Then
          xu = UCase(Left(li, fh - 1))
       Else
          xu = UCase(li)
       End If
       If xu = "WO" Then
          If EOF(1) = True Then
             end_frag = True
             Exit Sub
          End If
          Line Input #1, li
          li = Trim(li)
          If Len(li) = 0 Then
             mb = 0
             Exit Sub
          Else
             i = 1
             Do
                fh = InStr(li, " ")
                mb = i
                If fh = 0 Then
                   mbcch(i) = li
                   Exit Do
                Else
                   mbcch(i) = Left(li, fh - 1)
                   li = Right(li, Len(li) - fh)
                End If
                i = i + 1
             Loop
          End If
       Else
          i = 1
          Do
             fh = InStr(li, " ")
             mb = i
             If fh = 0 Then
                mbcch(i) = li
                Exit Do
             Else
                mbcch(i) = Left(li, fh - 1)
                li = Trim(Right(li, Len(li) - fh))
             End If
             i = i + 1
          Loop
       End If
    Else
       mb = 0
    End If
End Sub

Sub getno(bs_no)
    Dim lines As String, gg As String
    Dim ff As Integer
    On Error Resume Next
    Do While Not EOF(1)
       Line Input #1, lines
       lines = Trim(lines)
       If Len(lines) > 0 Then
          If UCase(lines) = "CELL" Then Exit Do
       End If
    Loop
    If EOF(1) = True Then
       end_frag = True
       Exit Sub
    End If
    Line Input #1, lines
    lines = Trim(lines)
    ff = InStr(lines, " ")
    If ff > 0 Then
       gg = Left(lines, ff - 1)
       If UCase(gg) = "WO" Then
          If EOF(1) = True Then
             end_frag = True
             Exit Sub
          End If
          Line Input #1, lines
          lines = Trim(lines)
       End If
    End If
    bs_no = lines
End Sub

Sub bs_ali(ByVal bs As String, a, b, c)
    On Error Resume Next
    i = 0
    mapinfo.Do "fetch first from cell"
    msg = "select * from cell where col2 = " + Chr(34) + UCase(bs) + Chr(34) + " or col2 = " + Chr(34) + bs + Chr(34) + " into temp"
    mapinfo.Do msg
    row = Val(mapinfo.eval("tableinfo(temp,8)"))
    If row = 0 Then
       a = ""
       b = ""
       c = ""
    Else
       a = mapinfo.eval("temp.arfcn")
       b = mapinfo.eval("temp.ci")
       c = mapinfo.eval("temp.bsic")
    End If

End Sub

Sub bs_lac(ByVal bs As String, a, b, c)
    On Error Resume Next
    i = 0
    mapinfo.Do "fetch first from cell"
    msg = "select * from cell where col2 = " + Chr(34) + UCase(bs) + Chr(34) + " or col2 = " + Chr(34) + bs + Chr(34) + " into temp"
    mapinfo.Do msg
    row = Val(mapinfo.eval("tableinfo(temp,8)"))
    If row = 0 Then
       b = ""
       a = ""
       c = ""
    Else
       a = mapinfo.eval("temp.col1")
       b = mapinfo.eval("temp.ci")
       c = mapinfo.eval("temp.arfcn")
    End If
    mapinfo.Do "close table temp"
End Sub

Sub ncell_motorola()
    Dim recordno As Integer
    Dim linetxt As String
    Dim sameci As String * 5
    Dim largeci As String * 5
    Dim AA As String * 1
    Dim finds As Integer
    Dim FindChar As String * 1
    Dim s_data As NewCellStru
    Dim NcellCol(1 To 16) As String
    Dim MyCi As String
    
    On Error Resume Next
    txtname = convert_filename(1)
    dbfname = Gsm_Path + "\map\cell.dbf"
    Label1.Caption = "正在转换 " + txtname
    load_new = 0
    load_sam = 0

    AA = Chr$(26)
    large = 0
    endfile = 0
    endfile1 = 0
    recordno = 0

    lenth = FileLen(txtname)
    nline = lenth / 300
    bline = Int(nline / 100)
    percent_step = 1
    If bline = 0 Then
       bline = 1
       percent_step = Int(100 / nline + 0.5)
    End If
    bss = 1
    scnline = 0
    ProgressBar1.Value = 0
    
    Open dbfname For Binary As #2
    Open txtname For Input As #3
    Do While Not EOF(3)
       Line Input #3, linetxt
       If Len(linetxt) > 10 Then
          finds = InStr(linetxt, Chr(9))
          If finds > 0 Then
             FindChar = Chr(9)
             linetxt = Right(linetxt, Len(linetxt) - finds)
          Else
             FindChar = " "
          End If
          linetxt = Trim(linetxt)
          Txt = Mid(linetxt, 1, 5)
          If Txt = "460-0" Then
             Exit Do
          End If
       End If
    Loop
    Do While Not EOF(3)
       nb_data = new_data
       recordno = recordno + 1
       scnline = scnline + 1
       If scnline = bss * bline And ProgressBar1.Value < 99 Then
          ProgressBar1.Value = ProgressBar1.Value + percent_step
          bss = bss + 1
       End If
       For i = 1 To 16
           NcellCol(i) = ""
       Next
       no = 1
       large = 0
       findspa = InStr(linetxt, FindChar)
       ci = Left(linetxt, findspa - 1)
       MyCi = Right(ci, 4)
       linetxt = Trim(Right(linetxt, Len(linetxt) - findspa))
       'Call getnb(linetxt, nb_data.col(no).arfcn_c, nb_data.col(no).ci_c, nb_data.col(no).bsic_c)
       Call getnb(linetxt, NcellCol(no))
       Do While no < 16
          Do While Not EOF(3)
             endfile = 0
             Line Input #3, linetxt
             If Len(linetxt) > 10 Then
                finds = InStr(linetxt, FindChar)
                If finds > 0 Then
                   linetxt = Trim(Right(linetxt, Len(linetxt) - finds))
                End If
                Txt = Mid(linetxt, 1, 5)
                If Txt = "460-0" Then
                   Exit Do
                End If
             End If
             endfile = 1
          Loop
          If endfile = 1 Then
             GoTo s2
          End If
          findspa = InStr(linetxt, FindChar)
          ci = Left(linetxt, findspa - 1)
          sameci = (Right(ci, 4))
          If Trim(sameci) = Trim(MyCi) Then
             no = no + 1
             If no = 16 Then
                large = 1
             End If
             linetxt = Trim(Right(linetxt, Len(linetxt) - findspa))
             Call getnb(linetxt, NcellCol(no))
          Else
             Exit Do
          End If
       Loop
    If large = 1 Then
         Do While Not EOF(3)
             endfile1 = 0
             Line Input #3, linetxt
             If Len(linetxt) > 10 Then
                finds = InStr(linetxt, FindChar)
                If finds > 0 Then
                   linetxt = Trim(Right(linetxt, Len(linetxt) - finds))
                End If
                Txt = Mid(linetxt, 1, 5)
                If Txt = "460-0" Then
                   findspa = InStr(linetxt, FindChar)
                   la = Left(linetxt, findspa - 1)
                   largeci = Right(la, 4)
                   If Trim(largeci) <> Trim(MyCi) Then
                      Exit Do
                   End If
                End If
             End If
             endfile1 = 1
         Loop
    End If
    
s2:
       Pos = 0
       Do While Not EOF(2)
          Seek #2, 1026 + Pos * 309
          Get #2, , s_data
          If Trim(s_data.ci) = Trim(MyCi) Then
             load_sam = load_sam + 1
             For i = 1 To 16
                 s_data.NCELL(i) = NcellCol(i)
             Next
             s_data.time = DATE
             Seek #2, 1026 + Pos * 309
             Put #2, , s_data
             load_new = load_new + 1
             Exit Do
          End If
          Pos = Pos + 1
       Loop
       Close #2
       Open dbfname For Binary As #2
       
       DoEvents
       If Convert_Stop = True Then
          Close
          Exit Sub
       End If
       If endfile1 = 1 Then
          GoTo s1
       End If
    Loop

s1:
    Close
    If ProgressBar1.Value < 100 Then
       ProgressBar1.Value = 100
    End If
End Sub

Sub cell_ericsson(sinput1, sinput2, sinput3, sinput4, soutput) 'Convert program
    'Dim field_data As cell_stru
    'Dim s_data As cell_stru
    Dim field_data As NewCellStru
    Dim s_data As NewCellStru
    Dim lines As String
    Dim ci As String * 5
    Dim lee As String * 1
    Dim recordno As Long
    Dim test As String * 1
    Dim bs_no As String * 10
    Dim power_type As String * 3
    Dim bsc_stsge As String * 7
    Dim bsc_type As String * 5
    Dim bts_type As String * 9
    Dim Bcchtemp As String
    On Error Resume Next
    load_sam = 0
    load_new = 0
    Pos = 0
    wrec = 0
    lee = Chr$(26)
    FileCopy sinput2, soutput
    lenth = FileLen(sinput1)
    nline = lenth / 200
    bline = Fix(nline / 100)
    percent_step = 1
    If bline = 0 Then
       bline = 1
       percent_step = 100 / nline
    End If
    bs = 1
    scnline = 0
    ProgressBar1.Value = 0
    Label1.Caption = "正在转换 " + sinput1
    Label1.Refresh
    Open sinput1 For Input As #1
    Open sinput2 For Binary As #2
    Open soutput For Binary As #3
    outlen = FileLen(soutput)
    Seek #2, 5
    Get #2, , recordno
    Seek #3, FileLen(soutput)
    Get #3, , test
    If test <> lee Then
       outlen = outlen + 1
    End If

    Do While Not EOF(1)
    
       field_data.b = space(1)
       field_data.time = space(8)
       field_data.lon = space(12)
       field_data.lat = space(12)
       field_data.Name = space(10)
       field_data.bs_no = space(10)
       field_data.bearing = space(3)
       field_data.downtilt = space(3)
       field_data.max_bts = space(2)
       field_data.max_ms = space(2)
       field_data.ci = space(5)
       field_data.ARFCN = space(3)
       field_data.BSIC = space(3)
       field_data.Lac = space(5)
       field_data.Nonbcch = space(32)
       For i = 1 To 16
           field_data.NCELL(i) = space(10)
       Next
       
       scnline = scnline + 1
       If scnline = bs * bline And ProgressBar1.Value < 99 Then
          ProgressBar1.Value = ProgressBar1.Value + percent_step
          bs = bs + 1
       End If

       Call getline_ericsson(lines)
       If lines = "" Then
          Exit Do
       End If
       finds = InStr(lines, " ")
       field_data.bs_no = Left(lines, finds - 1)
       newbs_no = field_data.bs_no
       lines = LTrim$(Right(lines, Len(lines) - finds))
       finds = InStr(lines, " ")
       lin = Left(lines, finds - 1)
       lines = LTrim$(Right(lines, Len(lines) - finds))
       field_data.Lac = Mid(lin, 8, 4)
       NewLac = field_data.Lac
       For i = 1 To 3
           finds = InStr(lin, "-")
           lin = Right(lin, Len(lin) - finds)
       Next
       field_data.ci = lin
       NewCi = field_data.ci
       finds = InStr(lines, " ")
       field_data.BSIC = Left(lines, finds - 1)
       NewBsic = field_data.BSIC
       lines = LTrim$(Right(lines, Len(lines) - finds))
       finds = InStr(lines, " ")
       field_data.ARFCN = Left(lines, finds - 1)
       NewArfcn = field_data.ARFCN
       Pos = 0
       Do While Not EOF(2)
          'Seek #2, 962 + pos * 153
          Seek #2, 1026 + Pos * 309
          Get #2, , s_data
          If UCase(Trim(s_data.bs_no)) = UCase(Trim(field_data.bs_no)) Then
             load_sam = load_sam + 1
             field_data = s_data
             field_data.bs_no = newbs_no
             field_data.Lac = NewLac
             field_data.ci = NewCi
             field_data.BSIC = NewBsic
             field_data.ARFCN = NewArfcn
             field_data.time = DATE
             Seek #3, 1026 + Pos * 309
             Put #3, , field_data
             GoTo s2
          End If
          Pos = Pos + 1
       Loop
       Close #2
       Open sinput2 For Binary As #2
       load_new = load_new + 1
s2:
       DoEvents
       If Convert_Stop = True Then
          Close
          Exit Sub
       End If
          
    Loop
    Close

    lenth = FileLen(sinput3)
    nline = lenth / 60
    bline = Fix(nline / 100)
    percent_step = 1
    If bline = 0 Then
       bline = 1
       percent_step = 100 / nline
    End If
    bs = 1
    scnline = 0
    ProgressBar1.Value = 0
    Label1.Caption = "正在转换 " + sinput3
    Label1.Refresh
    Open soutput For Binary As #3
    Open sinput3 For Input As #7
    Do While Not EOF(7)

       scnline = scnline + 1
       If scnline = bs * bline And ProgressBar1.Value + percent_step < 100 Then
          ProgressBar1.Value = ProgressBar1.Value + percent_step
          bs = bs + 1
       End If
 
       Call getin(lines)
       If lines = "" Then
          Exit Do
       End If
       Call getfield(lines, bs_no)
       po = 0
       pget = 0
       Do While Not EOF(3)
          Seek #3, 1026 + po * 309
          Get #3, , s_data
          If UCase(Trim(s_data.bs_no)) = UCase(Trim(bs_no)) Then
             Seek #3, 1026 + po * 309
             pget = 1
             Exit Do
          End If
          po = po + 1
       Loop
       If pget = 1 Then
          'Call getfield(lines, s_data.power_type)
          'Call getfield(lines, s_data.bsc_stsge)
          'Call getfield(lines, s_data.bsc_type)
          'Call getfield(lines, s_data.bts_type)
          Put #3, , s_data
       End If
       Close #3
       Open soutput For Binary As #3
       DoEvents
       If Convert_Stop = True Then
          Close
          Exit Sub
       End If
    Loop
    Close
    
    '************
    
    lenth = FileLen(sinput4)
    nline = lenth / 200
    bline = Fix(nline / 100)
    percent_step = 1
    If bline = 0 Then
       bline = 1
       percent_step = 100 / nline
    End If
    bs = 1
    scnline = 0
    ProgressBar1.Value = 0
    Label1.Caption = "正在转换 " + sinput4
    Label1.Refresh
    Open soutput For Binary As #3
    Open sinput4 For Input As #7
    Do While Not EOF(7)

       scnline = scnline + 1
       If scnline = bs * bline And ProgressBar1.Value < 99 Then
          ProgressBar1.Value = ProgressBar1.Value + percent_step
          bs = bs + 1
       End If
       Call getbsno(bs_no)
       If Trim(bs_no) = "" Then Exit Do
       po = 0
       pget = 0
       Do While Not EOF(3)
          Seek #3, 1026 + po * 309
          Get #3, , s_data
          If UCase(Trim(s_data.bs_no)) = UCase(Trim(bs_no)) Then
             Seek #3, 1026 + po * 309
             pget = 1
             Exit Do
          End If
          po = po + 1
       Loop
       If pget = 1 Then
          For i = 1 To 16
              bc(i) = "   "
          Next
          Call getbcch
          Bcchtemp = ""
          For i = 1 To 16
              If Val(bc(i)) > 0 Then
                 Bcchtemp = Bcchtemp & Trim(bc(i)) & ","
              End If
          Next
          If Len(Bcchtemp) > 0 Then
             Bcchtemp = Left(Bcchtemp, Len(Bcchtemp) - 1)
          End If
          s_data.Nonbcch = Bcchtemp
          Put #3, , s_data
       End If
       Close #3
       Open soutput For Binary As #3
       DoEvents
       If Convert_Stop = True Then
          Close
          Exit Sub
       End If
    Loop
    If ProgressBar1.Value < 100 Then
       ProgressBar1.Value = 100
    End If
    Close
End Sub

Sub getbcch()
    Dim bc_line As String
    Dim ff As Integer
    Dim buff As String
    On Error Resume Next
gg:
    Call in_4(bc_line)
    ff = InStr(bc_line, " ")
    buff = Left(bc_line, ff - 1)
    If UCase(buff) <> "CHGR" Then GoTo gg
    For i = 1 To 16
        Call in_4(bc_line)
        ff = InStr(bc_line, " ")
        If ff = 0 Then
           If UCase(bc_line) = "CELL" Or UCase(bc_line) = "END" Then Exit For
           bc(i) = bc_line
           GoTo nextbc
        Else
           buff = Left(bc_line, ff - 1)
           If UCase(buff) = "CELL" Or UCase(bc_line) = "END" Then Exit For
           If UCase(buff) = "WO" Then Call in_4(bc_line)
           If UCase(bc_line) = "CELL" Or UCase(bc_line) = "END" Then Exit For
        End If
        ff = InStr(bc_line, " ")
        If ff = 0 Then
           bc(i) = bc_line
        Else:
           Do
              bc_line = Trim(Right(bc_line, Len(bc_line) - ff))
              ff = InStr(bc_line, " ")
              If ff = 0 Then
                 bc(i) = bc_line
                 Exit Do
              End If
           Loop
        End If
nextbc:
    Next
       
End Sub
Sub getbsno(bs_no)
    Dim bs_line As String
    On Error Resume Next
ag:
    Call in_4(bs_line)
    If bs_line = "" Then
       bs_no = ""
    Else
       If Len(bs_line) > 10 Or UCase(bs_line) = "CELL" Then GoTo ag
       bs_no = bs_line
    End If
End Sub

Sub getline_ericsson(lines)  'Read data from source file(sinput1)
    On Error Resume Next
    Do While Not EOF(1)
s1:
       Line Input #1, lines
       lines = Trim(lines)
       If Len(lines) = 0 And EOF(1) = 0 Then
          GoTo s1
       End If
       finds = InStr(lines, " ")
       If finds = 0 And EOF(1) = 0 Then
          GoTo s1
       End If
    '  lines = LTrim$(Right(lines, Len(lines) - finds))
       lin = Trim(Right(lines, Len(lines) - finds))
       If Mid(lin, 1, 5) = "460-0" Then
          Exit Sub
       End If
    Loop
    lines = ""
End Sub
Sub getin(lines)  'Read data from source file(sinput3)
    On Error Resume Next
    Do While Not EOF(7)
s7:
       Line Input #7, lines
       lines = Trim(lines)
       If Len(lines) = 0 And EOF(7) = 0 Then
          GoTo s7
       End If
       finds = InStr(lines, " ")
       If finds = 0 And EOF(7) = 0 Then
          GoTo s7
       Else
          Exit Sub
       End If
    Loop
    lines = ""
End Sub

Sub in_4(lines)  'Read data from source file(sinput4)
    On Error Resume Next
    Do While Not EOF(7)
s7:
       Line Input #7, lines
       lines = Trim(lines)
       If Len(lines) = 0 And EOF(7) = 0 Then
          GoTo s7
       Else
          Exit Sub
       End If
    Loop
    lines = ""
End Sub

Sub getfield(lines, field)  ' Get field data
    Dim finds As Integer
    Dim FindChar As String * 1
    
    On Error Resume Next
    finds = InStr(lines, Chr(9))
    If finds > 0 Then
       FindChar = Chr(9)
    Else
       FindChar = " "
    End If
    finds = InStr(lines, FindChar)
    If finds = 0 Then
       field = lines
       lines = ""
    Else
       field = Left(lines, finds - 1)
       lines = Trim(Right(lines, Len(lines) - finds))
    End If
End Sub

Sub sts_ericsson()
    Dim Convert_UseName As String
    Dim convert_tabname As String
    Dim convert_num As Integer
    Dim f_cch As e_cch
    Dim f_tch As e_tch
    Dim i As Integer, p As Integer, j As Integer, k As Integer
    Dim cre_name As String
    Dim tch_cch As String, n_name As String
    Dim xls_all
    Dim d_end As String * 1
    Dim recordno As Long, filelong As Long
    Dim per As Integer, PerCount As Integer, finds As Integer
    Dim shz As Long
'*******************************************************8
    Dim row, temp_row, bs_no, lat, lon, bearing, radius
    Dim radiuslon, radiuslat
'*******************************************************
    On Error Resume Next
    convert_num = 1
    Convert_UseName = convert_filename(1)
    finds = InStr(Convert_UseName, ".")
    If finds > 0 Then
       Convert_UseName = Left(Convert_UseName, finds - 1)
    End If
    Err = 0
    convert_tabname = Convert_UseName + ".tab"
    mapinfo.Do "Register Table " + Chr(34) + convert_filename(1) + Chr(34) + " TYPE XLS Into " + Chr(34) + convert_tabname + Chr(34)
    mapinfo.Do "open table " + Chr(34) + convert_tabname + Chr(34)
    If Err Then
       j = MsgBox("无法打开文件   " + convert_filename(1) + " ,或该文件不是 Excel 文件格式", 48, "打开文件")
       Unload Me
       Exit Sub
    End If
    finds = InStr(Convert_UseName, "\")
    Do While finds > 0
       Convert_UseName = Right(Convert_UseName, Len(Convert_UseName) - finds)
       finds = InStr(Convert_UseName, "\")
    Loop
    If Asc(Left(Convert_UseName, 1)) > 47 And Asc(Left(Convert_UseName, 1)) < 58 Then
       Convert_UseName = "_" + Convert_UseName
    End If
    
    d_end = Chr(26)
    ProgressBar1.Value = 1
    For i = 1 To convert_num
        ProgressBar1.Value = 1
        mapinfo.Do "fetch first from " & Convert_UseName
        xls_all = mapinfo.eval("tableinfo(" + Convert_UseName + ",8)")
        per = Int((xls_all / 100) + 0.5)
        PerCount = 0
        mapinfo.Do "fetch next from " & Convert_UseName
        mapinfo.Do "fetch next from " & Convert_UseName
        tch_cch = mapinfo.eval(Convert_UseName + ".col1")
        If tch_cch = "TCH" Then
           cre_name = Gsm_Path + "\sts\tch_sts.dbf"
           Label1.Caption = "正在生成 " + cre_name
           Gsm_FileName = Gsm_Path + "\e_tch.dbf"
           FileCopy Gsm_FileName, cre_name
           filelong = FileLen(Gsm_FileName) + 1
           Open cre_name For Binary As #3
           Seek #3, filelong
           mapinfo.Do "fetch next from " & Convert_UseName
           f_tch.b = " "
           f_tch.TCH = mapinfo.eval(Convert_UseName + ".col1")
           f_tch.ID = mapinfo.eval(Convert_UseName + ".col2")
           f_tch.CA = mapinfo.eval(Convert_UseName + ".col3")
           f_tch.CS = mapinfo.eval(Convert_UseName + ".col4")
           f_tch.CA = LTrim$(str(Int(Val(f_tch.CA) + 0.5)))
           f_tch.CS = LTrim$(str(Int(Val(f_tch.CS) + 0.5)))
           For j = 1 To 11
               k = j + 4
               f_tch.tch_col(j) = mapinfo.eval(Convert_UseName + ".col" & k)
               If j <> 6 And j <> 7 Then
                  f_tch.tch_col(j) = Format$(f_tch.tch_col(j), "fixed")
               Else
                  f_tch.tch_col(j) = LTrim$(str(Int(Val(f_tch.tch_col(j)) + 0.5)))
               End If
           Next
           Put #3, , f_tch
           recordno = 1
           mapinfo.Do "fetch next from " & Convert_UseName
           mapinfo.Do "fetch next from " & Convert_UseName
           For p = 1 To xls_all - 6
               If PerCount = per And ProgressBar1.Value < 90 Then
                  ProgressBar1.Value = ProgressBar1.Value + 1
                  PerCount = 0
               End If
               PerCount = PerCount + 1
               mapinfo.Do "fetch next from " & Convert_UseName
               n_name = mapinfo.eval(Convert_UseName + ".col1")
               If n_name <> "" Then f_tch.TCH = n_name
               f_tch.ID = mapinfo.eval(Convert_UseName + ".col2")
               f_tch.CA = mapinfo.eval(Convert_UseName + ".col3")
               f_tch.CS = mapinfo.eval(Convert_UseName + ".col4")
               f_tch.CA = LTrim$(str(Int(Val(f_tch.CA) + 0.5)))
               f_tch.CS = LTrim$(str(Int(Val(f_tch.CS) + 0.5)))
               For j = 1 To 11
                   k = j + 4
                   f_tch.tch_col(j) = mapinfo.eval(Convert_UseName + ".col" & k)
                   If j <> 6 And j <> 7 Then
                      f_tch.tch_col(j) = Format$(f_tch.tch_col(j), "fixed")
                   Else
                      f_tch.tch_col(j) = LTrim$(str(Int(Val(f_tch.tch_col(j)) + 0.5)))
                   End If
               Next
               recordno = recordno + 1
               Put #3, , f_tch
               DoEvents
               If Convert_Stop = True Then
                  Exit For
               End If
           Next
           Put #3, , d_end
           Seek #3, 5
           Put #3, , recordno
'******************************************************************
           mapinfo.Do "close table " & Convert_UseName
           Close #3
           mapinfo.Do "Register Table  " + " " + Chr(34) + Gsm_Path + "\sts\tch_sts.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + Gsm_Path + "\sts\tch_sts.tab" + Chr(34)
           mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\sts\tch_sts.tab" + Chr(34)
           mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\map\cell" + Chr(34)
           mapinfo.Do "fetch first from tch_sts"
           row = Val(mapinfo.eval("tableinfo(tch_sts,8)"))
           mapinfo.Do "Create Map For tch_sts CoordSys Earth Projection 1, 0 "
           mapinfo.Do "Set Style Pen MakePen(1,60,0)"
           mapinfo.Do "set style brush  makebrush(2,0,0) "
           mapinfo.Do "Set Style Symbol MakeSymbol(33,0,2)"
           For j = 1 To row
               bs_no = mapinfo.eval("tch_sts.col2")
               mapinfo.Do "select * from cell where col2 = " + Chr(34) + bs_no + Chr(34) + " into temp"
               temp_row = Val(mapinfo.eval("tableinfo(temp,8)"))
               If temp_row > 0 Then
                  lon = mapinfo.eval("temp.lon")
                  lat = mapinfo.eval("temp.lat")
                  bearing = mapinfo.eval("temp.bearing")
'                  radius = mapinfo.eval(tch_sts.col6)
'                  radius = Int(radius * 50)
                  lon = lon + 0.0015 * Sin(bearing * 0.01745329252)
                  lat = lat + 0.0015 * Cos(bearing * 0.01745329252)
 '                radiuslon = radius + lon
 '                 radiuslat = radius + lat
                  mapinfo.Do " update tch_sts set lon = " + str(lon) + ",lat = " + str(lat) + ",bearing = " + str(bearing) + " where rowid = " & j
                  mapinfo.Do "create point into variable sts_mypoint (" & lon & "," & lat & ") symbol(34,7585792,2)"
                  mapinfo.Do "update tch_sts set Obj=sts_mypoint  where rowid=" & j
               End If
               mapinfo.Do "fetch next from tch_sts"
               mapinfo.Do "fetch first from cell"
           Next
           mapinfo.Do "commit table tch_sts"
           mapinfo.Do "close table cell"
           mapinfo.Do "close table tch_sts"
'***************************************************************************************
        Else
           cre_name = Gsm_Path + "\sts\cch_sts.dbf"
           Label1.Caption = "正在生成 " + cre_name
           Gsm_FileName = Gsm_Path + "\e_cch.dbf"
           FileCopy Gsm_FileName, cre_name
           filelong = FileLen(Gsm_FileName) + 1
           Open cre_name For Binary As #3
           Seek #3, filelong
           f_cch.b = " "
           recordno = 0
           For p = 1 To xls_all - 3
               If PerCount = per And ProgressBar1.Value < 90 Then
                  ProgressBar1.Value = ProgressBar1.Value + 1
                  PerCount = 0
               End If
               PerCount = PerCount + 1
               mapinfo.Do "fetch next from " & Convert_UseName
               n_name = mapinfo.eval(Convert_UseName + ".col1")
               If n_name <> "" Then f_cch.CCH = n_name
               f_cch.ID = mapinfo.eval(Convert_UseName + ".col2")
               f_cch.CSR = mapinfo.eval(Convert_UseName + ".col3")
               f_cch.CSR = Format$(Val(f_cch.CSR), "fixed")
               f_cch.SA = mapinfo.eval(Convert_UseName + ".col4")
               f_cch.ss = mapinfo.eval(Convert_UseName + ".col5")
               f_cch.SA = LTrim$(str(Int(Val(f_cch.SA) + 0.5)))
               f_cch.ss = LTrim$(str(Int(Val(f_cch.ss) + 0.5)))
               For j = 1 To 6
                   k = j + 5
                   f_cch.cch_col(j) = mapinfo.eval(Convert_UseName + ".col" & k)
                   f_cch.cch_col(j) = Format$(Val(f_cch.cch_col(j)), "fixed")
               Next
               recordno = recordno + 1
               Put #3, , f_cch
               DoEvents
               If Convert_Stop = True Then
                  Exit For
               End If
           Next
           Put #3, , d_end
           Seek #3, 5
           Put #3, , recordno
'******************************************************
           Close #3
           mapinfo.Do "close table " & Convert_UseName
           mapinfo.Do "Register Table  " + " " + Chr(34) + Gsm_Path + "\sts\cch_sts.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + Gsm_Path + "\sts\cch_sts.tab" + Chr(34)
           mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\sts\cch_sts.tab" + Chr(34)
           mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\map\cell" + Chr(34)
           mapinfo.Do "fetch first from cch_sts"
           row = Val(mapinfo.eval("tableinfo(cch_sts,8)"))
           mapinfo.Do "Create Map For cch_sts CoordSys Earth Projection 1, 0 "
           mapinfo.Do "Set Style Pen MakePen(1,60,0)"
           mapinfo.Do "set style brush  makebrush(2,0,0) "
           mapinfo.Do "Set Style Symbol MakeSymbol(33,0,2)"
           For j = 1 To row
               bs_no = mapinfo.eval("cch_sts.col2")
               mapinfo.Do "select * from cell where col2 = " + Chr(34) + bs_no + Chr(34) + " into temp"
               temp_row = Val(mapinfo.eval("tableinfo(temp,8)"))
               If temp_row > 0 Then
                  lon = mapinfo.eval("temp.lon")
                  lat = mapinfo.eval("temp.lat")
                  bearing = mapinfo.eval("temp.bearing")
                  lon = lon + 0.0015 * Sin(bearing * 0.01745329252)
                  lat = lat + 0.0015 * Cos(bearing * 0.01745329252)
                  mapinfo.Do " update cch_sts set lon = " + str(lon) + ",lat = " + str(lat) + ",bearing = " + str(bearing) + " where rowid = " & j
                  mapinfo.Do "create point into variable sts_mypoint (" & lon & "," & lat & ") symbol(34,7585792,2)"
                  mapinfo.Do "update cch_sts set Obj=sts_mypoint  where rowid=" & j
               End If
               mapinfo.Do "fetch next from cch_sts"
               mapinfo.Do "fetch first from cell"
           Next
           mapinfo.Do "commit table cch_sts"
           mapinfo.Do "close table cell"
           mapinfo.Do "close table cch_sts"
'*******************************************************
        End If
'        mapinfo.do "Register Table  " + " " + Chr(34) + "c:\gsm\sts\tch_sts.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + "c:\gsm\sts\tch_sts.tab" + Chr(34)
'        mapinfo.do "Register Table  " + " " + Chr(34) + "c:\gsm\sts\cch_sts.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + "c:\gsm\sts\cch_sts.tab" + Chr(34)
    Next
End Sub

Private Sub Command1_Click()
    On Error Resume Next

    If (MsgBox("确实要中止转换吗？", 33, "提示")) = 1 Then
       Convert_Stop = True
    End If
End Sub

Private Sub Timer1_Timer()
    Dim sinput1 As String, sinput2 As String, sinput3 As String
    Dim sinput4 As String, soutput As String
    Dim gsmname As String
    Dim finds As Integer
    Dim MyRecord As Record
    Dim temp As String
    Dim Cell_Rows As Integer
    Dim Suffix As String
    Dim InString As String
    Dim NcellFile As Boolean
    Dim MyInString As String * 200
    
    On Error Resume Next
    NcellFile = False
    Convert_Stop = False
    Timer1.Enabled = False
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    Open Gsm_FileName For Binary As #1
    Get #1, 1, MyRecord  ' Read third record.
    Close #1
    If Menu_Flag = 9999 Or Menu_Flag = 9998 Or Menu_Flag = 9997 Then
       UpdateCell
       Unload Me
       Exit Sub
    End If
    If Menu_Flag = 2301 Then
       If UCase(Right(convert_filename(1), 4)) <> ".XLS" Then
          StsEricssonHex
       Else
          sts_ericsson
       End If
       Unload Me
    End If
    If Menu_Flag = 2302 Then
       save_mark = False
       If UCase(Right(convert_filename(1), 4)) = ".XLS" Then
          mapinfo.Do "Register Table " + Chr(34) + convert_filename(1) + Chr(34) + " TYPE XLS Into " + Chr(34) + Gsm_Path + "\CellTemp.tab" + Chr(34)
          mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\CellTemp.tab" + Chr(34)
          If Err Then
             MsgBox "无法打开文件 " & convert_filename(1) & "或文件格式错误", 64, "提示"
             Screen.MousePointer = 0
             Unload Me
             Exit Sub
          End If
          If (Val(mapinfo.eval("tableinfo(celltemp,4)")) = 29 And UCase(mapinfo.eval("Columninfo( celltemp,COL1, 1)")) = "CELL_NAME") Then
             CellExcel
             Screen.MousePointer = 0
             Unload Me
             Exit Sub
          Else
             If Val(MyRecord.exchange) <> 1 Then
                mapinfo.Do "close table celltemp"
                Gsm_FileName = Gsm_Path + "\celltemp.*"
                Kill Gsm_FileName
             End If
          End If
       End If
       
       If Val(MyRecord.exchange) = 0 Then
         If InStr(UCase(convert_filename(1)), "RLNRP") > 0 Then
            Ncell_Ericsson
            save_mark = False
         Else
          sinput = ""
          temp = convert_filename(1)
          finds = InStr(temp, "\")
          Do While finds > 0
             sinput = sinput + Left(temp, finds)
             temp = Right(temp, Len(temp) - finds)
             finds = InStr(temp, "\")
          Loop
          sinput = Left(sinput, Len(sinput) - 1)
          finds = InStr(temp, ".")
          If finds > 0 Then
             Suffix = Right(temp, Len(temp) - finds)
          Else
             Suffix = ""
          End If
          sinput1 = sinput + "\rldep" + "." + Suffix
          If UCase(dir(sinput1, 0)) = "" Then
             MsgBox "数据文件RLDEP不存在！", 64, "提示"
             Unload Me
             Exit Sub
          End If
          soutput = Gsm_Path + "\map\cell1.dbf"
          sinput2 = Gsm_Path + "\map\cell.dbf"
          sinput3 = sinput + "\rlcpp" + "." + Suffix
          If UCase(dir(sinput3, 0)) = "" Then
             MsgBox "数据文件RLCPP不存在！", 64, "提示"
             Unload Me
             Exit Sub
          End If
          sinput4 = sinput + "\rlcfp" + "." + Suffix
          If UCase(dir(sinput4, 0)) = "" Then
             MsgBox "数据文件RLCFP不存在！", 64, "提示"
             Unload Me
             Exit Sub
          End If
          save_mark = False
          Call cell_ericsson(sinput1, sinput2, sinput3, sinput4, soutput)
          FileCopy soutput, sinput2
         End If
          mapinfo.Do "Register Table  " + " " + Chr(34) + Gsm_Path + "\map\cell.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + Gsm_Path + "\map\cell.tab" + Chr(34)
          mapinfo.Do "Open Table " + Chr(34) + Gsm_Path + "\map\cell.tab" + Chr(34)
          Screen.MousePointer = 0
          Unload Me
          If load_new = 0 Then
             MsgBox "CELL库数据已全部更新!", 64, "提示"
          Else
             MsgBox "CELL库数据已更新了" & load_sam & "项!" + Chr(10) + "交换数据中还有" & load_new & "项在CELL库中没有" + Chr(10) + "相同的BS_NO与它们相匹配！", 64, "提示"
          End If
          mapinfo.Do "close table cell"
       Else
          If Val(MyRecord.exchange) = 4 Then
                Open convert_filename(1) For Input As #8
                Do While Not EOF(8)
                   Line Input #8, InString
                   If InStr(UCase(InString), "<<") > 0 Then
                      If InStr(UCase(InString), "IADCEL") > 0 Then
                         NcellFile = True
                      Else
                         NcellFile = False
                      End If
                      Exit Do
                   End If
                Loop
                Close #8
             If NcellFile Then
                Ncell_Italtel
             Else
                Cell_Italtel
             End If
             save_mark = False
             mapinfo.Do "Register Table  " + " " + Chr(34) + Gsm_Path + "\map\cell.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + Gsm_Path + "\map\cell.tab" + Chr(34)
             mapinfo.Do "Open Table " + Chr(34) + Gsm_Path + "\map\cell.tab" + Chr(34)
             Cell_Rows = Val(mapinfo.eval("tableinfo(cell,8)"))
             Screen.MousePointer = 0
             Unload Me
             If Cell_Rows = load_new Then
                MsgBox "CELL库数据已全部更新!", 64, "提示"
             Else
                MsgBox "CELL库数据已更新了" & load_new & "项!" + Chr(10) + "CELL库中还有" & (Cell_Rows - load_new) & "项没被更新", 64, "提示"
             End If
             mapinfo.Do "close table cell"
          ElseIf Val(MyRecord.exchange) = 5 Then
                Open convert_filename(1) For Binary As #8
                Get #8, , MyInString
                InString = Trim(Left(MyInString, InStr(MyInString, Chr(10)) - 1))
                Suffix = ","
                finds = InStr(InString, Suffix)
                If finds = 0 Then
                   Suffix = Chr(9)
                   finds = InStr(InString, Suffix)
                   If finds = 0 Then
                      Suffix = Chr(20)
                      finds = InStr(InString, Suffix)
                      If finds = 0 Then
                         Close #8
                         GoTo incorrectType
                      End If
                   End If
                End If
                Do While Right(InString, 1) = Suffix Or Right(InString, 1) = Chr(10) Or Right(InString, 1) = Chr(13)
                   InString = Trim(Left(InString, Len(InString) - 1))
                Loop
                InString = Left(InString, Len(InString) - 1)
                If Right(InString, 1) = Suffix Then
                   NcellFile = False
                Else
                   NcellFile = True
                End If
                Close #8
                If NcellFile Then
                   NCell_Nortel
                Else
                   Cell_Nortel
                End If
incorrectType:
                save_mark = False
                mapinfo.Do "Register Table  " + " " + Chr(34) + Gsm_Path + "\map\cell.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + Gsm_Path + "\map\cell.tab" + Chr(34)
                mapinfo.Do "Open Table " + Chr(34) + Gsm_Path + "\map\cell.tab" + Chr(34)
                Cell_Rows = Val(mapinfo.eval("tableinfo(cell,8)"))
                Screen.MousePointer = 0
                Unload Me
                If Cell_Rows = load_new Then
                   MsgBox "CELL库数据已全部更新!", 64, "提示"
                Else
                   MsgBox "CELL库数据已更新了" & load_new & "项!" + Chr(10) + "CELL库中还有" & (Cell_Rows - load_new) & "项没被更新", 64, "提示"
                End If
                mapinfo.Do "close table cell"
          Else
             save_mark = False
             If UCase(Right(convert_filename(1), 4)) = ".XLS" Then
                CellMotorolaXLS
             Else
                sinput1 = convert_filename(1)
                sinput2 = Gsm_Path + "\map\cell.dbf"
                Open sinput1 For Input As #8
                Do While Not EOF(8)
                   Line Input #8, InString
                   If InStr(UCase(InString), "BSIC") > 0 Then
                      If InStr(UCase(InString), "NEIGHB") > 0 Then
                         NcellFile = True
                      Else
                         NcellFile = False
                      End If
                      Exit Do
                   End If
                Loop
                Close #8
                If NcellFile Then
                   ncell_motorola
                Else
                   Call cell_motorola(sinput1, sinput2)
                End If
             End If
             mapinfo.Do "Register Table  " + " " + Chr(34) + Gsm_Path + "\map\cell.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + Gsm_Path + "\map\cell.tab" + Chr(34)
             mapinfo.Do "Open Table " + Chr(34) + Gsm_Path + "\map\cell.tab" + Chr(34)
             Cell_Rows = mapinfo.eval("tableinfo(cell,8)")
             Screen.MousePointer = 0
             Unload Me
             If Cell_Rows = load_new Then
                MsgBox "CELL库数据已全部更新!", 64, "提示"
             Else
                MsgBox "CELL库数据已更新了" & load_new & "项!" + Chr(10) + "CELL库中还有" & (Cell_Rows - load_new) & "项没被更新", 64, "提示"
             End If
             mapinfo.Do "close table cell"
             Exit Sub
          End If
       End If
    End If
    If Menu_Flag = 2303 Then
       If Val(MyRecord.exchange) = 0 Then
          Ncell_Ericsson
       Else
          If Val(MyRecord.exchange) = 4 Then
             Ncell_Italtel
          Else
             If UCase(Right(convert_filename(1), 4)) = ".XLS" Then
                NcellMotorolaXLS
             Else
                ncell_motorola
             End If
          End If
       End If
       Unload Me
    End If
End Sub

Sub Cell_Italtel()
    Dim Old_Cell As NewCellStru
    Dim recordno As Long
    Dim Line_Char As String
    Dim End_Char As String * 1
    Dim i As Integer, File_num As Integer, j As Integer
    Dim FileLenth As Long, nline As Integer, bline As Integer
    Dim PercentStep As Integer, bs As Integer, scnline As Integer
    Dim my_Pos As Long
    
    On Error Resume Next
    File_num = 0
    Do While Trim(convert_filename(File_num + 1)) <> ""
       File_num = File_num + 1
    Loop
    End_Char = Chr$(26)
    load_new = 0
    Gsm_FileName = Gsm_Path + "\map\cell.dbf"
    Open Gsm_FileName For Binary As #2
    For i = 1 To File_num
        FileLenth = FileLen(convert_filename(i))
        nline = FileLenth / 3000
        bline = Fix(nline / 100)
        PercentStep = 1
        If bline = 0 Then
           bline = 1
           PercentStep = 100 / nline
        End If
        bs = 1
        scnline = 0
        ProgressBar1.Value = 0
        Label1.Caption = "正在转换 " + convert_filename(i)
        Label1.Refresh
        end_frag = False
        Open convert_filename(i) For Input As #1
        Have_Read = False
        Do While 1
           New_Cell.ci = space(5)
           New_Cell.BSIC = space(3)
           New_Cell.ARFCN = space(3)
           New_Cell.bs_no = space(10)
           New_Cell.max_bts = space(2)
           New_Cell.max_ms = space(2)
           New_Cell.Nonbcch = space(32)
           scnline = scnline + 1
           If scnline = bs * bline And ProgressBar1.Value < 99 Then
              ProgressBar1.Value = ProgressBar1.Value + PercentStep
              bs = bs + 1
           End If
           Read_Cell
           If end_frag = True Then
              Exit Do
           End If
           my_Pos = 0
           Do While Not EOF(2)
              Seek #2, 1026 + my_Pos * 309
              Get #2, , Old_Cell
              If Trim(UCase(Old_Cell.bs_no)) = Trim(New_Cell.bs_no) Then
                 Other_Field
                 If end_frag = True Then
                    Exit Do
                 End If
                 If Len(Italtel_Nonbcch) > 0 Then
                    Italtel_Nonbcch = Left(Italtel_Nonbcch, Len(Italtel_Nonbcch) - 1)
                 End If
                 Old_Cell.ARFCN = New_Cell.ARFCN
                 Old_Cell.Lac = New_Cell.Lac
                 Old_Cell.ci = New_Cell.ci
                 Old_Cell.BSIC = New_Cell.BSIC
                 Old_Cell.max_bts = New_Cell.max_bts
                 Old_Cell.time = DATE
                 Old_Cell.Nonbcch = Italtel_Nonbcch
                 Seek #2, 1026 + my_Pos * 309
                 Put #2, , Old_Cell
                 load_new = load_new + 1
                 Exit Do
              End If
              my_Pos = my_Pos + 1
           Loop
           If end_frag = True Then
              Exit Do
           End If
           DoEvents
           If Convert_Stop = True Then
              Close
              Exit Sub
           End If
           Close #2
           Open Gsm_FileName For Binary As #2
        Loop
        Close #1
        If ProgressBar1.Value < 100 Then
           ProgressBar1.Value = 100
        End If
    Next
    Close
End Sub

Sub Read_Cell()
    Dim Read_Line As String
    Dim finds As Integer
    
    On Error Resume Next
    If Have_Read = True Then
       Have_Read = False
       Read_Line = Public_Line
       GoTo Find_again
    End If
Read_again:
    Do While Not EOF(1)
       Line Input #1, Read_Line
       If InStr(Read_Line, "<<") > 0 Then
          Exit Do
       End If
    Loop
    If EOF(1) Then
       end_frag = True
       Exit Sub
    End If
Find_again:
    Read_Line = UCase(Trim(Read_Line))
    finds = InStr(Read_Line, "IIDCEL")
    If finds > 0 Then
       Read_Line = Right(Read_Line, Len(Read_Line) - finds)
       finds = InStr(Read_Line, ":")
       If finds > 0 Then
          Read_Line = Right(Read_Line, Len(Read_Line) - finds)
          finds = InStr(Read_Line, ";")
          If finds > 0 Then
             Read_Line = Left(Read_Line, finds - 1)
             New_Cell.bs_no = Trim(Read_Line)
          Else
             GoTo Read_again
          End If
       Else
          GoTo Read_again
       End If
    Else
       If EOF(1) Then
          end_frag = True
          Exit Sub
       End If
       Line Input #1, Read_Line
       GoTo Find_again
    End If
End Sub

Sub Other_Field()
    Dim Read_Line As String
    Dim finds As Integer
    Dim ncc As String, bscc As String
    Dim DchNo_num As Integer
    Dim i As Integer
    Dim OneChar As String * 1
    
    On Error Resume Next
    Italtel_Nonbcch = ""
    Do While Not EOF(1)
       Line Input #1, Read_Line
       Read_Line = UCase(Trim(Read_Line))
       finds = InStr(Read_Line, "LAC:")
       If finds > 0 Then
          Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
          finds = InStr(Read_Line, ":")
          Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
          finds = InStr(Read_Line, " ")
          If finds > 0 Then
             New_Cell.Lac = Trim(Left(Read_Line, finds - 1))
          End If
          finds = InStr(Read_Line, "CCID:")
          If finds > 0 Then
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, ":")
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, " ")
             If finds > 0 Then
                New_Cell.ci = Trim(Left(Read_Line, finds - 1))
             End If
          End If
          Exit Do
       End If
    Loop
    If EOF(1) Then
       end_frag = True
       Exit Sub
    End If
    Do While Not EOF(1)
       Line Input #1, Read_Line
       Read_Line = UCase(Trim(Read_Line))
       finds = InStr(Read_Line, "NCC:")
       If finds > 0 Then
          Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
          finds = InStr(Read_Line, ":")
          Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
          finds = InStr(Read_Line, " ")
          If finds > 0 Then
             ncc = Trim(Left(Read_Line, finds - 1))
          End If
          finds = InStr(Read_Line, "BSCC:")
          If finds > 0 Then
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, ":")
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, " ")
             If finds > 0 Then
                bscc = Trim(Left(Read_Line, finds - 1))
             End If
          End If
          New_Cell.BSIC = ncc & bscc
          Exit Do
       End If
    Loop
    If EOF(1) Then
       end_frag = True
       Exit Sub
    End If
    Do While Not EOF(1)
       Line Input #1, Read_Line
       Read_Line = UCase(Trim(Read_Line))
       finds = InStr(Read_Line, "BSPWR:")
       If finds > 0 Then
          Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
          finds = InStr(Read_Line, ":")
          Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
          finds = InStr(Read_Line, " ")
          If finds > 0 Then
             New_Cell.max_bts = Trim(Left(Read_Line, finds - 1))
          End If
          finds = InStr(Read_Line, "BCCH:")
          If finds > 0 Then
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, ":")
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, " ")
             If finds > 0 Then
                New_Cell.ARFCN = Trim(Left(Read_Line, finds - 1))
             End If
          End If
          Exit Do
       End If
    Loop
'************************************* DCHNO
    DchNo_num = 0
    Do While Not EOF(1)
       Line Input #1, Read_Line
       Read_Line = UCase(Trim(Read_Line))
       finds = InStr(Read_Line, "CELL")
       'If finds > 0 And InStr(Read_Line, "ALLOCATION") > 0 Then
       If finds > 0 And InStr(Read_Line, "ALL") > 0 Then
          Line Input #1, Read_Line
          Read_Line = Trim(Read_Line)
Dchno1:
          For i = 1 To Len(Read_Line)
              If Asc(Left(Read_Line, 1)) > 47 And Asc(Left(Read_Line, 1)) < 58 Then
                 Exit For
              Else
                 Read_Line = Trim(Right(Read_Line, Len(Read_Line) - 1))
                 If Len(Read_Line) <= 1 Then
                    Exit Do
                 End If
              End If
          Next
          finds = InStr(Read_Line, " ")
          If finds > 0 Then
             'New_Cell.bcch(DchNo_num + 1) = Trim(Left(Read_Line, finds - 1))
             Italtel_Nonbcch = Italtel_Nonbcch & Trim(Left(Read_Line, finds - 1)) & ","
             DchNo_num = DchNo_num + 1
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             GoTo Dchno1
             finds = InStr(Read_Line, " ")
          Else
             'New_Cell.bcch(DchNo_num + 1) = Read_Line
             Italtel_Nonbcch = Italtel_Nonbcch & Trim(Read_Line) & ","
             DchNo_num = DchNo_num + 1
          End If
       End If
    Loop
    Do While Not EOF(1)
       Line Input #1, Read_Line
       Read_Line = UCase(Trim(Read_Line))
       If InStr(Read_Line, "<<") > 0 Then
          Have_Read = True
          Public_Line = Read_Line
          Exit Do
       End If
       finds = InStr(Read_Line, "CELL")
       'If finds > 0 And InStr(Read_Line, "ALLOCATION") > 0 Then
       If finds > 0 And InStr(Read_Line, "ALL") > 0 Then
          Line Input #1, Read_Line
          Read_Line = Trim(Read_Line)
Dchno2:
          For i = 1 To Len(Read_Line)
              If Asc(Left(Read_Line, 1)) > 47 And Asc(Left(Read_Line, 1)) < 58 Then
                 Exit For
              Else
                 Read_Line = Trim(Right(Read_Line, Len(Read_Line) - 1))
                 If Len(Read_Line) <= 1 Then
                    GoTo Dchno3
                 End If
              End If
          Next
          finds = InStr(Read_Line, " ")
          If finds > 0 Then
             'New_Cell.bcch(DchNo_num + 1) = Trim(Left(Read_Line, finds - 1))
             Italtel_Nonbcch = Italtel_Nonbcch & Trim(Left(Read_Line, finds - 1)) & ","
             DchNo_num = DchNo_num + 1
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             GoTo Dchno2
             finds = InStr(Read_Line, " ")
          Else
             'New_Cell.bcch(DchNo_num + 1) = Read_Line
             Italtel_Nonbcch = Italtel_Nonbcch & Trim(Read_Line) & ","
             DchNo_num = DchNo_num + 1
          End If
       End If
Dchno3:
    Loop
    If EOF(1) Then
       end_frag = True
    End If
End Sub

Sub Ncell_Italtel()
    Dim recordno As Integer, File_num As Integer, i As Integer
    Dim linetxt As String
    Dim End_Char As String * 1
    Dim File_Lenth As Long, nline As Integer, bline As Integer
    Dim percent_step As Integer, bss As Integer, scnline As Integer
    Dim j As Integer
    Dim s_data As NewCellStru
    
    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\map\cell.dbf"
    Open Gsm_FileName For Binary As #2

    File_num = 0
    Do While Trim(convert_filename(File_num + 1)) <> ""
       File_num = File_num + 1
    Loop
    For i = 1 To File_num
        Open convert_filename(i) For Input As #1
        File_Lenth = FileLen(convert_filename(i))
        nline = File_Lenth / 5000
        bline = Int(nline / 100)
        percent_step = 2
        If bline = 0 Then
           bline = 1
           percent_step = Int(100 / nline + 0.5) + 1
        End If
        bss = 1
        scnline = 0
        ProgressBar1.Value = 0
        Label1.Caption = "正在转换 " + convert_filename(i)
        end_frag = False
        Get_Bs_no = False
        Do While 1
           scnline = scnline + 1
           If scnline = bss * bline And ProgressBar1.Value < 98 Then
              ProgressBar1.Value = ProgressBar1.Value + percent_step
              bss = bss + 1
           End If
           If Get_Bs_no = False Then
              Read_Ncell
           End If
           If end_frag = True Then
              Exit Do
           End If
           'Call bs_lac(Trim(New_Ncell.bs_no), New_Ncell.bs_name, New_Ncell.ci, New_Ncell.Lac)
           For j = 1 To 16
               New_Ncell.NCELL(j) = ""
           Next
           For j = 1 To 16
               Call Ncell_Field(j)
               If Get_Bs_no = True Then
                  Exit For
               End If
           Next
           
       Pos = 0
       Do While Not EOF(2)
          'Seek #2, 962 + pos * 153
          Seek #2, 1026 + Pos * 309
          Get #2, , s_data
          If UCase(Trim(s_data.bs_no)) = UCase(Trim(New_Ncell.bs_no)) Then
             load_sam = load_sam + 1
             For j = 1 To 16
                 s_data.NCELL(j) = New_Ncell.NCELL(j)
             Next
             s_data.time = DATE
             Seek #2, 1026 + Pos * 309
             Put #2, , s_data
             load_new = load_new + 1
             Exit Do
          End If
          Pos = Pos + 1
       Loop
       Close #2
       Open Gsm_FileName For Binary As #2
           
           'Put #2, , New_Ncell
           DoEvents
           If Convert_Stop = True Then
              Close
              Exit Sub
           End If
           If Get_Bs_no = True Then
              New_Ncell.bs_no = Prev_Bs_no
           End If
        Loop
        If ProgressBar1.Value < 100 Then
           ProgressBar1.Value = 100
        End If
        Close #1
    Next
    Close
End Sub

Sub Read_Ncell()
    Dim Read_Line As String
    Dim finds As Integer
    
    On Error Resume Next
Read_again:
    Do While Not EOF(1)
       Line Input #1, Read_Line
       If InStr(Read_Line, "<<") > 0 Then
          Exit Do
       End If
    Loop
    If EOF(1) Then
       end_frag = True
       Exit Sub
    End If
Find_again:
    Read_Line = UCase(Trim(Read_Line))
    finds = InStr(Read_Line, "IADCEL")
    If finds > 0 Then
       Read_Line = Right(Read_Line, Len(Read_Line) - finds)
       finds = InStr(Read_Line, ":")
       If finds > 0 Then
          Read_Line = Right(Read_Line, Len(Read_Line) - finds)
          finds = InStr(Read_Line, ";")
          If finds > 0 Then
             Read_Line = Left(Read_Line, finds - 1)
             New_Ncell.bs_no = Trim(Read_Line)
          Else
             GoTo Read_again
          End If
       Else
          GoTo Read_again
       End If
    Else
       If EOF(1) Then
          end_frag = True
          Exit Sub
       End If
       Line Input #1, Read_Line
       GoTo Find_again
    End If

End Sub

Sub Ncell_Field(col_num As Integer)
    Dim Read_Line As String
    Dim finds As Integer
    Dim ncc As String, bscc As String
    
    On Error Resume Next
    Get_Bs_no = False
    Do While Not EOF(1)
       Line Input #1, Read_Line
       Read_Line = UCase(Trim(Read_Line))
       
       finds = InStr(Read_Line, "<<")
       If finds > 0 Then
Find_again:
          Read_Line = UCase(Trim(Read_Line))
          finds = InStr(Read_Line, "IADCEL")
          If finds > 0 Then
             Read_Line = Right(Read_Line, Len(Read_Line) - finds)
             finds = InStr(Read_Line, ":")
             If finds > 0 Then
                Read_Line = Right(Read_Line, Len(Read_Line) - finds)
                finds = InStr(Read_Line, ";")
                If finds > 0 Then
                   Read_Line = Left(Read_Line, finds - 1)
                   Prev_Bs_no = Trim(Read_Line)
                   Get_Bs_no = True
                   Exit Sub
                Else
                   If EOF(1) Then
                      end_frag = True
                      Exit Sub
                   End If
                   Line Input #1, Read_Line
                   GoTo Find_again
                End If
             Else
                If EOF(1) Then
                   end_frag = True
                   Exit Sub
                End If
                Line Input #1, Read_Line
                GoTo Find_again
             End If
          Else
             If EOF(1) Then
                end_frag = True
                Exit Sub
             End If
             Line Input #1, Read_Line
             GoTo Find_again
          End If
       End If
       finds = InStr(Read_Line, "CCID:")
       If finds > 0 Then
          Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
          finds = InStr(Read_Line, ":")
          Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
          finds = InStr(Read_Line, " ")
          If finds > 0 Then
             New_Ncell.NCELL(col_num) = Trim(Left(Read_Line, finds - 1))
          End If
          Exit Do
       End If
    Loop
    If EOF(1) Then
       end_frag = True
       Exit Sub
    End If
    Do While Not EOF(1)
       Line Input #1, Read_Line
       Read_Line = UCase(Trim(Read_Line))
       finds = InStr(Read_Line, "NCC:")
       If finds > 0 Then
          Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
          finds = InStr(Read_Line, ":")
          Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
          finds = InStr(Read_Line, " ")
          If finds > 0 Then
             ncc = Trim(Left(Read_Line, finds - 1))
          End If
          finds = InStr(Read_Line, "BSCC:")
          If finds > 0 Then
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, ":")
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, " ")
             If finds > 0 Then
                bscc = Trim(Left(Read_Line, finds - 1))
             End If
          End If
          'New_Ncell.col(col_num).bsic_c = ncc & bscc
          finds = InStr(Read_Line, "BCCH:")
          If finds > 0 Then
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, ":")
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, " ")
             If finds > 0 Then
           '     New_Ncell.col(col_num).arfcn_c = Trim(Left(Read_Line, finds - 1))
             End If
          End If
          Exit Do
       End If
    Loop
    If EOF(1) Then
       end_frag = True
    End If
End Sub

Sub cell_motorola(sinput1, sinput2) 'Convert program
    Dim field_data As NewCellStru
    Dim s_data As NewCellStru
    Dim recordno As Long
    Dim lines As String
    Dim test As String * 1
    Dim lee As String * 1
    Dim cilac As String
    Dim sameci As String * 5
    Dim newbcch(1 To 6) As String * 3
    Dim FLAG As Integer
    Dim NewLac As String, NewCi As String, NewBsic As String, NewArfcn As String
    Dim NewMax_ms As String, NewMax_bts As String, NewPref As String, NewMicrocell As String
    Dim NonBcchtemp As String, NonBcch_one As String
    
    On Error Resume Next
    Pos = 0
    wrec = 0
    lee = Chr$(26)
    FLAG = 0
    load_new = 0
    lenth = FileLen(sinput1)
    nline = lenth / 220
    bline = Fix(nline / 100)
    percent_step = 1
    If bline = 0 Then
       bline = 1
       percent_step = 100 / nline
    End If
    bs = 1
    scnline = 0
    ProgressBar1.Value = 0
    Label1.Caption = "正在转换 " + sinput1
    Label1.Refresh
    Open sinput1 For Input As #1
    Open sinput2 For Binary As #2

    Seek #2, 5
    Get #2, , recordno

    Do While 1
       field_data.b = space(1)
       field_data.time = space(8)
       field_data.lon = space(12)
       field_data.lat = space(12)
       field_data.Name = space(10)
       field_data.bs_no = space(10)
       field_data.bearing = space(3)
       field_data.downtilt = space(3)
       field_data.max_bts = space(2)
       field_data.max_ms = space(2)
       field_data.ci = space(5)
       field_data.ARFCN = space(3)
       field_data.BSIC = space(3)
       field_data.Lac = space(5)
       field_data.Nonbcch = space(32)
       For i = 1 To 16
           field_data.NCELL(i) = space(10)
       Next

       scnline = scnline + 1
       If scnline = bs * bline And ProgressBar1.Value < 99 Then
          ProgressBar1.Value = ProgressBar1.Value + percent_step
          bs = bs + 1
       End If

       If FLAG = 0 Then
          Call getline_motorola(lines)
          If lines = "" Then
             Exit Do
          End If
          Call getfield(lines, cilac)
       End If
       If lines = "" Then
          Exit Do
       End If
       For i = 1 To 2
           finds = InStr(cilac, "-")
           cilac = Right(cilac, Len(cilac) - finds)
       Next
       finds = InStr(cilac, "-")
       field_data.Lac = Left(cilac, finds - 1)
       field_data.ci = Right(cilac, Len(cilac) - finds)
       Call getfield(lines, field_data.BSIC)
       For i = 1 To 2
           Call getfield(lines, nonuse)
       Next
       Call getfield(lines, field_data.ARFCN)
       Call getfield(lines, field_data.max_ms)
       Call getfield(lines, nonuse)
       Call getfield(lines, field_data.max_bts)
       Call getfield(lines, nonuse)
       Call getfield(lines, nonuse)
       'Call getfield(lines, field_data.pref)
       For i = 1 To 10
           Call getfield(lines, nonuse)
       Next
       Call getfield(lines, field_data.microcell)
       
       NewLac = field_data.Lac
       NewCi = field_data.ci
       NewBsic = field_data.BSIC
       NewArfcn = field_data.ARFCN
       NewMax_ms = field_data.max_ms
       NewMax_bts = field_data.max_bts
       'NewPref = field_data.pref
       NewMicrocell = field_data.microcell
        
       j = 1
       FLAG = 0
       NonBcchtemp = ""
       Do While 1

          Call getline_motorola(lines)
          If lines = "" Then
             Exit Do
          End If
          Call getfield(lines, cilac)
          cci = cilac
          For i = 1 To 3
              finds = InStr(cci, "-")
              cci = Right(cci, Len(cci) - finds)
          Next
          sameci = cci
          If Trim(sameci) = Trim(field_data.ci) Then
             For i = 1 To 3
                 Call getfield(lines, nonuse)
             Next
             If j <= 16 Then
                Call getfield(lines, NonBcch_one)
                'newbcch(j) = field_data.bcch(j)
                If NonBcch_one <> "" Then
                   NonBcchtemp = NonBcchtemp & Trim(NonBcch_one) & ","
                   j = j + 1
                End If
             End If
          Else
             FLAG = 1
             Exit Do
          End If
       Loop
       If Len(NonBcchtemp) > 0 Then
          NonBcchtemp = Left(NonBcchtemp, Len(NonBcchtemp) - 1)
       End If
       Pos = 0
       Do While Not EOF(2)
          Seek #2, 1026 + Pos * 309
          Get #2, , s_data
          If s_data.ci = field_data.ci Then
             field_data = s_data
             field_data.Lac = Trim(NewLac)
             field_data.ci = Trim(NewCi)
             field_data.BSIC = Oct(Val(NewBsic))
             field_data.ARFCN = Trim(NewArfcn)
             field_data.max_ms = Trim(NewMax_ms)
             field_data.max_bts = Trim(NewMax_bts)
             'field_data.pref = Trim(NewPref)
             field_data.microcell = Trim(NewMicrocell)
             field_data.time = DATE
             field_data.Nonbcch = NonBcchtemp

             Seek #2, 1026 + Pos * 309
             Put #2, , field_data
             load_new = load_new + 1
             GoTo s2
          End If
          Pos = 1 + Pos
       Loop
       Close #2
       Open sinput2 For Binary As #2

s2:
       DoEvents
       If Convert_Stop = True Then
          Close
          Exit Sub
       End If
    Loop

    Close

    If ProgressBar1.Value < 100 Then
       ProgressBar1.Value = 100
    End If

End Sub

Sub getline_motorola(lines)  'Read data from source file
    Dim finds As Integer
    Dim FindChar As String * 1
    On Error Resume Next
    
    Do While Not EOF(1)
s1:
       Line Input #1, lines
       lines = Trim(lines)
       If Len(lines) = 0 And EOF(1) = 0 Then
          GoTo s1
       End If
       finds = InStr(lines, Chr(9))
       If finds > 0 Then
          FindChar = Chr(9)
          lines = Trim(Right(lines, Len(lines) - finds))
       Else
          FindChar = " "
       End If
       finds = InStr(lines, FindChar)
       If finds = 0 And EOF(1) = False Then
          GoTo s1
       End If
       If EOF(1) Then
          Exit Do
       End If
       lin = Trim(Left(lines, finds - 1))
       If Mid(lin, 1, 5) = "460-0" Then
          Exit Sub
       End If
    Loop
    lines = ""
End Sub

Sub getnb(linetxt, a2)
    Dim finds As Integer
    Dim FindChar As String * 1
    Dim a1 As String, a3 As String
    
    On Error Resume Next
    finds = InStr(linetxt, Chr(9))
    If finds > 0 Then
       FindChar = Chr(9)
    Else
       FindChar = " "
    End If
    finds = InStr(linetxt, FindChar)
    nbci = Left(linetxt, finds - 1)
    a2 = Right(nbci, 4)
    a2 = (a2)
    linetxt = Trim(Right(linetxt, Len(linetxt) - finds))
    finds = InStr(linetxt, FindChar)
    a3 = Left(linetxt, finds - 1)
    linetxt = Trim(Right(linetxt, Len(linetxt) - finds))
    For i = 1 To 4
       finds = InStr(linetxt, FindChar)
       linetxt = Trim(Right(linetxt, Len(linetxt) - finds))
    Next
    finds = InStr(linetxt, FindChar)
    a1 = Left(linetxt, finds)
End Sub

Sub StsEricssonHex()
    Dim i As Integer, FileNums As Integer
    Dim LastString As String, FunctionStr As String
    Dim CorrectFile As Boolean
    Dim ConvertTch As Boolean, ConvertCch As Boolean
    Dim StsRows As Integer
        
    On Error Resume Next
    
    ConvertTch = False
    ConvertCch = False
    FileNums = 1
    FileNumber3 = FreeFile
    Do While Trim(convert_filename(FileNums)) <> ""
       CorrectFile = False
       Open convert_filename(FileNums) For Input As #FileNumber3
       ProgressBar1.Value = 1
       LinesPer = Int(FileLen(convert_filename(FileNums)) / 350 / 100)
       StsPercentStep = 1
       If LinesPer = 0 Then
          LinesPer = 1
          StsPercentStep = Int(100 / FileLen(convert_filename(FileNums)) / 350 + 0.5)
       End If
       IncreasePer = 1
       IncreaseLine = 0
       
       FunctionStr = ReadFileHead(CorrectFile)
       If Not CorrectFile Then
          MsgBox "文件格式或内容不对！", 64, "提示"
          GoTo NextFile
       End If
       FileNumber1 = FreeFile
       If TchFlag And (Not ConvertTch) Then
          ConvertTch = True
          Gsm_File2 = Gsm_Path + "\sts\tch_sts.dbf"
          Gsm_FileName = Gsm_Path + "\e_tch.dbf"
          FileCopy Gsm_FileName, Gsm_File2
          Open Gsm_File2 For Binary As FileNumber1
          Seek #FileNumber1, 610
          Label1.Caption = "正在生成 " & Gsm_File2
       ElseIf (Not TchFlag) And (Not ConvertCch) Then
          ConvertCch = True
          Gsm_File2 = Gsm_Path + "\sts\cch_sts.dbf"
          Gsm_FileName = Gsm_Path + "\e_cch.dbf"
          FileCopy Gsm_FileName, Gsm_File2
          Open Gsm_File2 For Binary As FileNumber1
          Seek #FileNumber1, 482
          Label1.Caption = "正在生成 " & Gsm_File2
       Else
          GoTo NextFile
       End If
       FileNumber2 = FreeFile
       Gsm_File2 = Gsm_Path + "\map\cell.dbf"
       Open Gsm_File2 For Binary As FileNumber2
       
       For i = 1 To ObjectTypeNum
           LastString = FunctionStr
           FunctionStr = ObjTypeRec(LastString)
           LastString = FunctionStr
           FunctionStr = ObjectRec(LastString)
           If Convert_Stop Then
              Exit For
           End If
       Next
       Seek #FileNumber1, 5
       Put #FileNumber1, , StsRecordno
       Close #FileNumber1
'********************************************************
       If TchFlag Then
          mapinfo.Do "Register Table " + Chr(34) + Gsm_Path + "\sts\tch_sts.dbf" + Chr(34) + " Type " + Chr(34) + "DBF" + Chr(34) + " Into " + Chr(34) + Gsm_Path + "\sts\tch_sts.tab" + Chr(34)
          mapinfo.Do "Open Table " + Chr(34) + Gsm_Path + "\sts\tch_sts.tab" + Chr(34)
          mapinfo.Do "fetch first from tch_sts"
          StsRows = Val(mapinfo.eval("tableinfo(tch_sts,8)"))
          mapinfo.Do "Create Map For tch_sts CoordSys Earth Projection 1, 0 "
          mapinfo.Do "Set Style Pen MakePen(1,60,0)"
          mapinfo.Do "set style brush  makebrush(2,0,0) "
          mapinfo.Do "Set Style Symbol MakeSymbol(33,0,2)"
          For i = 1 To StsRows
              mapinfo.Do "create point into variable sts_mypoint (tch_sts.lon ,tch_sts.lat) symbol(34,7585792,2)"
              mapinfo.Do "update tch_sts set Obj=sts_mypoint where rowid=" & i
              mapinfo.Do "fetch next from tch_sts"
          Next
          mapinfo.Do "commit table tch_sts"
          mapinfo.Do "close table tch_sts"
       Else
          mapinfo.Do "Register Table " + Chr(34) + Gsm_Path + "\sts\cch_sts.dbf" + Chr(34) + " Type " + Chr(34) + "DBF" + Chr(34) + " Into " + Chr(34) + Gsm_Path + "\sts\cch_sts.tab" + Chr(34)
          mapinfo.Do "Open Table " + Chr(34) + Gsm_Path + "\sts\cch_sts.tab" + Chr(34)
          mapinfo.Do "fetch first from cch_sts"
          StsRows = Val(mapinfo.eval("tableinfo(cch_sts,8)"))
          mapinfo.Do "Create Map For cch_sts CoordSys Earth Projection 1, 0 "
          mapinfo.Do "Set Style Pen MakePen(1,60,0)"
          mapinfo.Do "set style brush  makebrush(2,0,0) "
          mapinfo.Do "Set Style Symbol MakeSymbol(33,0,2)"
          For i = 1 To StsRows
              mapinfo.Do "create point into variable sts_mypoint (cch_sts.lon ,cch_sts.lat) symbol(34,7585792,2)"
              mapinfo.Do "update cch_sts set Obj=sts_mypoint where rowid=" & i
              mapinfo.Do "fetch next from cch_sts"
          Next
          mapinfo.Do "commit table cch_sts"
          mapinfo.Do "close table cch_sts"
       End If
'********************************************************

NextFile:
       Close #FileNumber3
       Close #FileNumber2
       FileNums = FileNums + 1
    Loop

End Sub

Function ReadFileHead(ResultVar As Boolean) As String
    Dim LineString As String
    Dim i As Integer
    Dim TempStr As String
    
    On Error Resume Next
    ResultVar = False
    Do While Not EOF(FileNumber3)
       Line Input #FileNumber3, LineString
       If InStr(UCase(LineString), "FILE DUMP") > 0 Then
          Do While Not EOF(FileNumber3)
             Line Input #FileNumber3, LineString
             If InStr(UCase$(LineString), "CELLTCH") > 0 Or InStr(UCase$(LineString), "CELLCCH") > 0 Then
                If InStr(UCase$(LineString), "CELLTCH") > 0 Then
                   TchFlag = True
                Else
                   TchFlag = False
                End If
                ResultVar = True
                Exit Do
             ElseIf Asc(Left(Trim(LineString), 1)) < 58 And Asc(Left(Trim(LineString), 1)) > 47 Then
                Exit Function
             End If
          Loop
          Exit Do  'Is it necessary?
       End If
    Loop
    If Not ResultVar Then
       Exit Function
    End If
    For i = 1 To 3
        LineString = StsReadLine
    Next
    TempStr = ""
    LineString = Trim(Right(LineString, Len(LineString) - 31))
    For i = 1 To 4
        If Left(LineString, 2) <> "00" Then
           TempStr = TempStr & Chr(Val("&H" & Left(LineString, 2)))
        End If
        LineString = Trim(Right(LineString, Len(LineString) - 2))
    Next
    ObjectTypeNum = Val(TempStr)
    ReadFileHead = Trim(Right(LineString, 13))
End Function

Function ObjTypeRec(RecString As String) As String
    Dim LineString As String, ReadTemp As String
    Dim TempString1 As String, TempString2 As String
    Dim i As Integer, j As Integer
        
    On Error Resume Next
    LineString = RecString & " " & Trim(StsReadLine)
    TempString1 = Mid(LineString, 42, 13)
    TempString1 = Left(TempString1, InStr(TempString1, " ") - 1) & Right(TempString1, Len(TempString1) - InStr(TempString1, " "))
    TempString2 = ""
    For i = 1 To 6
        If Mid(TempString1, i * 2 - 1, 2) <> "00" Then
           TempString2 = TempString2 & Chr(Val("&H" & Mid(TempString1, i * 2 - 1, 2)))
        End If
    Next
    ObjectRecNum = Val(Trim(TempString2))
    TempString1 = Mid(LineString, 55, 4)
    CounterNum = Chr(Val("&H" & Left(TempString1, 2))) & Chr(Val("&H" & Right(TempString1, 2)))
    ReDim CounterType(1 To CounterNum)
    LineString = Right(LineString, Len(LineString) - 59)
    For i = 1 To CounterNum
        If Len(LineString) < 64 Then
           ReadTemp = StsReadLine
           If ReadTemp = "" Then
              Convert_Stop = True
              Exit Function
           Else
              LineString = Trim(LineString & " " & ReadTemp)
           End If
        End If
        TempString1 = ""
        For j = 1 To 16
            TempString1 = TempString1 & Chr(Val("&H" & Left(LineString, 2)))
            LineString = Trim(Right(LineString, Len(LineString) - 2))
        Next
        TempString1 = Trim(TempString1)
        Select Case TempString1
           Case "TAVAACC", "CAVAACC"
               CounterType(i) = 1
           Case "TAVASCAN", "CAVASCAN"
               CounterType(i) = 2
           Case "TCALLS", "CCALLS"
               CounterType(i) = 3
           Case "TCONGS", "CCONGS"
               CounterType(i) = 4
           Case "TCONGSSUB", "CCONGSSUB"
               CounterType(i) = 5
           Case "TDISQA", "CDISQA"
               CounterType(i) = 6
           Case "TDISSS", "CDISSS"
               CounterType(i) = 7
           Case "TMSESTB", "CMSESTB"
               CounterType(i) = 8
           Case "TNDROP", "CNDROP"
               CounterType(i) = 9
           Case "TNSCAN", "CNSCAN"
               CounterType(i) = 10
           Case "TTRALACC", "CTRALACC"
               CounterType(i) = 11
           Case Else
               CounterType(i) = 0
        End Select
    Next
    ObjTypeRec = LineString
End Function

Function ObjectRec(RecString As String) As String
    Dim i As Integer, k As Integer
    Dim TempString As String, ReadTemp As String
    Dim TchData As e_tch
    Dim CchData As e_cch
    Dim j As Long
    Dim CellData As NewCellStru
    
    On Error Resume Next
    LineString = Trim(RecString & " " & StsReadLine)
    For i = 1 To ObjectRecNum
        IncreaseLine = IncreaseLine + 1
        If IncreaseLine = IncreasePer * LinesPer And ProgressBar1.Value + StsPercentStep < 100 Then
           ProgressBar1.Value = ProgressBar1.Value + StsPercentStep
           IncreasePer = IncreasePer + 1
        End If
        If Len(LineString) < 64 Then
           ReadTemp = StsReadLine
           If ReadTemp = "" Then
              Convert_Stop = True
              Exit Function
           Else
              LineString = Trim(LineString & " " & ReadTemp)
           End If
        End If
        TempString = LineString
        LineString = ValidProcess(TempString)
        TempString = ""
        For j = 1 To 16
            If Left(LineString, 2) <> "00" Then
               TempString = TempString & Chr(Val("&H" & Left(LineString, 2)))
            End If
            LineString = Trim(Right(LineString, Len(LineString) - 2))
        Next
        TchData.ID = Trim(TempString)
        CchData.ID = Trim(TempString)
        For j = 1 To CounterNum
            TempString = ""
            If Len(LineString) < 64 Then
               LineString = Trim(LineString & " " & StsReadLine)
            End If
            For k = 1 To 10
                If Left(LineString, 2) <> "00" Then
                   TempString = TempString & Chr(Val("&H" & Left(LineString, 2)))
                End If
                LineString = Trim(Right(LineString, Len(LineString) - 2))
            Next
            If CounterType(j) > 0 Then
               If Trim(TempString) = "" Then
                  CounterData(CounterType(j)) = "0"
               Else
                  CounterData(CounterType(j)) = Trim(TempString)
               End If
            End If
        Next
        If TchFlag Then
           If Val((CounterData(10))) = 0 Or Val((CounterData(1))) = 0 Then    '每线话务量
              TchData.tch_col(2) = "0.00"
           Else
              TchData.tch_col(2) = Format((Val((CounterData(11)) * Val(CounterData(2))) / Val((CounterData(10)) * Val(CounterData(1)))), "0.00")
           End If
           If Val((CounterData(2))) = 0 Then     '可用信道
              TchData.tch_col(7) = "0"
           Else
              TchData.tch_col(7) = Val(CounterData(1)) / Val(CounterData(2))
           End If
           If Val((CounterData(3))) = 0 Then
              TchData.tch_col(3) = "0.00"        '拥塞率
              TchData.tch_col(1) = "0.00"        '接通率
           Else
              TchData.tch_col(3) = Format(((Val(CounterData(4)) + Val(CounterData(5))) / Val(CounterData(3))) * 100, "0.00")
              TchData.tch_col(1) = Format((Val(CounterData(8)) / Val(CounterData(3))) * 100, "0.00")
           End If
           If Val((CounterData(8))) = 0 Then     '掉话率
              TchData.tch_col(5) = "0.00"
           Else
              TchData.tch_col(5) = Format((Val(CounterData(9)) / Val(CounterData(8))) * 100, "0.00")
           End If
           If Val((CounterData(9))) = 0 Then
              TchData.tch_col(9) = "0.00"        '质差断线
              TchData.tch_col(10) = "0.00"       '弱信号断线
           Else
              TchData.tch_col(9) = Format((Val(CounterData(6)) / Val(CounterData(9))) * 100, "0.00")
              TchData.tch_col(10) = Format((Val(CounterData(7)) / Val(CounterData(9))) * 100, "0.00")
           End If
           j = 0
           TchData.TCH = ""
           TchData.lon = ""
           TchData.lat = ""
           TchData.bearing = ""
           Do While Not EOF(FileNumber2)
              Seek #FileNumber2, 1026 + 309 * j
              Get #FileNumber2, , CellData
              If UCase(Trim(CellData.bs_no)) = UCase(Trim(TchData.ID)) Then
                 TchData.TCH = CellData.Name
                 TchData.bearing = CellData.bearing
                 TchData.lon = Val(CellData.lon) + 0.0015 * Sin(Val(CellData.bearing) * 0.01745329252)
                 TchData.lat = Val(CellData.lat) + 0.0015 * Cos(Val(CellData.bearing) * 0.01745329252)
                 Exit Do
              End If
              j = j + 1
           Loop
           Seek #FileNumber2, 1
           Put #FileNumber1, , TchData
        Else
           If Val((CounterData(10))) = 0 Or Val((CounterData(1))) = 0 Then
              CchData.cch_col(2) = "0.00"
           Else
              CchData.cch_col(2) = Format((Val((CounterData(11)) * Val(CounterData(2))) / Val((CounterData(10)) * Val(CounterData(1)))) * 100, "0.00")
           End If
           If Val((CounterData(3))) = 0 Then
              CchData.cch_col(3) = "0.00"
              'CchData.cch_col(1) = "0.00"
           Else
              CchData.cch_col(3) = Format(((Val(CounterData(4)) + Val(CounterData(5))) / Val(CounterData(3))) * 100, "0.00")
              'CchData.cch_col(1) = Format((Val(CounterData(8)) / Val(CounterData(3))) * 100, "0.00")
           End If
           If Val((CounterData(9))) = 0 Then
              CchData.cch_col(5) = "0.00"        '质差断线
              CchData.cch_col(6) = "0.00"       '弱信号断线
           Else
              CchData.cch_col(5) = Format((Val(CounterData(6)) / Val(CounterData(9))) * 100, "0.00")
              CchData.cch_col(6) = Format((Val(CounterData(7)) / Val(CounterData(9))) * 100, "0.00")
           End If
           
           j = 0
           CchData.CCH = ""
           CchData.lon = ""
           CchData.lat = ""
           CchData.bearing = ""
           Do While Not EOF(FileNumber2)
              Seek #FileNumber2, 1026 + 309 * j
              Get #FileNumber2, , CellData
              If UCase(Trim(CellData.bs_no)) = UCase(Trim(CchData.ID)) Then
                 CchData.CCH = CellData.Name
                 CchData.bearing = CellData.bearing
                 CchData.lon = Val(CellData.lon) + 0.0015 * Sin(Val(CellData.bearing) * 0.01745329252)
                 CchData.lat = Val(CellData.lat) + 0.0015 * Cos(Val(CellData.bearing) * 0.01745329252)
                 Exit Do
              End If
              j = j + 1
           Loop
           Seek #FileNumber2, 1
           Put #FileNumber1, , CchData
        End If
        StsRecordno = StsRecordno + 1
        DoEvents
        If Convert_Stop Then
           Exit Function
        End If
    Next
End Function

Function StsReadLine() As String
    Dim ReturnStr As String
    
    On Error Resume Next
    Do While Not EOF(FileNumber3)
       Line Input #FileNumber3, ReturnStr
       ReturnStr = Trim(ReturnStr)
       If Asc(Left(ReturnStr, 1)) < 58 And Asc(Left(ReturnStr, 1)) > 47 Then
          StsReadLine = Trim(ReturnStr)
          Exit Function
       End If
    Loop
    StsReadLine = ""
End Function

Function ValidProcess(ProcessString As String) As String
    Dim i As Integer
    
    On Error Resume Next
CheckAgain:
    Do
       If Left(ProcessString, 2) = "00" Then
          If Len(ProcessString) > 2 Then
             ProcessString = Trim(Right(ProcessString, Len(ProcessString) - 2))
          Else
             ProcessString = ""
             Exit Do
          End If
       Else
          ValidProcess = ProcessString
          Exit Function
       End If
    Loop
    If ProcessString = "" Then
       ProcessString = StsReadLine
       GoTo CheckAgain
    Else
       If Len(ProcessString) < 60 Then
          ProcessString = ProcessString & " " & StsReadLine
       End If
    End If
    ValidProcess = ProcessString
End Function

Sub CellMotorolaXLS()
    Dim XLSRows As Integer, i As Integer, j As Integer
    Dim CellPos As Long
    Dim CellData As NewCellStru
    Dim XLS_Ci As String, XLS_Lac As String
    Dim FileNumber As Integer
    Dim PerCount As Integer, PerRows As Integer
              
    On Error Resume Next
    FileNumber = FreeFile
    Gsm_FileName = Gsm_Path + "\map\cell.dbf"
    Open Gsm_FileName For Binary As #FileNumber
    XLSRows = mapinfo.eval("tableinfo(celltemp,8)")
    PerRows = Int((XLSRows / 100) + 0.5)
    PerCount = 0
    ProgressBar1.Value = 0
    Label1.Caption = "正在转换 " + convert_filename(1)
    mapinfo.Do "fetch first from celltemp"
    For i = 1 To 3
        mapinfo.Do "fetch next from celltemp"
    Next
    load_new = 0
    For i = 1 To XLSRows - 3
        If PerCount = PerRows And ProgressBar1.Value < 90 Then
           ProgressBar1.Value = ProgressBar1.Value + 1
           PerCount = 0
        End If
        PerCount = PerCount + 1
        If Trim(mapinfo.eval("celltemp.col1")) = "" Then
           Exit For
        End If
        XLS_Ci = mapinfo.eval("celltemp.col6")
        CellPos = 0
        Do While Not EOF(FileNumber)
           Seek #FileNumber, 1026 + CellPos * 309
           Get #FileNumber, , CellData
           If Trim(XLS_Ci) = Trim(CellData.ci) Then
              XLS_Lac = Trim(mapinfo.eval("celltemp.col5"))
              If UCase(Right(XLS_Lac, 1)) = "H" Or UCase(Right(XLS_Lac, 1)) = "O" Then
                 XLS_Lac = Format(Val("&" & Right(XLS_Lac, 1) & Format(Val(XLS_Lac))))
              End If
              CellData.Lac = XLS_Lac
              CellData.BSIC = Left(Trim(CellData.BSIC), 1) & Trim(mapinfo.eval("celltemp.col7"))
              CellData.ARFCN = Trim(mapinfo.eval("celltemp.col8"))
              CellData.time = DATE
              'For j = 1 To 4
              '    CellData.bcch(j) = Trim(mapinfo.eval("celltemp.col" & Format(j + 8)))
              'Next
              'CellData.bcch(5) = ""
              'CellData.bcch(6) = ""
              CellData.Nonbcch = ""
              For j = 1 To 4
                  CellData.Nonbcch = Trim(CellData.Nonbcch) & Trim(mapinfo.eval("celltemp.col" & Format(j + 8))) & ","
              Next
              If Len(Trim(CellData.Nonbcch)) > 0 Then
                 CellData.Nonbcch = Left(CellData.Nonbcch, Len(Trim(CellData.Nonbcch)) - 1)
              End If
              Seek #FileNumber, 1026 + CellPos * 309
              Put #FileNumber, , CellData
              load_new = load_new + 1
              Exit Do
           End If
           CellPos = 1 + CellPos
        Loop
        Close #FileNumber
        Open Gsm_FileName For Binary As #FileNumber
        DoEvents
        If Convert_Stop = True Then
           Exit For
        End If
        mapinfo.Do "fetch next from celltemp"
    Next
    Close #FileNumber
    mapinfo.Do "close table celltemp"
    Gsm_FileName = Gsm_Path + "\celltemp.*"
    Kill Gsm_FileName
    
End Sub

Sub CellExcel()
    
    On Error Resume Next
    
End Sub

Sub NcellMotorolaXLS()
    Dim XLSRows As Integer, i As Integer, j As Integer
    Dim Ncelldata As aell
    Dim FileNumber As Integer
    Dim PerCount As Integer, PerRows As Integer
    Dim Ncellrecord As Integer
    
    On Error Resume Next
    mapinfo.Do "Register Table " + Chr(34) + convert_filename(1) + Chr(34) + " TYPE XLS Into " + Chr(34) + Gsm_Path + "\NcellTemp.tab" + Chr(34)
    mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\NcellTemp.tab" + Chr(34)
    If Err Then
       MsgBox "无法打开文件 " & convert_filename(1) & "或文件格式错误", 64, "提示"
       Screen.MousePointer = 0
       Exit Sub
    End If
    Gsm_File2 = Gsm_Path + "\map\cell"
    mapinfo.Do "open table " + Chr(34) + Gsm_File2 + Chr(34)
    mapinfo.Do "fetch first from cell"
    FileNumber = FreeFile
    Gsm_FileName = Gsm_Path + "\map\ncell.dbf"
    Gsm_File2 = Gsm_Path + "\ncell.dbf"
    FileCopy Gsm_File2, Gsm_FileName
    Open Gsm_FileName For Binary As #FileNumber
    Seek #FileNumber, 2210
    XLSRows = mapinfo.eval("tableinfo(ncelltemp,8)")
    PerRows = Int((XLSRows / 100) + 0.5)
    PerCount = 0
    ProgressBar1.Value = 0
    Ncellrecord = 0
    Label1.Caption = "正在转换 " + convert_filename(1)
    mapinfo.Do "fetch first from ncelltemp"
    For i = 1 To 3
        mapinfo.Do "fetch next from ncelltemp"
    Next
    For i = 1 To XLSRows - 3
        If PerCount = PerRows And ProgressBar1.Value < 90 Then
           ProgressBar1.Value = ProgressBar1.Value + 1
           PerCount = 0
        End If
        PerCount = PerCount + 1
        If Trim(mapinfo.eval("ncelltemp.col1")) = "" Then
           Exit For
        End If
        Ncelldata.ci = mapinfo.eval("ncelltemp.col6")
        Ncelldata.bs_name = FindNArfcn(Ncelldata.ci, Ncelldata.Lac, Ncelldata.bs_no)
        For j = 1 To 16
            Ncelldata.col(j).ci_c = Trim(mapinfo.eval("ncelltemp.col" & Format(j + 12)))
        Next
        Put #FileNumber, , Ncelldata
        Ncellrecord = Ncellrecord + 1
        DoEvents
        If Convert_Stop = True Then
           Exit For
        End If
        mapinfo.Do "fetch next from ncelltemp"
        
    Next
    Seek #FileNumber, 5
    Put #FileNumber, , Ncellrecord
    Close #FileNumber
    mapinfo.Do "close table ncelltemp"
    mapinfo.Do "close table cell"
    Gsm_FileName = Gsm_Path + "\ncelltemp.*"
    Kill Gsm_FileName
    
End Sub

Sub UpdateCell()
    Dim CellData As NewCellStru
    Dim CellRows As Variant
    Dim PerCount As Integer, PerRows As Integer
    Dim i As Integer, j As Long
    Dim NonBcchtemp As String
    Dim Cellrecord As Long
    Dim OldCellData As Oldcellstru
    Dim OldCellData1 As Oldcellstru1
    Dim MyMicrocell As String * 1
    
    On Error Resume Next
    Label1.Caption = "系统正在更新基站库,旧的基站库将改名为 Cell.Old"
    Me.Caption = "更新系统"
    'FileCopy Gsm_Path + "\cellstru.dbf", Gsm_Path + "\map\cell.new"
    If dir(Gsm_Path + "\map\cell.new", 0) <> "" Then
       Kill Gsm_Path + "\map\cell.new"
    End If
    Open Gsm_Path + "\map\cell.new" For Binary As #8
    hDbfFile = 8
    MakeCellFile
    'Open Gsm_Path + "\map\cell.new" For Binary As #8
    Seek #8, 1026
    mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\map\cell" + Chr(34)
    CellRows = mapinfo.eval("tableinfo(cell,8)")
    mapinfo.Do "close table cell"
    PerRows = Int((CellRows / 100) + 0.5)
    PerCount = 0
    ProgressBar1.Value = 0
    CellData.b = " "
    For i = 1 To 16
        CellData.NCELL(i) = space$(10)
    Next
    Open Gsm_Path + "\map\cell.dbf" For Binary As #9
    If Menu_Flag = 9998 Then
       Seek #9, 930
    ElseIf Menu_Flag = 9999 Then
       Seek #9, 962
    Else
       Seek #9, 1026
    End If
    For i = 1 To CellRows
        If PerCount = PerRows And ProgressBar1.Value < 90 Then
           ProgressBar1.Value = ProgressBar1.Value + 1
           PerCount = 0
        End If
        PerCount = PerCount + 1
        If Menu_Flag = 9997 Then
            Get #9, , OldCellData1
            CellData.Name = Trim(OldCellData1.Name)
            If InStr(CellData.Name, Chr(0)) > 0 Then
               CellData.Name = Left(CellData.Name, InStr(CellData.Name, Chr(0)))
            End If
            CellData.bs_no = Trim(OldCellData1.bs_no)
            CellData.ci = Trim(OldCellData1.ci)
            CellData.ARFCN = Trim(OldCellData1.ARFCN)
            CellData.BSIC = Trim(OldCellData1.BSIC)
            CellData.bearing = Trim(OldCellData1.bearing)
            CellData.Lac = Trim(OldCellData1.Lac)
            CellData.downtilt = Trim(OldCellData1.downtilt)
            CellData.max_bts = Trim(OldCellData1.max_bts)
            CellData.max_ms = Trim(OldCellData1.max_ms)
            CellData.time = Trim(OldCellData1.time)
            CellData.lon = Trim(OldCellData1.lon)
            CellData.lat = Trim(OldCellData1.lat)
            CellData.microcell = Trim(OldCellData1.microcell)
            CellData.Nonbcch = Trim(OldCellData1.Nonbcch)
            For j = 1 To 16
                CellData.NCELL(j) = Trim(OldCellData1.NCELL(j))
            Next
        Else
            Get #9, , OldCellData
            CellData.Name = Trim(OldCellData.Name)
            If InStr(CellData.Name, Chr(0)) > 0 Then
               CellData.Name = Left(CellData.Name, InStr(CellData.Name, Chr(0)))
            End If
            CellData.bs_no = Trim(OldCellData.bs_no)
            CellData.ci = Trim(OldCellData.ci)
            CellData.ARFCN = Trim(OldCellData.ARFCN)
            CellData.BSIC = Trim(OldCellData.BSIC)
            CellData.bearing = Trim(OldCellData.bearing)
            CellData.Lac = Trim(OldCellData.Lac)
            CellData.downtilt = Trim(OldCellData.downtilt)
            CellData.max_bts = Trim(OldCellData.max_bts)
            CellData.max_ms = Trim(OldCellData.max_ms)
            CellData.time = Trim(OldCellData.time)
            CellData.lon = Trim(OldCellData.lon)
            CellData.lat = Trim(OldCellData.lat)
            If Menu_Flag = 9999 Then
               Get #9, , MyMicrocell
               CellData.microcell = MyMicrocell
            Else
               CellData.microcell = "0"
            End If
            NonBcchtemp = ""
            For j = 1 To 6
                If Val(OldCellData.bcch(j)) <> 0 Then
                   NonBcchtemp = NonBcchtemp + Trim(OldCellData.bcch(j)) + ","
                End If
            Next
            If Len(NonBcchtemp) > 0 Then
               NonBcchtemp = Left(NonBcchtemp, Len(NonBcchtemp) - 1)
            End If
            CellData.Nonbcch = NonBcchtemp
        End If
        Put #8, , CellData
    Next
    Cellrecord = CellRows
    Seek #8, 5
    Put #8, , Cellrecord
    Close #8
    Close #9
    FileCopy Gsm_Path + "\map\cell.dbf", Gsm_Path + "\map\cell.old"
    FileCopy Gsm_Path + "\map\cell.new", Gsm_Path + "\map\cell.dbf"
    mapinfo.Do "Register Table " + Chr(34) + Gsm_Path + "\map\cell.dbf" + Chr(34) + "Type ""DBF"" Into " + Chr(34) + Gsm_Path + "\map\cell.tab" + Chr(34)
    mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\map\cell.tab" + Chr(34)
         
          For i = 1 To Cellrecord
          mapinfo.Do " x1 = cell.Lon + 0.002 * Sin(Cell.bearing * 0.01745329252)" '  DEG_2_RAD)"
          mapinfo.Do " y1 = cell.Lat + 0.002 * Cos(Cell.bearing * 0.01745329252)"  ' DEG_2_RAD)"

          j = Val(mapinfo.eval("cell.arfcn"))
          If j <> 0 Then
          Select Case j
                 Case 1
                       j = 16711680
                 Case 2
                       j = 65280
                 Case 3
                       j = 255
                 Case 4
                       j = 16711935
                 Case 5
                       j = 16776960
                 Case 6
                       j = 65535
                 Case 7
                       j = 8388608
                 Case 8
                       j = 32768
                 Case 9
                       j = 128
                 Case 10
                       j = 8388736
                 Case 11
                       j = 8421376
                 Case 12
                       j = 32896
                 Case 13
                       j = 16744576
                 Case 14
                       j = 8454016
                 Case 15
                       j = 8421631
                 Case 16
                       j = 16744703
                 Case 17
                       j = 16777088
                 Case 18
                       j = 8454143
                 Case 19
                       j = 8405056
                 Case 20
                       j = 4227136
                 Case 21
                       j = 4210816
                 Case 22
                       j = 8405120
                 Case 23
                       j = 8421440
                 Case 24
                       j = 4227200
                 Case 25
                       j = 16761024
                 Case 26
                       j = 12648384
                 Case 27
                       j = 12632319
                 Case 28
                       j = 16761087
                 Case 29
                       j = 16777152
                 Case 30
                       j = 12648447
                 Case 31
                       j = 8413280
                 Case 32
                       j = 6324320
                 Case 33
                       j = 6316160
                 Case 34
                       j = 8413312
                 Case 35
                       j = 8421472
                 Case 36
                       j = 6324352

                 Case 37
                       j = 16711680
                 Case 38
                       j = 65280
                 Case 39
                       j = 255
                 Case 40
                       j = 16711935
                 Case 41
                       j = 16776960
                 Case 42
                       j = 65535
                 Case 43
                       j = 8388608
                 Case 44
                       j = 32768
                 Case 45
                       j = 128
                 Case 46
                       j = 8388736
                 Case 47
                       j = 8421376
                 Case 48
                       j = 32896
                 Case 49
                       j = 16744576
                 Case 50
                       j = 8454016
                 Case 51
                       j = 8421631
                 Case 52
                       j = 16744703
                 Case 53
                       j = 16777088
                 Case 54
                       j = 8454143
                 Case 55
                       j = 8405056
                 Case 56
                       j = 4227136
                 Case 57
                       j = 4210816
                 Case 58
                       j = 8405120
                 Case 59
                       j = 8421440
                 Case 60
                       j = 4227200
                 Case 61
                       j = 16761024
                 Case 62
                       j = 12648384
                 Case 63
                       j = 12632319
                 Case 64
                       j = 16761087
                 Case 65
                       j = 16777152
                 Case 66
                       j = 12648447
                 Case 67
                       j = 8413280
                 Case 68
                       j = 6324320
                 Case 69
                       j = 6316160
                 Case 70
                       j = 8413312
                 Case 71
                       j = 8421472
                 Case 72
                       j = 6324352
                 Case 73
                       j = 16711680
                 Case 74
                       j = 65280
                 Case 75
                       j = 255
                 Case 76
                       j = 16711935
                 Case 77
                       j = 16776960
                 Case 78
                       j = 65535
                 Case 79
                       j = 8388608
                 Case 80
                       j = 32768
                 Case 81
                       j = 128
                 Case 82
                       j = 8388736
                 Case 83
                       j = 8421376
                 Case 84
                       j = 32896
                 Case 85
                       j = 16744576
                 Case 86
                       j = 8454016
                 Case 87
                       j = 8421631
                 Case 88
                       j = 16744703
                 Case 89
                       j = 16777088
                 Case 90
                       j = 8454143
                 Case 91
                       j = 8405056
                 Case 92
                       j = 4227136
                 Case 93
                       j = 4210816
                 Case 94
                       j = 8405120
                 Case 95
                       j = 8421440
                 Case 96
                       j = 4227200
                 Case 97
                       j = 16761024
                 Case 98
                       j = 12648384
                 Case 99
                       j = 12632319
                 Case 100
                       j = 16761087
                 Case 101
                       j = 16777152
                 Case 102
                       j = 12648447
                 Case 103
                       j = 8413280
                 Case 104
                       j = 6324320
                 Case 105
                       j = 6316160
                 Case 106
                       j = 8413312
                 Case 107
                       j = 8421472
                 Case 108
                       j = 6324352
                 Case 109
                       j = 16711680
                 Case 110
                       j = 65280
                 Case 111
                       j = 255
                 Case 112
                       j = 16711935
                 Case 113
                       j = 16776960
                 Case 114
                       j = 65535
                 Case 115
                       j = 8388608
                 Case 116
                       j = 32768
                 Case 117
                       j = 128
                 Case 118
                       j = 8388736
                 Case 119
                       j = 8421376
                 Case 120
                       j = 32896
                 Case 121
                       j = 16744576
                 Case 122
                       j = 8454016
                 Case 123
                       j = 8421631
          End Select
              If Val(mapinfo.eval("cell.microcell")) = 0 Then
                 mapinfo.Do "Set Style Pen MakePen(1,60," & j & ")"
                 mapinfo.Do "update cell  set Obj= CreateLine(x1,y1,cell.lon, cell.Lat)  where rowid=" & i
              Else
                 mapinfo.Do "Set Style Pen MakePen(1,58," & j & ")"
                 mapinfo.Do "update cell  set Obj= CreateLine(x1,y1,cell.lon, cell.Lat)  where rowid=" & i
              
              End If
              
          End If
          mapinfo.Do "fetch next from cell"
      Next
      mapinfo.Do "commit table cell"
      mapinfo.Do "close table cell"
End Sub

Sub Cell_Nortel()
    Dim Old_Cell As NewCellStru
    Dim recordno As Long
    Dim Line_Char As String
    Dim End_Char As String * 1
    Dim i As Integer, File_num As Integer, j As Integer
    Dim FileLenth As Long, nline As Integer, bline As Integer
    Dim PercentStep As Integer, bs As Integer, scnline As Integer
    Dim my_Pos As Long
    
    On Error Resume Next
    File_num = 0
    Do While Trim(convert_filename(File_num + 1)) <> ""
       File_num = File_num + 1
    Loop
    End_Char = Chr$(26)
    load_new = 0
    Gsm_FileName = Gsm_Path + "\map\cell.dbf"
    Open Gsm_FileName For Binary As #2
    For i = 1 To File_num
        FileLenth = FileLen(convert_filename(i))
        nline = FileLenth / 300
        bline = Fix(nline / 100)
        PercentStep = 1
        If bline = 0 Then
           bline = 1
           PercentStep = 100 / nline
        End If
        bs = 1
        scnline = 0
        ProgressBar1.Value = 0
        Label1.Caption = "正在转换 " + convert_filename(i)
        Label1.Refresh
        end_frag = False
        Open convert_filename(i) For Binary As #1
        Have_Read = False
        Do While 1
           New_Cell.ci = space(5)
           New_Cell.BSIC = space(3)
           'New_Cell.ARFCN = space(3)
           New_Cell.bs_no = space(10)
           'New_Cell.max_bts = space(2)
           'New_Cell.max_ms = space(2)
           'New_Cell.Nonbcch = space(32)
           New_Cell.Lac = ""
           scnline = scnline + 1
           If scnline = bs * bline And ProgressBar1.Value < 99 Then
              ProgressBar1.Value = ProgressBar1.Value + PercentStep
              bs = bs + 1
           End If
           Read_NortelCell
           If end_frag = True Then
              Exit Do
           End If
           my_Pos = 0
           Do While Not EOF(2)
              Seek #2, 1026 + my_Pos * 309
              Get #2, , Old_Cell
              If Trim(UCase(Old_Cell.bs_no)) = Trim(New_Cell.bs_no) Then
                 If end_frag = True Then
                    Exit Do
                 End If
                 Old_Cell.Lac = New_Cell.Lac
                 Old_Cell.ci = New_Cell.ci
                 Old_Cell.BSIC = New_Cell.BSIC
                 Old_Cell.time = DATE
                 'Old_Cell.Nonbcch = Italtel_Nonbcch
                 Seek #2, 1026 + my_Pos * 309
                 Put #2, , Old_Cell
                 load_new = load_new + 1
                 Exit Do
              End If
              my_Pos = my_Pos + 1
           Loop
           If end_frag = True Then
              Exit Do
           End If
           DoEvents
           If Convert_Stop = True Then
              Close
              Exit Sub
           End If
           Close #2
           Open Gsm_FileName For Binary As #2
        Loop
        Close #1
        If ProgressBar1.Value < 100 Then
           ProgressBar1.Value = 100
        End If
    Next
    Close
End Sub

Sub Read_NortelCell()
    Dim finds As Integer
    Dim FindChar As String * 1
    Dim Read_Line As String
    Dim InString As String * 200
    
    On Error Resume Next
    'If Have_Read = True Then
    '   Have_Read = False
    '   Read_Line = Public_Line
    '   GoTo Find_again
    'End If
Read_again:
    If InStr(Nortel_Line, Chr(10)) = 0 Then
        Do While Not EOF(1)
           Get #1, , InString
           If Trim(InString) <> "" Then
              Exit Do
           End If
        Loop
        Nortel_Line = Nortel_Line & InString
    End If
    If EOF(1) Then
       end_frag = True
       Exit Sub
    End If
Find_again:
    If InStr(Nortel_Line, Chr(10)) = 0 Then
       GoTo Read_again
    End If
    Read_Line = Left(Nortel_Line, InStr(Nortel_Line, Chr(10)) - 1)
    Nortel_Line = Right(Nortel_Line, Len(Nortel_Line) - InStr(Nortel_Line, Chr(10)))
    Read_Line = Trim(Read_Line)
    FindChar = ","
    finds = InStr(Read_Line, FindChar)
    If finds = 0 Then
       FindChar = Chr(9)
       finds = InStr(Read_Line, FindChar)
       If finds = 0 Then
          FindChar = Chr(20)
          finds = InStr(Read_Line, FindChar)
          If finds = 0 Then
             If EOF(1) Then
                end_frag = True
                Exit Sub
             End If
             GoTo Read_again
          End If
       End If
    End If
    Do While Left(Read_Line, 1) = FindChar
       Read_Line = Trim(Right(Read_Line, Len(Read_Line) - 1))
    Loop
    finds = InStr(Read_Line, FindChar)
    'If finds > 0 Then
       Read_Line = Right(Read_Line, Len(Read_Line) - finds)
       finds = InStr(Read_Line, FindChar)
       If finds > 0 Then
          New_Cell.bs_no = Trim(Left(Read_Line, finds - 1))
          Read_Line = Right(Read_Line, Len(Read_Line) - finds)
          finds = InStr(Read_Line, FindChar)
          Read_Line = Right(Read_Line, Len(Read_Line) - finds)
          finds = InStr(Read_Line, FindChar)
          New_Cell.ci = Trim(Left(Read_Line, finds - 1))
          Read_Line = Right(Read_Line, Len(Read_Line) - finds)
          finds = InStr(Read_Line, FindChar)
          New_Cell.Lac = Trim(Left(Read_Line, finds - 1))
          Read_Line = Right(Read_Line, Len(Read_Line) - finds)
          finds = InStr(Read_Line, FindChar)
          'New_Cell.Lac = Trim(Left(Read_Line, finds - 1))
          Read_Line = Right(Read_Line, Len(Read_Line) - finds)
          finds = InStr(Read_Line, FindChar)
          Read_Line = Right(Read_Line, Len(Read_Line) - finds)
          finds = InStr(Read_Line, FindChar)
          New_Cell.BSIC = Trim(Left(Read_Line, finds - 1))
          Read_Line = Right(Read_Line, Len(Read_Line) - finds)
          finds = InStr(Read_Line, FindChar)
          New_Cell.BSIC = Trim(Left(Read_Line, finds - 1)) & Trim(New_Cell.BSIC)
       Else
          GoTo Read_again
       End If
    'Else
    '   If EOF(1) Then
    '      end_frag = True
    '      Exit Sub
    '   End If
    'End If
End Sub

Sub NCell_Nortel()
    Dim Old_Cell As NewCellStru
    Dim recordno As Long
    Dim Line_Char As String
    Dim End_Char As String * 1
    Dim i As Integer, File_num As Integer, j As Integer
    Dim FileLenth As Long, nline As Integer, bline As Integer
    Dim PercentStep As Integer, bs As Integer, scnline As Integer
    Dim my_Pos As Long
    
    On Error Resume Next
    File_num = 0
    Do While Trim(convert_filename(File_num + 1)) <> ""
       File_num = File_num + 1
    Loop
    End_Char = Chr$(26)
    load_new = 0
    Gsm_FileName = Gsm_Path + "\map\cell.dbf"
    Open Gsm_FileName For Binary As #2
    For i = 1 To File_num
        FileLenth = FileLen(convert_filename(i))
        nline = FileLenth / 300
        bline = Fix(nline / 100)
        PercentStep = 1
        If bline = 0 Then
           bline = 1
           PercentStep = 100 / nline
        End If
        bs = 1
        scnline = 0
        ProgressBar1.Value = 0
        Label1.Caption = "正在转换 " + convert_filename(i)
        Label1.Refresh
        end_frag = False
        Open convert_filename(i) For Binary As #1
        Have_Read = False
        Do While 1
           For j = 1 To 16
               New_Ncell.NCELL(j) = ""
           Next
           New_Ncell.bs_no = space(10)
           
           scnline = scnline + 1
           If scnline = bs * bline And ProgressBar1.Value < 99 Then
              ProgressBar1.Value = ProgressBar1.Value + PercentStep
              bs = bs + 1
           End If
           Read_NortelNCell
           If end_frag = True Then
              Exit Do
           End If
           my_Pos = 0
           Do While Not EOF(2)
              Seek #2, 1026 + my_Pos * 309
              Get #2, , Old_Cell
              If Trim(UCase(Old_Cell.bs_no)) = Trim(New_Ncell.bs_no) Then
                 If end_frag = True Then
                    Exit Do
                 End If
                 For j = 1 To 16
                     Old_Cell.NCELL(j) = New_Ncell.NCELL(j)
                 Next
                 Old_Cell.time = DATE
                 Seek #2, 1026 + my_Pos * 309
                 Put #2, , Old_Cell
                 load_new = load_new + 1
                 Exit Do
              End If
              my_Pos = my_Pos + 1
           Loop
           If end_frag = True Then
              Exit Do
           End If
           DoEvents
           If Convert_Stop = True Then
              Close
              Exit Sub
           End If
           Close #2
           Open Gsm_FileName For Binary As #2
        Loop
        Close #1
        If ProgressBar1.Value < 100 Then
           ProgressBar1.Value = 100
        End If
    Next
    Close
End Sub

Sub Read_NortelNCell()
    Dim finds As Integer
    Dim FindChar As String * 1
    Dim Read_Line As String
    Dim InString As String * 200
    
    On Error Resume Next
Read_again:
    If InStr(Nortel_Line, Chr(10)) = 0 Then
        Do While Not EOF(1)
           Get #1, , InString
           If Trim(InString) <> "" Then
              Exit Do
           End If
        Loop
        Nortel_Line = Nortel_Line & InString
    End If
    If EOF(1) Then
       end_frag = True
       Exit Sub
    End If
Find_again:
    If InStr(Nortel_Line, Chr(10)) = 0 Then
       GoTo Read_again
    End If
    Read_Line = Left(Nortel_Line, InStr(Nortel_Line, Chr(10)) - 1)
    Nortel_Line = Right(Nortel_Line, Len(Nortel_Line) - InStr(Nortel_Line, Chr(10)))
    Read_Line = Trim(Read_Line)
    FindChar = ","
    finds = InStr(Read_Line, FindChar)
    If finds = 0 Then
       FindChar = Chr(9)
       finds = InStr(Read_Line, FindChar)
       If finds = 0 Then
          FindChar = Chr(20)
          finds = InStr(Read_Line, FindChar)
          If finds = 0 Then
             If EOF(1) Then
                end_frag = True
                Exit Sub
             End If
             GoTo Read_again
          End If
       End If
    End If
    Do While Left(Read_Line, 1) = FindChar
       Read_Line = Trim(Right(Read_Line, Len(Read_Line) - 1))
    Loop
    finds = InStr(Read_Line, FindChar)
    'If finds > 0 Then
       Read_Line = Right(Read_Line, Len(Read_Line) - finds)
       finds = InStr(Read_Line, FindChar)
       If finds > 0 Then
          New_Ncell.bs_no = Trim(Left(Read_Line, finds - 1))
          Read_Line = Right(Read_Line, Len(Read_Line) - finds)
          finds = InStr(Read_Line, FindChar)
          For j = 1 To 16
              Read_Line = Right(Read_Line, Len(Read_Line) - finds)
              finds = InStr(Read_Line, FindChar)
              If finds = 0 Then
                 New_Ncell.NCELL(j) = Read_Line
                 Exit For
              Else
                 New_Ncell.NCELL(j) = Trim(Left(Read_Line, finds - 1))
              End If
          Next
       Else
          GoTo Read_again
       End If
    'Else
    '   If EOF(1) Then
    '      end_frag = True
    '      Exit Sub
    '   End If
    'End If
End Sub

