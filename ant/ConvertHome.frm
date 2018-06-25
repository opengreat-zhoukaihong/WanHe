VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form Data_Convert 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数据转换"
   ClientHeight    =   1425
   ClientLeft      =   825
   ClientTop       =   6960
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
      TabIndex        =   2
      Top             =   1035
      Width           =   1080
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   315
      TabIndex        =   1
      Top             =   585
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   450
      _Version        =   327680
      Appearance      =   0
      MouseIcon       =   "Convert.frx":030A
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
Dim bc(1 To 6) As String * 3
Dim mbcch(1 To 16) As String * 3
Dim end_frag As Boolean
Dim mb As Integer
Dim New_Cell As cell_stru
Dim New_Ncell As aell
Dim Prev_Bs_no As String * 10
Dim Get_Bs_no As Boolean
Dim Public_Line As String
Dim Have_Read As Boolean
'*****************************
Dim FileNumber3 As Integer
'*****************************

Sub Ncell_Ericsson()
    Dim header As String * 2209
    Dim gline As String
    Dim d_end As String * 1
    Dim recordno As Long
    Dim a_put As aell
    Dim f_spa As Integer
    Dim f_cell As String
    Dim bs_no As String * 10
    Dim File_num As Integer
    Dim CheckFlag As Boolean
    Dim BaseNoSave() As String
    
    On Error Resume Next
    File_num = 0
    Do While Trim(convert_filename(File_num + 1)) <> ""
       File_num = File_num + 1
    Loop
    d_end = Chr(26)
    recordno = 0
    'Label1.Caption = "正在生成 " + Gsm_Path + "\map\ncell.dbf"
    Gsm_FileName = Gsm_Path + "\map\cell.dbf"
    'lenth = FileLen(Gsm_FileName)
    'nline = lenth / 300
    'bline = Int(nline / 100)
    'percent_step = 1
    'If bline = 0 Then
    '   bline = 1
    '   percent_step = Int(100 / nline + 0.5)
    'End If
    'bss = 1
    'scnline = 0
    ProgressBar1.Value = 0
    
    Gsm_FileName = Gsm_Path + "\map\ncell.dbf"
    Gsm_File2 = Gsm_Path + "\map\ncell.old"
    FileCopy Gsm_FileName, Gsm_File2
    Kill Gsm_FileName
'**************************0
    Gsm_File2 = Gsm_Path + "\ncell.dbf"
    'Open Gsm_File2 For Binary As #3
'**************************
    'Get #3, , header
    'Close #3
    FileCopy Gsm_File2, Gsm_FileName
    Open Gsm_FileName For Binary As #2
    Seek #2, 2210
    'Put #2, , header
    Gsm_File2 = Gsm_Path + "\map\cell"
    mapinfo.Do "open table " + Chr(34) + Gsm_File2 + Chr(34)
    mapinfo.Do "fetch first from cell"
    
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

    a_put.b = " "
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
       a_put.bs_name = space$(10)
       a_put.bs_no = space$(10)
       a_put.ci = space$(5)
       a_put.Lac = space$(5)

       For i = 1 To 16
           a_put.col(i).arfcn_c = space$(3)
           a_put.col(i).ci_c = space$(5)
           a_put.col(i).bsic_c = space$(3)
           a_put.col(i).bs_no_c = space$(10)
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
       a_put.bs_no = gline
       Call bs_lac(gline, a_put.bs_name, a_put.ci, a_put.Lac)
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
                        If Trim(a_put.bs_no) = BaseNoSave(File_num) Then
                           GoTo end_p
                        End If
                    Next
                 Else
                    BaseNoSave(File_num) = Trim(a_put.bs_no)
                 End If
                 BaseNoSave(File_num) = Trim(a_put.bs_no)
                 CheckFlag = True
              End If
              recordno = recordno + 1
              Put #2, , a_put
              GoTo none
           Else
              If UCase(f_cell) = "WO" Then GoTo wo
           End If
           a_put.col(k).bs_no_c = f_cell
           Call bs_ali(f_cell, a_put.col(k).arfcn_c, a_put.col(k).ci_c, a_put.col(k).bsic_c)
       Next
              If CheckFlag = False Then
                 If j > 1 Then
                    For i = 1 To j - 1
                        If Trim(a_put.bs_no) = BaseNoSave(i) Then
                           GoTo end_p
                        End If
                    Next
                 End If
                 BaseNoSave(j) = Trim(a_put.bs_no)
                 CheckFlag = True
              End If
       
       recordno = recordno + 1
       Put #2, , a_put
       DoEvents
       If Convert_Stop = True Then
          Put #2, , d_end
          Seek #2, 5
          Put #2, , recordno
          Close
          Exit Sub
       End If
       GoTo hasfind
another:
    Loop
end_p:
    
    Close #1
  Next
    
    Put #2, , d_end
    Seek #2, 5
    Put #2, , recordno
    Close
    Gsm_File2 = Gsm_Path + "\map\cell"
    mapinfo.Do "close table cell"

'888*********************************************************888
'    Open measname For Input As #1
'    Gsm_FileName = Gsm_Path + "\map\ncell.dbf"
'    Open Gsm_FileName For Binary As #2
'    Do While Not EOF(1)
'       Line Input #1, lines
'       lines = Trim(lines)
'       If Len(lines) > 0 Then
'          f_spa = InStr(lines, " ")
'          If f_spa > 0 Then
'             f_cell = UCase(Left(lines, f_spa - 1))
'             If f_cell = "CELL" Then Exit Do
'          Else
'             If UCase(lines) = "CELL" Then Exit Do
'          End If
'       End If
'    Loop'

'    Do
'       Call getno(bs_no)
'       If end_frag = True Then Exit Do
'       getmbcch
'       If end_frag = True Then Exit Do'
'
'       If mb > 0 Then
'          po = 0
'          pget = 0
 '         Do While Not EOF(2)
 ''            Seek #2, 2210 + po * 223
  '           Get #2, , a_put
'             If Trim(a_put.bs_no) = Trim(bs_no) Then
'                Seek #2, 2210 + po * 223
 '               pget = 1
'                Exit Do
'             End If
'             po = po + 1
'          Loop
'          If pget = 1 Then
'             For j = 1 To mb
'                 For k = 1 To 16
'                     If Trim(a_put.col(k).arfcn_c) = Trim(mbcch(j)) Then
'                        a_put.col(k).meas_c = "Y"
'                        Exit For
'                     End If
'                 Next
'              Next
'              Put #2, , a_put
'          End If
'          Seek #2, 1     '*********************************'
'       End If'

'    Loop

'    Close
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
    Dim nb_data As aell
    Dim new_data As aell
    Dim dbfhead As String * 2209
    Dim recordno As Integer
    Dim linetxt As String
    Dim sameci As String * 5
    Dim largeci As String * 5
    Dim AA As String * 1
    Dim finds As Integer
    Dim FindChar As String * 1
    
    On Error Resume Next
    txtname = convert_filename(1)
    dbfname = Gsm_Path + "\map\ncell.dbf"
    Label1.Caption = "正在生成 " + Gsm_Path + "\map\ncell.dbf"
    Gsm_File2 = Gsm_Path + "\map\cell"
    mapinfo.Do "open table " + Chr(34) + Gsm_File2 + Chr(34)
    mapinfo.Do "fetch first from cell"

    new_data.b = " "
    new_data.bs_name = space$(10)
    new_data.bs_no = space$(10)
    new_data.ci = space$(5)
    new_data.Lac = space$(5)
    For i = 1 To 16
        new_data.col(i).arfcn_c = space$(3)
        new_data.col(i).ci_c = space$(5)
        new_data.col(i).bsic_c = space$(3)
        new_data.col(i).bs_no_c = space$(10)
    Next
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
    
    Gsm_FileName = Gsm_Path + "\ncell.dbf"
    'Open Gsm_FileName For Binary As #1
    'Open dbfname For Binary As #2
    'Get #1, , dbfhead
    'Put #2, , dbfhead
    FileCopy Gsm_FileName, dbfname
    Open dbfname For Binary As #2
    Seek #2, 2210
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

       no = 1
       large = 0
       findspa = InStr(linetxt, FindChar)
       ci = Left(linetxt, findspa - 1)
       nb_data.ci = Right(ci, 4)
       linetxt = Trim(Right(linetxt, Len(linetxt) - findspa))
       Call getnb(linetxt, nb_data.col(no).arfcn_c, nb_data.col(no).ci_c, nb_data.col(no).bsic_c)
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
          If sameci = nb_data.ci Then
             no = no + 1
             If no = 16 Then
                large = 1
             End If
             linetxt = Trim(Right(linetxt, Len(linetxt) - findspa))
             Call getnb(linetxt, nb_data.col(no).arfcn_c, nb_data.col(no).ci_c, nb_data.col(no).bsic_c)
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
                   If largeci <> nb_data.ci Then
                      Exit Do
                   End If
                End If
             End If
             endfile1 = 1
         Loop
    End If
    
s2:

'    nb_data.bs_name = FindArfcn(nb_data.ci, nb_data.ARFCN)   ppppp
       nb_data.bs_name = FindNArfcn(nb_data.ci, nb_data.Lac, nb_data.bs_no)
       nb_data.bs_name = Trim(nb_data.bs_name)
       nb_data.Lac = Trim(nb_data.Lac)
       nb_data.bs_no = Trim(nb_data.bs_no)
       Put #2, , nb_data
       DoEvents
       If Convert_Stop = True Then
          Put #2, , AA
          Seek #2, 5
          Put #2, , recordno
          Close
          Exit Sub
       End If
       If endfile1 = 1 Then
          GoTo s1
       End If
    Loop

s1:
    Put #2, , AA
    Seek #2, 5
    Put #2, , recordno
    Close
    If ProgressBar1.Value < 100 Then
       ProgressBar1.Value = 100
    End If
End Sub

Sub cell_ericsson(sinput1, sinput2, sinput3, sinput4, soutput) 'Convert program
    Dim field_data As cell_stru
    Dim s_data As cell_stru
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
    On Error Resume Next
    load_sam = 0
    load_new = 0
    pos = 0
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
       field_data.photo = space(12)
       field_data.time = space(8)
       field_data.lon = space(12)
       field_data.lat = space(12)
       field_data.Name = space(10)
       field_data.bs_no = space(10)
       field_data.bearing = space(3)
       field_data.ant_type = space(12)
       field_data.ant_angle = space(3)
       field_data.downtilt = space(3)
       field_data.max_bts = space(2)
       field_data.max_ms = space(2)
       field_data.pref = space(2)
       field_data.tch_num = space(2)
       field_data.bsc_stsge = space(7)
       field_data.bsc_type = space(5)
       field_data.bts_type = space(9)
       field_data.power_type = space(3)
       field_data.ci = space(5)
       field_data.ARFCN = space(3)
       field_data.BSIC = space(3)
       field_data.Lac = space(5)
       For i = 1 To 6
           field_data.bcch(i) = space(3)
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
       pos = 0
       Do While Not EOF(2)
          Seek #2, 962 + pos * 153
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
             Seek #3, 962 + pos * 153
             Put #3, , field_data
             GoTo s2
          End If
          pos = pos + 1
       Loop
       Close #2
       Open sinput2 For Binary As #2
       load_new = load_new + 1
'       Seek #3, outlen + wrec * 152
'       Put #3, , field_data
'       wrec = wrec + 1
s2:
       DoEvents
       If Convert_Stop = True Then
          Close
          Exit Sub
       End If
          
    Loop
'    Seek #3, outlen
'    Put #3, , lee
'    recordno = recordno + wrec
'    Seek #3, 5
'    Put #3, , recordno
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
          Seek #3, 962 + po * 153
          Get #3, , s_data
          If UCase(Trim(s_data.bs_no)) = UCase(Trim(bs_no)) Then
             Seek #3, 962 + po * 153
             pget = 1
             Exit Do
          End If
          po = po + 1
       Loop
       If pget = 1 Then
          Call getfield(lines, s_data.power_type)
          Call getfield(lines, s_data.bsc_stsge)
          Call getfield(lines, s_data.bsc_type)
          Call getfield(lines, s_data.bts_type)
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
          Seek #3, 962 + po * 153
          Get #3, , s_data
          If UCase(Trim(s_data.bs_no)) = UCase(Trim(bs_no)) Then
             Seek #3, 962 + po * 153
             pget = 1
             Exit Do
          End If
          po = po + 1
       Loop
       If pget = 1 Then
          For i = 1 To 6
              bc(i) = "   "
          Next
          Call getbcch
          For i = 1 To 6
              s_data.bcch(i) = bc(i)
          Next
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
    For i = 1 To 6
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
        Else
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
    Dim per As Integer, percount As Integer, finds As Integer
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
        percount = 0
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
               If percount = per And ProgressBar1.Value < 90 Then
                  ProgressBar1.Value = ProgressBar1.Value + 1
                  percount = 0
               End If
               percount = percount + 1
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
               If percount = per And ProgressBar1.Value < 90 Then
                  ProgressBar1.Value = ProgressBar1.Value + 1
                  percount = 0
               End If
               percount = percount + 1
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
    
    On Error Resume Next
    Convert_Stop = False
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    Open Gsm_FileName For Binary As #1
    Get #1, 1, MyRecord  ' Read third record.
    Close #1
    If Menu_Flag = 2301 Then
       sts_ericsson
       Unload Me
    End If
    If Menu_Flag = 2302 Then
       If Val(MyRecord.exchange) = 0 Then
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
             Cell_Italtel
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
             sinput1 = convert_filename(1)
             sinput2 = Gsm_Path + "\map\cell.dbf"
             'soutput = Gsm_Path + "\map\cell2.dbf"
             save_mark = False
             Call cell_motorola(sinput1, sinput2)
             'Gsm_FileName = Gsm_Path + "\map\cell2.dbf"
             'Gsm_File2 = Gsm_Path + "\map\cell.dbf"
             'FileCopy Gsm_FileName, Gsm_File2
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
          End If
       End If
    End If
    If Menu_Flag = 2303 Then
       If Val(MyRecord.exchange) = 0 Then
          'gsmname = convert_filename(1)
          'measname = Left(convert_filename(1), Len(convert_filename(1)) - 5) + "rlmfp"
          Ncell_Ericsson
       Else
          If Val(MyRecord.exchange) = 4 Then
             Ncell_Italtel
          Else
             ncell_motorola
          End If
       End If
       Unload Me
    End If
End Sub

Sub Cell_Italtel()
    Dim Old_Cell As cell_stru
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
           For j = 1 To 6
               New_Cell.bcch(j) = ""
           Next
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
              Seek #2, 962 + my_Pos * 153
              Get #2, , Old_Cell
              If Trim(UCase(Old_Cell.bs_no)) = Trim(New_Cell.bs_no) Then
                 Other_Field
                 If end_frag = True Then
                    Exit Do
                 End If
                 Old_Cell.ARFCN = New_Cell.ARFCN
                 Old_Cell.Lac = New_Cell.Lac
                 Old_Cell.ci = New_Cell.ci
                 Old_Cell.BSIC = New_Cell.BSIC
                 Old_Cell.max_bts = New_Cell.max_bts
                 Old_Cell.time = DATE
                 For j = 1 To 6
                     Old_Cell.bcch(j) = New_Cell.bcch(j)
                 Next
                 Seek #2, 962 + my_Pos * 153
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
             New_Cell.bcch(DchNo_num + 1) = Trim(Left(Read_Line, finds - 1))
             DchNo_num = DchNo_num + 1
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             GoTo Dchno1
             finds = InStr(Read_Line, " ")
          Else
             New_Cell.bcch(DchNo_num + 1) = Read_Line
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
             New_Cell.bcch(DchNo_num + 1) = Trim(Left(Read_Line, finds - 1))
             DchNo_num = DchNo_num + 1
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             GoTo Dchno2
             finds = InStr(Read_Line, " ")
          Else
             New_Cell.bcch(DchNo_num + 1) = Read_Line
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
    
    On Error Resume Next
    Gsm_File2 = Gsm_Path + "\map\cell"
    mapinfo.Do "open table " + Chr(34) + Gsm_File2 + Chr(34)
    mapinfo.Do "fetch first from cell"
    Gsm_FileName = Gsm_Path + "\map\ncell.dbf"
    Gsm_File2 = Gsm_Path + "\ncell.dbf"
    FileCopy Gsm_File2, Gsm_FileName
    
    End_Char = Chr$(26)
    recordno = 0
    Open Gsm_FileName For Binary As #2
    Seek #2, 2210
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
        New_Ncell.b = " "
        Do While 1
           scnline = scnline + 1
           If scnline = bss * bline And ProgressBar1.Value < 98 Then
              ProgressBar1.Value = ProgressBar1.Value + percent_step
              bss = bss + 1
           End If
           New_Ncell.bs_name = space$(10)
           New_Ncell.ci = space$(5)
           New_Ncell.Lac = space$(5)
           For j = 1 To 16
               New_Ncell.col(j).arfcn_c = space$(3)
               New_Ncell.col(j).ci_c = space$(5)
               New_Ncell.col(j).bsic_c = space$(3)
               New_Ncell.col(j).bs_no_c = space$(10)
           Next
           If Get_Bs_no = False Then
              Read_Ncell
           End If
           If end_frag = True Then
              Exit Do
           End If
           Call bs_lac(Trim(New_Ncell.bs_no), New_Ncell.bs_name, New_Ncell.ci, New_Ncell.Lac)
           For j = 1 To 16
               Call Ncell_Field(j)
               If Get_Bs_no = True Then
                  Exit For
               End If
           Next
           recordno = recordno + 1
           Put #2, , New_Ncell
           DoEvents
           If Convert_Stop = True Then
              Put #2, , End_Char
              Seek #2, 5
              Put #2, , recordno
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
    Put #2, , End_Char
    Seek #2, 5
    Put #2, , recordno
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
             New_Ncell.col(col_num).ci_c = Trim(Left(Read_Line, finds - 1))
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
          New_Ncell.col(col_num).bsic_c = ncc & bscc
          finds = InStr(Read_Line, "BCCH:")
          If finds > 0 Then
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, ":")
             Read_Line = Trim(Right(Read_Line, Len(Read_Line) - finds))
             finds = InStr(Read_Line, " ")
             If finds > 0 Then
                New_Ncell.col(col_num).arfcn_c = Trim(Left(Read_Line, finds - 1))
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
    Dim field_data As cell_stru
    Dim s_data As cell_stru
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
    
    On Error Resume Next
    pos = 0
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
       field_data.photo = space(12)
       field_data.time = space(8)
       field_data.lon = space(12)
       field_data.lat = space(12)

       field_data.Name = space(10)
       field_data.ci = space(5)
       field_data.BSIC = space(3)
       field_data.ARFCN = space(3)
       field_data.bs_no = space(10)
       field_data.bearing = space(3)
       field_data.ant_type = space(12)
       field_data.ant_angle = space(3)
       field_data.downtilt = space(3)
       field_data.max_bts = space(2)
       field_data.max_ms = space(2)
       field_data.pref = space(2)
       field_data.tch_num = space(2)
       field_data.bsc_stsge = space(7)
       field_data.bsc_type = space(5)
       field_data.bts_type = space(9)
       field_data.power_type = space(3)
       field_data.microcell = space(1)
       For i = 1 To 6
           field_data.bcch(i) = space(3)
           newbcch(i) = space(3)
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
       Call getfield(lines, field_data.pref)
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
       NewPref = field_data.pref
       NewMicrocell = field_data.microcell
        
       j = 1
       FLAG = 0
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
             If j <= 6 Then
                Call getfield(lines, field_data.bcch(j))
                newbcch(j) = field_data.bcch(j)
                j = j + 1
             End If
          Else
             FLAG = 1
             Exit Do
          End If
       Loop

       pos = 0
       Do While Not EOF(2)
          Seek #2, 962 + pos * 153
          Get #2, , s_data
          If s_data.ci = field_data.ci Then
             field_data = s_data
             field_data.Lac = Trim(NewLac)
             field_data.ci = Trim(NewCi)
             field_data.BSIC = Oct(Val(NewBsic))
             field_data.ARFCN = Trim(NewArfcn)
             field_data.max_ms = Trim(NewMax_ms)
             field_data.max_bts = Trim(NewMax_bts)
             field_data.pref = Trim(NewPref)
             field_data.microcell = Trim(NewMicrocell)
             field_data.time = DATE
             For i = 1 To 6
                 field_data.bcch(i) = newbcch(i)
             Next

             Seek #2, 962 + pos * 153
             Put #2, , field_data
             load_new = load_new + 1
             GoTo s2
          End If
          pos = 1 + pos
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

Sub getnb(linetxt, a1, a2, a3)
    Dim finds As Integer
    Dim FindChar As String * 1
    
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
    Dim FileNumber1 As Integer, FileNumber2 As Integer
    Dim LineString As String, LeftString As String
    Dim CorrectFile As Boolean
    
    On Error Resume Next
    
    FileNums = 1
    Gsm_File2 = Gsm_Path + "\sts\tch_sts.dbf"
    Gsm_FileName = Gsm_Path + "\e_tch.dbf"
    FileCopy Gsm_FileName, Gsm_File2
    FileNumber1 = FreeFile
    Open Gsm_File2 For Binary As FileNumber1
    Gsm_File2 = Gsm_Path + "\sts\cch_sts.dbf"
    Gsm_FileName = Gsm_Path + "\e_cch.dbf"
    FileCopy Gsm_FileName, Gsm_File2
    FileNumber2 = FreeFile
    Open Gsm_File2 For Binary As FileNumber2
    FileNumber3 = FreeFile
    Do While Trim(convert_filename(FileNums)) <> ""
       CorrectFile = False
       Open convert_filename(FileNums) For Input As #FileNumber3
       Call ReadFileHead(CorrectFile, LeftString)
       If Not CorrectFile Then
          GoTo NextFile
       End If
       ObjTypeRec (LeftString)
       
       
       
NextFile:
       Close #FileNumber3
       FileNums = FileNums + 1
    Loop
    
    Close #FileNumber1
    Close #FileNumber2
End Sub

Sub ReadFileHead(ResultVar As Boolean, ReturnString As String)
    Dim LineString As String
    Dim i As Integer, ObjTypeNum As Integer
    
    On Error Resume Next
    ResultVar = False
    Do While Not EOF(FileNumber3)
       Line Input #FileNumber3, LineString
       If InStr(UCase(LineString), "FILE DUMP") > 0 Then
          Do While Not EOF(FileNumber3)
             Line Input #FileNumber3, LineString
             If InStr(UCase$(LineString), "CELLTCH") > 0 Or InStr(UCase$(LineString), "CELLCCH") > 0 Then
                ResultVar = True
                Exit Do
             ElseIf Asc(Left(Trim(LineString), 1)) < 58 And Asc(Left(Trim(LineString), 1)) > 47 Then
                Exit Sub
             End If
          Loop
          Exit Do  'Is it necessary?
       End If
    Loop
    If Not ResultVar Then
       Exit Sub
    End If
    For i = 1 To 3
        Do While Not EOF(FileNumber3)
           Line Input #FileNumber3, LineString
           If InStr(UCase(LineString), "WO") = 0 And InStr(UCase(LineString), "END") = 0 Then
              Exit Do
           End If
        Loop
    Next
    ReturnString = Right(LineString, 13)
End Sub

Sub ObjTypeRec(RecString As String)
    
    On Error Resume Next
    if len(recstring)<
    
End Sub

Sub ObjectRec(RecString As String)
    On Error Resume Next
    
End Sub
