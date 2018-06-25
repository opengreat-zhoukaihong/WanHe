VERSION 5.00
Begin VB.Form Base 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "小区选择"
   ClientHeight    =   3090
   ClientLeft      =   3270
   ClientTop       =   1980
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Base.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3090
   ScaleWidth      =   4620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "小区选择"
      Height          =   2190
      Left            =   300
      TabIndex        =   9
      Top             =   720
      Width           =   2820
      Begin VB.CheckBox Check_tch 
         Caption         =   "TCH"
         Height          =   240
         Left            =   1605
         TabIndex        =   5
         Top             =   1755
         Width           =   585
      End
      Begin VB.CheckBox Check_cch 
         Caption         =   "CCH"
         Height          =   240
         Left            =   540
         TabIndex        =   4
         Top             =   1755
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.CheckBox Cell_3 
         Caption         =   "小区-3"
         Height          =   240
         Left            =   285
         TabIndex        =   15
         Top             =   1275
         Width           =   840
      End
      Begin VB.CheckBox Cell_2 
         Caption         =   "小区-2"
         Height          =   240
         Left            =   285
         TabIndex        =   14
         Top             =   885
         Width           =   840
      End
      Begin VB.CheckBox Cell_1 
         Caption         =   "小区-1"
         Height          =   240
         Left            =   300
         TabIndex        =   13
         Top             =   465
         Value           =   1  'Checked
         Width           =   840
      End
      Begin VB.TextBox Arfcn_1 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Enabled         =   0   'False
         Height          =   270
         Left            =   2025
         TabIndex        =   1
         Text            =   "  "
         Top             =   450
         Width           =   495
      End
      Begin VB.TextBox Arfcn_2 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Enabled         =   0   'False
         Height          =   270
         Left            =   2025
         TabIndex        =   2
         Text            =   " "
         Top             =   855
         Width           =   495
      End
      Begin VB.TextBox Arfcn_3 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Enabled         =   0   'False
         Height          =   270
         Left            =   2025
         TabIndex        =   3
         Top             =   1245
         Width           =   495
      End
      Begin VB.Label Cell1Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1440
         TabIndex        =   12
         Top             =   480
         Width           =   510
         WordWrap        =   -1  'True
      End
      Begin VB.Label Cell2Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1425
         TabIndex        =   11
         Top             =   885
         Width           =   540
      End
      Begin VB.Label Cell3Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1425
         TabIndex        =   10
         Top             =   1275
         Width           =   540
      End
   End
   Begin VB.ComboBox Combo1 
      DataField       =   " "
      DataSource      =   " "
      Height          =   300
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   1515
   End
   Begin VB.CommandButton SBSCancel 
      Cancel          =   -1  'True
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   3315
      TabIndex        =   7
      Top             =   1200
      Width           =   1080
   End
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   3315
      TabIndex        =   6
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "基站选择："
      Height          =   180
      Index           =   0
      Left            =   450
      TabIndex        =   8
      Top             =   285
      Width           =   900
   End
End
Attribute VB_Name = "Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
   Dim finds As Integer
   On Error Resume Next
  Select Case Menu_Flag
   Case 151
   Case 41, 461, 462, 463, 464, 465, 466
        If Combo1.Text <> "" Then
           i = 0
           row = Val(mapinfo.eval("tableinfo(base,8)"))
           mapinfo.Do "fetch First from base"
           msg = mapinfo.eval("base.bs_NAME")
           finds = InStr(msg, Chr(0))
           If finds > 0 Then
              msg = Left(msg, finds - 1)
           End If
           For i = 1 To row
               If Trim(UCase(msg)) = Trim(UCase(Combo1.Text)) Then
                  Exit For
               Else
                  mapinfo.Do "fetch next from base"
                  msg = mapinfo.eval("base.bs_name")
                  finds = InStr(msg, Chr(0))
                  If finds > 0 Then
                     msg = Left(msg, finds - 1)
                  End If
               End If
           Next
           While i <= row And msg <> Combo1.Text
              mapinfo.Do "fetch next from base"
              msg = mapinfo.eval("base.bs_NAME")
              finds = InStr(msg, Chr(0))
              If finds > 0 Then
                 msg = Left(msg, finds - 1)
              End If
              i = i + 1
           Wend
           Arfcn_1.Enabled = 1
           Arfcn_2.Enabled = 1
           Arfcn_3.Enabled = 1
           
           SBSOK.Enabled = 1
           SBSCancel.Enabled = 1
           
           If Menu_Flag = 462 Then
              Arfcn_1.Text = mapinfo.eval("base.BSIC_1")
              Arfcn_2.Text = mapinfo.eval("base.BSIC_2")
              Arfcn_3.Text = mapinfo.eval("base.BSIC_3")
           
              Cell1Label.Caption = "BSIC:"
              Cell2Label.Caption = "BSIC:"
              Cell3Label.Caption = "BSIC:"
            Else
              Arfcn_1.Text = mapinfo.eval("base.BCCH_1")
              Arfcn_2.Text = mapinfo.eval("base.BCCH_2")
              Arfcn_3.Text = mapinfo.eval("base.BCCH_3")
           End If
     End If
  End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer, row As Integer
    
    On Error Resume Next
    row = Val(mapinfo.eval("tableinfo(base,8)"))
    mapinfo.Do "fetch First from base"
    For i = 1 To row
        Combo1.AddItem mapinfo.eval("base.bs_NAME")
        mapinfo.Do "fetch next from base"
    Next
    Combo1.ListIndex = 0
    If Menu_Flag = 462 Or Menu_Flag = 463 Or Menu_Flag = 465 Then
       Check_cch.Visible = False
       Check_tch.Visible = False
    End If
       
       
End Sub

Private Sub SBSCancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub SBSOK_Click()
    Dim X, Y As Double
    Dim ci(4), bs_name, bs_overlay  As String
    Dim i, row, BSIC(3), ARFCN(3), Lac As Integer
    Dim Rxlev1, Rxqual1  As Integer
    Dim MyRecord As Record
    Dim moto_type As Boolean
    Dim finds As Integer
    Dim MyNonBcch(1 To 16) As String
    
    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    Open Gsm_FileName For Binary As #1
    Get #1, 1, MyRecord  ' Read third record.
    Close #1
    If Val(MyRecord.exchange) = 1 Or Val(MyRecord.exchange) = 4 Or Val(MyRecord.exchange) = 5 Then
       moto_type = True
    End If
    
 Screen.MousePointer = 11
 If Combo1.Text <> "" Then
    i = 0
    row = Val(mapinfo.eval("tableinfo(base,8)"))
    mapinfo.Do "fetch First from base"
    msg = mapinfo.eval("base.bs_NAME")
    finds = InStr(msg, Chr(0))
    If finds > 0 Then
       msg = Left(msg, finds - 1)
    End If
    msg = Trim(msg)
    While i <= row And msg <> Trim(Combo1.Text)
             mapinfo.Do "fetch next from base"
             msg = mapinfo.eval("base.bs_NAME")
             finds = InStr(msg, Chr(0))
             If finds > 0 Then
                msg = Left(msg, finds - 1)
             End If
             msg = Trim(msg)
             i = i + 1
    Wend
  
        bs_name = Combo1.Text
        ARFCN(1) = Val(mapinfo.eval("base.BCCH_1"))
        ARFCN(2) = Val(mapinfo.eval("base.BCCH_2"))
        ARFCN(3) = Val(mapinfo.eval("base.BCCH_3"))

        BSIC(1) = Val(mapinfo.eval("base.BSIC_1"))
        BSIC(2) = Val(mapinfo.eval("base.BSIC_2"))
        BSIC(3) = Val(mapinfo.eval("base.BSIC_3"))

        ci(1) = CStr(Val(mapinfo.eval("base.ci_1")))
        ci(2) = CStr(Val(mapinfo.eval("base.ci_2")))
        ci(3) = CStr(Val(mapinfo.eval("base.ci_3")))
        If moto_type = False Then
           mapinfo.Do "select * from cell where ci = " + Chr(34) + ci(1) + Chr(34) + " into my_temp"
           ci(1) = mapinfo.eval("my_temp.bs_no")
           mapinfo.Do "close table my_temp"
           mapinfo.Do "select * from cell where ci = " + Chr(34) + ci(2) + Chr(34) + " into my_temp"
           ci(2) = mapinfo.eval("my_temp.bs_no")
           mapinfo.Do "close table my_temp"
           mapinfo.Do "select * from cell where ci = " + Chr(34) + ci(3) + Chr(34) + " into my_temp"
           ci(3) = mapinfo.eval("my_temp.bs_no")
           mapinfo.Do "close table my_temp"
        End If
'        Rxlev1 = Val(Rxlev_1.Text)
'        Rxqual1 = Val(Rxlev_1.Text)

  'SelBase.Hide
  Select Case Menu_Flag
  Case 461
    If Cell_1.Value = 1 And ARFCN(1) <> 0 Then
       If Check_cch.Value = 1 And Check_tch.Value = 0 Then
          mapinfo.Do "select * from cell where ARFCN = " & ARFCN(1) & " into same_arfcn1"
       Else
          If Check_cch.Value = 0 And Check_tch.Value = 1 Then
             'mapinfo.do "select * from cell where Non_bcch_1 = " & ARFCN(1) & " or Non_bcch_2 = " & ARFCN(1) & "or Non_bcch_3 = " & ARFCN(1) & "or Non_bcch_4 = " & ARFCN(1) & "or Non_bcch_5 = " & ARFCN(1) & "or Non_bcch_6 = " & ARFCN(1) & " into same_arfcn1"
             mapinfo.Do "Select * from cell where Like(Non_bcch,""%" & Trim(ARFCN(1)) & "%"","""") = 1 into same_arfcn1"
          Else
             'mapinfo.do "select * from cell where arfcn = " & ARFCN(1) & " or non_bcch_1 = " & ARFCN(1) & " or Non_bcch_2 = " & ARFCN(1) & "or Non_bcch_3 = " & ARFCN(1) & "or Non_bcch_4 = " & ARFCN(1) & "or Non_bcch_5 = " & ARFCN(1) & "or Non_bcch_6 = " & ARFCN(1) & " into same_arfcn1"
             mapinfo.Do "Select * from cell where arfcn = " & ARFCN(1) & " or Like(Non_bcch,""%" & Trim(ARFCN(1)) & "%"","""") = 1 into same_arfcn1"
          End If
       End If

        row = Val(mapinfo.eval("tableinfo(same_arfcn1,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "same_arfcn1" + Chr(34)
        mapinfo.Do msg

        msg = "shade window Frontwindow() same_arfcn1 with ARFCN values " + Chr(34) & ARFCN(1) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg

        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_1 同频分析 " + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   same_arfcn1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
       End If
    End If

    If Cell_2.Value = 1 And ARFCN(2) <> 0 Then
       If Check_cch = 1 And Check_tch = 0 Then
          mapinfo.Do "select * from cell where ARFCN = " & ARFCN(2) & " into same_arfcn2"
       Else
          If Check_cch = 0 And Check_tch = 1 Then
             mapinfo.Do "Select * from cell where Like(Non_bcch,""%" & Trim(ARFCN(2)) & "%"","""") = 1 into same_arfcn2"
          Else
             mapinfo.Do "Select * from cell where arfcn = " & ARFCN(2) & " or Like(Non_bcch,""%" & Trim(ARFCN(2)) & "%"","""") = 1 into same_arfcn2"
          End If
       End If

        row = Val(mapinfo.eval("tableinfo(same_arfcn2,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
        mapinfo.Do "Add Map Auto Layer same_arfcn2"
        msg = "shade window   Frontwindow()   same_arfcn2 with ARFCN values  " + Chr(34) & ARFCN(2) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg

        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_2 同频分析 " + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   same_arfcn2"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
        End If
    End If

    If Cell_3.Value = 1 And ARFCN(3) <> 0 Then
       If Check_cch = 1 And Check_tch = 0 Then
          mapinfo.Do "select * from cell where ARFCN = " & ARFCN(3) & " into same_arfcn3"
       Else
          If Check_cch = 0 And Check_tch = 1 Then
             mapinfo.Do "Select * from cell where Like(Non_bcch,""%" & Trim(ARFCN(3)) & "%"","""") = 1 into same_arfcn3"
          Else
             mapinfo.Do "Select * from cell where arfcn = " & ARFCN(3) & " or Like(Non_bcch,""%" & Trim(ARFCN(3)) & "%"","""") = 1 into same_arfcn3"
          End If
       End If

        row = Val(mapinfo.eval("tableinfo(same_arfcn3,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
        mapinfo.Do "Add Map Auto Layer same_arfcn3"
        msg = "shade window Frontwindow() same_arfcn3 with ARFCN values " + Chr(34) & ARFCN(3) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg

        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_3 同频分析 " + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()   Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
        mapinfo.Do "browse * from   same_arfcn3"
        mapinfo.Do "set window Frontwindow() Position(0,3) Width 8 Height 1 "
       End If
    End If

  Case 462
    If Cell_1.Value = 1 And BSIC(1) <> 0 Then
        mapinfo.Do "select  *  from cell where BSIC = " & BSIC(1) & " into same_BSIC1"
        row = Val(mapinfo.eval("tableinfo(same_bsic1,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "same_BSIC1" + Chr(34)
        mapinfo.Do msg
         
        msg = "shade window   Frontwindow()  same_BSIC1 with BSIC values  " + Chr(34) & BSIC(1) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg
        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_1 同BSIC分析 " + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   same_BSIC1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
      End If
    End If

    If Cell_2.Value = 1 And BSIC(2) <> 0 Then
        mapinfo.Do "select  *  from cell where BSIC = " & BSIC(2) & " into same_BSIC2"
        row = Val(mapinfo.eval("tableinfo(same_bsic2,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "same_BSIC2" + Chr(34)
        mapinfo.Do msg
         
        msg = "shade window   Frontwindow()  same_BSIC2 with BSIC values  " + Chr(34) & BSIC(2) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg
        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_2 同BSIC分析 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   same_BSIC2"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
       End If
    End If

    If Cell_3.Value = 1 And BSIC(3) <> 0 Then
        mapinfo.Do "select  *  from cell where BSIC = " & BSIC(3) & " into same_BSIC3"
        row = Val(mapinfo.eval("tableinfo(same_bsic3,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "same_BSIC3" + Chr(34)
        mapinfo.Do msg
         
        msg = "shade window   Frontwindow()  same_BSIC3 with BSIC values  " + Chr(34) & BSIC(3) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg
        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_3 同BSIC分析 " + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "set window Frontwindow() Position(0,3) Width 8 Height 1 "
        mapinfo.Do "browse * from   same_BSIC3"
       End If
    End If


  Case 463

        Lac = mapinfo.eval("base.lac")

        mapinfo.Do "select  *  from Base where LAC = " & Lac & " into same_lac"
        mapinfo.Do "Add Map Auto Layer same_lac"
        msg = "shade window   Frontwindow()  same_lac with LAC values  " + Chr(34) & Lac & Chr(34) + " Symbol (66,16711935,12)"
        mapinfo.Do msg

        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 同LAC分析 " + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()   Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   same_lac"
        mapinfo.Do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
  Case 464
    If Cell_1.Value = 1 And ARFCN(1) <> 0 Then
        mapinfo.Do "select  *  from cell where ARFCN = " & ARFCN(1) & " AND BSIC = " & BSIC(1) & " into Bsic_arfc1"
        row = Val(mapinfo.eval("tableinfo(bsic_arfc1,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "Bsic_arfc1" + Chr(34)
        mapinfo.Do msg

        msg = "shade window   Frontwindow()  Bsic_arfc1 with ARFCN values  " + Chr(34) & ARFCN(1) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg

        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_1 同频同BSIC 分析 " + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   Bsic_arfc1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
       End If
    End If

    If Cell_2.Value = 1 And ARFCN(2) <> 0 Then
        mapinfo.Do "select  *  from cell where ARFCN = " & ARFCN(2) & " AND BSIC = " & BSIC(2) & " into Bsic_arfc2"
        row = Val(mapinfo.eval("tableinfo(bsic_arfc2,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "Bsic_arfc2" + Chr(34)
        mapinfo.Do msg
        msg = "shade window   Frontwindow()  Bsic_arfc2 with ARFCN values  " + Chr(34) & ARFCN(2) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg
        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_2 同频同BSIC 分析 " + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   Bsic_arfc2"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
       End If
    End If
    If Cell_3.Value = 1 And ARFCN(3) <> 0 Then
        mapinfo.Do "select  *  from cell where ARFCN = " & ARFCN(3) & " AND BSIC = " & BSIC(3) & " into Bsic_arfc3"
        row = Val(mapinfo.eval("tableinfo(bsic_arfc3,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "Bsic_arfc3" + Chr(34)
        mapinfo.Do msg
        msg = "shade window   Frontwindow()  Bsic_arfc3 with ARFCN values  " + Chr(34) & ARFCN(3) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg
        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_3 同频同BSIC 分析 " + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   Bsic_arfc3"
        mapinfo.Do "set window Frontwindow() Position(0,3) Width 8 Height 1 "
       End If
    End If

  Case 465
    Dim NCI(16), ci_str As String

    k = 0
    'mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\map\ncell.tab" + Chr(34)
    If Cell_1.Value = 1 And ARFCN(1) <> 0 Then
       row = Val(mapinfo.eval("tableinfo(cell,8)"))
        mapinfo.Do "Fetch FIRST from CELL"
        'row = Val(mapinfo.eval("tableinfo(nCELL,8)"))
        If moto_type = True Then
           ci_str = mapinfo.eval("CELL.ci")
        Else
           ci_str = mapinfo.eval("cell.bs_no")
        End If
        i = 0
        While ci_str <> ci(1) And i < row
             mapinfo.Do "Fetch next from cell"
             If moto_type = True Then
                ci_str = mapinfo.eval("CELL.ci")
             Else
                ci_str = mapinfo.eval("cell.bs_no")
             End If
             i = i + 1
        Wend
        For i = 1 To 16 Step 1
            If moto_type = True Then
               msg = "cell.ncell" & i
            Else
               msg = "cell.ncell" & i
            End If
            NCI(i) = mapinfo.eval(msg)
            If NCI(i) = "0" Or NCI(i) = "" Then
               k = i - 1
               Exit For
            End If
        Next i
          
        If moto_type = True Then
           msg = "select  *  from cell where ci = " + Chr(34) + NCI(1) + Chr(34)
           For i = 2 To 16 Step 1
               If Trim(NCI(i)) = "" Then
                  Exit For
               End If
               msg = msg + " or ci = " + Chr(34) + NCI(i) + Chr(34)
           Next i
        Else
           msg = "select  *  from cell  where bs_no = " + Chr(34) + NCI(1) + Chr(34)
           For i = 2 To 16 Step 1
               If Trim(NCI(i)) = "" Then
                  Exit For
               End If
               msg = msg + " or bs_no = " + Chr(34) + NCI(i) + Chr(34)
           Next i
        End If
        msg = msg + "  into ncell1"
        mapinfo.Do msg

        row = Val(mapinfo.eval("tableinfo(ncell1,8)"))
        If row < 1 Then
           MsgBox "所查找的小区无相邻小区！", 64, "提示"
        Else

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   ncell1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
        
        If moto_type = True Then
           msg = "select  *  from ncell1  where ncell1 <> " + Chr(34) + ci(1) + Chr(34)
           For i = 2 To 16 Step 1
               msg = msg + " and ncell" & i & " <> " + Chr(34) + ci(1) + Chr(34)
           Next i
        Else
           msg = "select  *  from ncell1  where ncell1 <> " + Chr(34) + ci(1) + Chr(34)
           For i = 2 To 16 Step 1
               msg = msg + " and ncell" & i & " <> " + Chr(34) + ci(1) + Chr(34)
           Next i
        End If
        msg = msg + "  into wrong_ncell1"
        mapinfo.Do msg

        row = Val(mapinfo.eval("tableinfo(wrong_ncell1,8)"))
        If row > 0 Then
           mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
           mapinfo.Do "browse * from   wrong_ncell1"
           mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
        End If
      End If
    End If

    If Cell_2.Value = 1 And ARFCN(2) <> 0 Then
        mapinfo.Do "Fetch FIRST from cELL"
      '  row = Val(mapinfo.eval("tableinfo(nCELL,8)"))
        If moto_type = True Then
           ci_str = mapinfo.eval("cELL.ci")
        Else
           ci_str = mapinfo.eval("cell.bs_no")
        End If
       row = Val(mapinfo.eval("tableinfo(cell,8)"))
        mapinfo.Do "Fetch FIRST from CELL"
        
        i = 0
        While ci_str <> ci(2) And i < row
             mapinfo.Do "Fetch next from cell"
             If moto_type = True Then
                ci_str = mapinfo.eval("CELL.ci")
             Else
                ci_str = mapinfo.eval("cell.bs_no")
             End If
             i = i + 1
        Wend
        
        For i = 1 To 16 Step 1
            If moto_type = True Then
               msg = "cell.ncell" & i
            Else
               msg = "cell.ncell" & i
            End If
            NCI(i) = mapinfo.eval(msg)
            If NCI(i) = "0" Or NCI(i) = "" Then
               k = i - 1
               Exit For
            End If
        Next i
        
        If moto_type = True Then
           msg = "select  *  from cell  where ci = " + Chr(34) + NCI(1) + Chr(34)
           For i = 2 To 16 Step 1
               msg = msg + " or ci = " + Chr(34) + NCI(i) + Chr(34)
           Next i
        Else
           msg = "select  *  from cell  where bs_no = " + Chr(34) + NCI(1) + Chr(34)
           For i = 2 To 16 Step 1
               msg = msg + " or bs_no = " + Chr(34) + NCI(i) + Chr(34)
           Next i
        End If
        msg = msg + "  into ncell2"
        mapinfo.Do msg

        row = Val(mapinfo.eval("tableinfo(ncell2,8)"))
        If row < 1 Then
           MsgBox "所查找的小区无相邻小区！", 64, "提示"
        Else

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   ncell2"
        mapinfo.Do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
        If moto_type = True Then
           msg = "select  *  from ncell2  where ncell1 <> " + Chr(34) + ci(2) + Chr(34)
           For i = 2 To 16 Step 1
               msg = msg + " and ncell" & i & " <> " + Chr(34) + ci(2) + Chr(34)
           Next i
        Else
           msg = "select  *  from ncell2  where ncell1 <> " + Chr(34) + ci(2) + Chr(34)
           For i = 2 To 16 Step 1
               msg = msg + " and ncell" & i & " <> " + Chr(34) + ci(2) + Chr(34)
           Next i
        End If
        msg = msg + "  into wrong_ncell2"
        mapinfo.Do msg

        row = Val(mapinfo.eval("tableinfo(wrong_ncell2,8)"))
        If row > 0 Then
          mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
          mapinfo.Do "browse * from   wrong_ncell2"
          mapinfo.Do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
        End If
      End If
    End If

    If Cell_3.Value = 1 And ARFCN(3) <> 0 Then
        mapinfo.Do "Fetch FIRST from CELL"
        If moto_type = True Then
           ci_str = mapinfo.eval("CELL.ci")
        Else
           ci_str = mapinfo.eval("cell.bs_no")
        End If
        i = 0
       row = Val(mapinfo.eval("tableinfo(cell,8)"))
        mapinfo.Do "Fetch FIRST from CELL"
        
        While ci_str <> ci(3) And i < row
             mapinfo.Do "Fetch next from cell"
             If moto_type = True Then
                ci_str = mapinfo.eval("CELL.ci")
             Else
                ci_str = mapinfo.eval("cell.bs_no")
             End If
             i = i + 1
        Wend

        For i = 1 To 16 Step 1
            If moto_type = True Then
               msg = "cell.ncell" & i
            Else
               msg = "cell.ncell" & i
            End If
            NCI(i) = mapinfo.eval(msg)
            If NCI(i) = "0" Or NCI(i) = "" Then
               k = i - 1
               Exit For
            End If
        Next i
        
        If moto_type = True Then
           msg = "select  *  from cell  where ci = " + Chr(34) + NCI(1) + Chr(34)
           For i = 2 To 16 Step 1
               msg = msg + " or ci = " + Chr(34) + NCI(i) + Chr(34)
           Next i
        Else
           msg = "select  *  from cell  where bs_no = " + Chr(34) + NCI(1) + Chr(34)
           For i = 2 To 16 Step 1
               msg = msg + " or bs_no = " + Chr(34) + NCI(i) + Chr(34)
           Next i
        End If
        msg = msg + "  into ncell3"
        mapinfo.Do msg

        row = Val(mapinfo.eval("tableinfo(ncell3,8)"))
        If row < 1 Then
           MsgBox "所查找的小区无相邻小区！", 64, "提示"
        Else

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   ncell3"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
        If moto_type = True Then
           msg = "select  *  from ncell3  where ncell1 <> " + Chr(34) + ci(3) + Chr(34)
           For i = 2 To 16 Step 1
               msg = msg + " and ncell" & i & " <> " + Chr(34) + ci(3) + Chr(34)
           Next i
        Else
           msg = "select  *  from ncell3  where ncell1 <> " + Chr(34) + ci(3) + Chr(34)
           For i = 2 To 16 Step 1
               msg = msg + " and ncell" & i & " <> " + Chr(34) + ci(3) + Chr(34)
           Next i
        End If
        msg = msg + "  into wrong_ncell3"
        mapinfo.Do msg

        row = Val(mapinfo.eval("tableinfo(wrong_ncell3,8)"))
        If row > 0 Then
           mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
           mapinfo.Do "browse * from   wrong_ncell3"
           mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
        End If
      End If
    End If

  Case 466
     If Cell_1.Value = 1 And ARFCN(1) <> 0 Then
        If Check_cch = 1 And Check_tch = 0 Then
           mapinfo.Do "select * from cell where ABS(Arfcn - " & ARFCN(1) & ")=1 into neighber_arfcn1"
        Else
           If Check_cch = 0 And Check_tch = 1 Then
              'mapinfo.do "select * from cell where ABS(non_bcch_1 - " & ARFCN(1) & ")=1 or ABS(non_bcch_2 - " & ARFCN(1) & ")=1 or ABS(non_bcch_3 - " & ARFCN(1) & ")=1 or ABS(non_bcch_4 - " & ARFCN(1) & ")=1 or ABS(non_bcch_5 - " & ARFCN(1) & ")=1  or ABS(non_bcch_6 - " & ARFCN(1) & ")=1 into neighber_arfcn1"
              mapinfo.Do "Select * from cell where Like(Non_bcch,""%" & Format(Val(ARFCN(1)) + 1) & "%"","""") = 1 or Like(Non_bcch,""%" & Format(Val(ARFCN(1)) - 1) & "%"","""") = 1 into neighber_arfcn1"
           Else
              'mapinfo.do "select * from cell where ABS(Arfcn - " & ARFCN(1) & ")=1 or ABS(non_bcch_1 - " & ARFCN(1) & ")=1 or ABS(non_bcch_2 - " & ARFCN(1) & ")=1 or ABS(non_bcch_3 - " & ARFCN(1) & ")=1 or ABS(non_bcch_4 - " & ARFCN(1) & ")=1 or ABS(non_bcch_5 - " & ARFCN(1) & ")=1  or ABS(non_bcch_6 - " & ARFCN(1) & ")=1 into neighber_arfcn1"
              mapinfo.Do "Select * from cell where ABS(Arfcn - " & ARFCN(1) & ")=1 or Like(Non_bcch,""%" & Format(Val(ARFCN(1)) + 1) & "%"","""") = 1 or Like(Non_bcch,""%" & Format(Val(ARFCN(1)) - 1) & "%"","""") = 1 into neighber_arfcn1"
           End If
       End If

        row = Val(mapinfo.eval("tableinfo(neighber_arfcn1,8)"))
        If row < 1 Then
           MsgBox "所查找的邻频小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "neighber_arfcn1" + Chr(34)
        mapinfo.Do msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   neighber_arfcn1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
       End If
    End If

    If Cell_2.Value = 1 And ARFCN(2) <> 0 Then
       If Check_cch = 1 And Check_tch = 0 Then
          mapinfo.Do "select * from cell where ABS(Arfcn - " & ARFCN(2) & ")=1 into neighber_arfcn2"
       Else
          If Check_cch = 0 And Check_tch = 1 Then
             mapinfo.Do "Select * from cell where Like(Non_bcch,""%" & Format(Val(ARFCN(2)) + 1) & "%"","""") = 1 or Like(Non_bcch,""%" & Format(Val(ARFCN(2)) - 1) & "%"","""") = 1 into neighber_arfcn2"
          Else
             mapinfo.Do "Select * from cell where ABS(Arfcn - " & ARFCN(2) & ")=1 or Like(Non_bcch,""%" & Format(Val(ARFCN(2)) + 1) & "%"","""") = 1 or Like(Non_bcch,""%" & Format(Val(ARFCN(2)) - 1) & "%"","""") = 1 into neighber_arfcn2"
          End If
       End If

        row = Val(mapinfo.eval("tableinfo(neighber_arfcn2,8)"))
        If row < 1 Then
           MsgBox "所查找的邻频小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "neighber_arfcn2" + Chr(34)
        mapinfo.Do msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   neighber_arfcn2"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
       End If
    End If

    If Cell_3.Value = 1 And ARFCN(3) <> 0 Then
       If Check_cch = 1 And Check_tch = 0 Then
          mapinfo.Do "select * from cell where ABS(Arfcn - " & ARFCN(3) & ")=1 into neighber_arfcn3"
       Else
          If Check_cch = 0 And Check_tch = 1 Then
             mapinfo.Do "Select * from cell where Like(Non_bcch,""%" & Format(Val(ARFCN(3)) + 1) & "%"","""") = 1 or Like(Non_bcch,""%" & Format(Val(ARFCN(3)) - 1) & "%"","""") = 1 into neighber_arfcn3"
          Else
             mapinfo.Do "Select * from cell where ABS(Arfcn - " & ARFCN(3) & ")=1 or Like(Non_bcch,""%" & Format(Val(ARFCN(3)) + 1) & "%"","""") = 1 or Like(Non_bcch,""%" & Format(Val(ARFCN(3)) - 1) & "%"","""") = 1 into neighber_arfcn3"
          End If
       End If

        row = Val(mapinfo.eval("tableinfo(neighber_arfcn3,8)"))
        If row < 1 Then
           MsgBox "所查找的邻频小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "neighber_arfcn3" + Chr(34)
        mapinfo.Do msg

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   neighber_arfcn3"
        mapinfo.Do "set window Frontwindow() Position(0,3) Width 8 Height 1 "
       End If
    End If
  End Select
 End If
VER_OUT:
 Screen.MousePointer = 0
 Unload Me
End Sub

