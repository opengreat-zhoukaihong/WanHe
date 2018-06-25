VERSION 5.00
Begin VB.Form SelBase 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "小区选择"
   ClientHeight    =   3465
   ClientLeft      =   2070
   ClientTop       =   885
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Selbase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   4530
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   195
      TabIndex        =   13
      Top             =   2250
      Width           =   4140
      Begin VB.OptionButton Sub_Over 
         Caption         =   "Sub"
         Height          =   300
         Left            =   3240
         TabIndex        =   20
         Top             =   645
         Width           =   705
      End
      Begin VB.OptionButton Full_OVER 
         Caption         =   "Full"
         Height          =   300
         Left            =   3240
         TabIndex        =   19
         Top             =   270
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.TextBox Rxqual_1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1605
         MaxLength       =   1
         TabIndex        =   15
         Text            =   "3"
         Top             =   660
         Width           =   480
      End
      Begin VB.TextBox Rxlev_1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1605
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "17"
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "误码门限(小于):"
         Height          =   180
         Index           =   2
         Left            =   195
         TabIndex        =   18
         Top             =   705
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "场强门限(大于):"
         Height          =   180
         Index           =   1
         Left            =   195
         TabIndex        =   17
         Top             =   330
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "[-93dBm]"
         Height          =   180
         Index           =   3
         Left            =   2130
         TabIndex        =   16
         Top             =   330
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "小区选择"
      Height          =   1590
      Left            =   195
      TabIndex        =   4
      Top             =   600
      Width           =   2745
      Begin VB.TextBox Arfcn_3 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Enabled         =   0   'False
         Height          =   270
         Left            =   1965
         TabIndex        =   10
         Top             =   1170
         Width           =   495
      End
      Begin VB.TextBox Arfcn_2 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Enabled         =   0   'False
         Height          =   270
         Left            =   1965
         TabIndex        =   9
         Text            =   " "
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox Arfcn_1 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Enabled         =   0   'False
         Height          =   270
         Left            =   1965
         TabIndex        =   8
         Text            =   "  "
         Top             =   405
         Width           =   495
      End
      Begin VB.CheckBox Cell_3 
         Caption         =   "小区-3"
         Height          =   240
         Left            =   300
         TabIndex        =   7
         Top             =   1185
         Width           =   840
      End
      Begin VB.CheckBox Cell_2 
         Caption         =   "小区-2"
         Height          =   240
         Left            =   300
         TabIndex        =   6
         Top             =   810
         Width           =   855
      End
      Begin VB.CheckBox Cell_1 
         Caption         =   "小区-1"
         Height          =   240
         Left            =   300
         TabIndex        =   5
         Top             =   405
         Value           =   1  'Checked
         Width           =   840
      End
      Begin VB.Label Cell1Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1365
         TabIndex        =   21
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Cell3Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1365
         TabIndex        =   12
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Cell2Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1365
         TabIndex        =   11
         Top             =   825
         Width           =   540
      End
   End
   Begin VB.ComboBox Combo1 
      DataField       =   " "
      DataSource      =   " "
      Height          =   300
      Left            =   1155
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   1815
   End
   Begin VB.CommandButton SBSCancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   3270
      TabIndex        =   3
      Top             =   1080
      Width           =   1080
   End
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   3270
      TabIndex        =   2
      Top             =   690
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "基站选择："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   210
      Width           =   900
   End
End
Attribute VB_Name = "SelBase"
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
              msg = Trim(Left(msg, finds - 1))
           End If
           While i < row And UCase(msg) <> Trim(UCase(Combo1.Text))
              mapinfo.Do "fetch next from base"
              msg = mapinfo.eval("base.bs_NAME")
              finds = InStr(msg, Chr(0))
              If finds > 0 Then
                 msg = Trim(Left(msg, finds - 1))
              End If
              i = i + 1
           Wend
           Arfcn_1.Enabled = 1
           Arfcn_2.Enabled = 1
           Arfcn_3.Enabled = 1
           
           If Menu_Flag = 461 Or Menu_Flag = 462 Or Menu_Flag = 463 Or Menu_Flag = 463 Or Menu_Flag = 464 Or Menu_Flag = 465 Or Menu_Flag = 466 Or Menu_Flag = 467 Then
'              RXLEV.Enabled = 0
'              Rxqual.Enabled = 0
           Else
'              RXLEV.Enabled = 1
'              Rxqual.Enabled = 1
           End If
           
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
  Dim connect As String

  On Error Resume Next
  i = 0
  row = Val(mapinfo.eval("tableinfo(base,8)"))
  mapinfo.Do "fetch First from base"
  'Combo1.Text = mapinfo.eval("base.bs_NAME")
  While i < row
       Combo1.AddItem mapinfo.eval("base.bs_NAME")
       mapinfo.Do "fetch next from base"
       i = i + 1
  Wend

  Select Case Menu_Flag
   Case 41
          'Rxlev_1.Enabled = 1
          'Rxqual_1.Enabled = 1
          
          Arfcn_1.Text = mapinfo.eval("base.BCCH_1")
          Arfcn_2.Text = mapinfo.eval("base.BCCH_2")
          Arfcn_3.Text = mapinfo.eval("base.BCCH_3")
          Frame2.Visible = False
          Height = 2700
   Case 151, 463
       Cell_1.Enabled = 0
       Cell_2.Enabled = 0
       Cell_3.Enabled = 0

       Rxlev_1.Enabled = 0
       Rxqual_1.Enabled = 0

  Case 462, 464, 465, 466, 467
       Rxlev_1.Enabled = 0
       Rxqual_1.Enabled = 0
  End Select
  mapinfo.Do "fetch First from base"
  Combo1.Text = mapinfo.eval("base.bs_NAME")
  Full_Flag = 1
End Sub

Private Sub Full_OVER_Click()
    Full_Flag = 1
End Sub

Private Sub SBSCancel_Click()
   On Error Resume Next
  Unload Me
End Sub

Private Sub SBSOK_Click()
 Dim X, Y As Double
 Dim finds As Integer
 Dim ci(4), bs_name, bs_overlay  As String
 Dim i, row, BSIC(3), ARFCN(3), Lac As Integer
 Dim Rxlev1, Rxqual1  As Integer
  
 On Error Resume Next
 Screen.MousePointer = 11
 If Combo1.Text <> "" Then
    i = 0
    row = Val(mapinfo.eval("tableinfo(base,8)"))
    mapinfo.Do "fetch First from base"
    msg = mapinfo.eval("base.bs_NAME")
    finds = InStr(msg, Chr(0))
    If finds > 0 Then
       msg = Trim(Left(msg, finds - 1))
    End If
    While i < row And UCase(msg) <> Trim(UCase(Combo1.Text))
             mapinfo.Do "fetch next from base"
             msg = mapinfo.eval("base.bs_NAME")
             finds = InStr(msg, Chr(0))
             If finds > 0 Then
                msg = Trim(Left(msg, finds - 1))
             End If
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
        Rxlev1 = Val(Rxlev_1.Text)
        Rxqual1 = Val(Rxlev_1.Text)

  SelBase.Hide
  Select Case Menu_Flag
   Case 151
        mapinfo.Do "x1= base.lon"
        mapinfo.Do "y1= base.lat"
        
        msg = "set map Center(x1,y1) Smart redraw zoom 4.5 units " + Chr(34) + "km" + Chr(34)
        mapinfo.Do msg
   
   Case 41

        mapinfo.Do "x1= base.lon"
        mapinfo.Do "y1= base.lat"

        Dim ov_flag, ov_flag1 As Integer
        ov_flag1 = 2
        ov_flag = ov_flag1
        Over = 0
        For i = 1 To Val(mapinfo.eval("NumTables()"))
            If Left(UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")), 7) = "OVERLAY" Then
               mapinfo.Do "close table " & mapinfo.eval("tableinfo(" & Format(i) & ",1)")
            End If
        Next
    If Cell_1.Value = 1 And ARFCN(1) <> 0 Then
       bs_overlay = "overlaytemp"
       Over = Over + 1

       mapinfo.Do "fetch first from  " & tblname
       mapinfo.Do "select Time, Lon, Lat,Ci_Serv,Rxlev_f From " + tblname + " where CI_SERV=  " + Chr(34) + ci(1) + Chr(34) + " into " & bs_overlay
       Gsm_FileName = Gsm_Path + "\Overlay1.tab"
       mapinfo.Do "commit table overlaytemp as " + Chr(34) + Gsm_FileName + Chr(34)
       mapinfo.Do "open table " + Chr(34) + Gsm_FileName + Chr(34)
       mapinfo.Do "Alter Table ""overlay1"" ( add Serving Smallint)"
       mapinfo.Do "update overlay1 set serving = 1"
       For i = 1 To 6
           mapinfo.Do "select * from " + tblname + " where (bsic_n" & Format(i) & "= " + Format(BSIC(1)) + " and bcch_n" & Format(i) & "=" + Format(ARFCN(1)) + ") into overlaytemp"
           row = Val(mapinfo.eval("tableinfo(overlaytemp,8)"))
           If row > 0 Then
              mapinfo.Do "insert into overlay1 (col1,col2,col3,col4,col5) select time,lon,lat,ci_serv,rxlev_n" & Format(i) & " from overlaytemp"
           End If
       Next
       mapinfo.Do "commit table overlay1"
       mapinfo.Do "add map auto layer overlay1"
                  msg = " shade window FrontWindow() overlay1 With RXLEV_F "
                  If Legend_Tog = 0 Then
                       'msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,17: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                       msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 35 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                  Else
                       msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                  End If
                  mapinfo.Do msg
                  If legendid = 0 Then     'win95
                      mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"     'win95
                      mapinfo.Do "Create Legend From Window  Frontwindow()"     'win95
                      legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
                  End If     'win95
                  If Legend_Tog = 0 Then
                         'msg = " Title " + Chr(34) + "RxlevFull观测 " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34) + "ascending off ranges Font (""System"",0,8,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                         'msg = " Title " + Chr(34) + "确定小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                         msg = " Title " + Chr(34) + "确定小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                  Else
                         msg = " Title " + Chr(34) + "确定小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                  End If
                  mapinfo.Do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on " & msg
       mapinfo.Do "close table overlaytemp"
     End If

    If Cell_2.Value = 1 And ARFCN(2) <> 0 Then
       bs_overlay = "overlaytemp"
       Over = Over + 1

       mapinfo.Do "fetch first from  " & tblname
       mapinfo.Do "select Time, Lon, Lat,ci_serv, Rxlev_f From " + tblname + " where CI_SERV=  " + Chr(34) + ci(2) + Chr(34) + " into " & bs_overlay
       Gsm_FileName = Gsm_Path + "\Overlay2.tab"
       mapinfo.Do "commit table overlaytemp as " + Chr(34) + Gsm_FileName + Chr(34)
       mapinfo.Do "open table " + Chr(34) + Gsm_FileName + Chr(34)
       mapinfo.Do "Alter Table ""overlay2"" ( add Serving Smallint)"
       mapinfo.Do "update overlay2 set serving = 1"
       
       For i = 1 To 6
           mapinfo.Do "select * from " + tblname + " where (bsic_n" & Format(i) & "= " + Format(BSIC(2)) + " and bcch_n" & Format(i) & "=" + Format(ARFCN(2)) + ") into overlaytemp"
           row = Val(mapinfo.eval("tableinfo(overlaytemp,8)"))
           If row > 0 Then
              mapinfo.Do "insert into overlay2 (col1,col2,col3,col4,col5) select time,lon,lat,ci_serv,rxlev_n" & Format(i) & " from overlaytemp"
           End If
       Next
       mapinfo.Do "commit table overlay2"
       mapinfo.Do "add map auto layer overlay2"
                  msg = " shade window FrontWindow() overlay2 With RXLEV_F "
                  If Legend_Tog = 0 Then
                       'msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,17: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                       msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 35 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                  Else
                       msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                  End If
                  mapinfo.Do msg
                  If legendid = 0 Then     'win95
                      mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"     'win95
                      mapinfo.Do "Create Legend From Window  Frontwindow()"     'win95
                      legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
                  End If     'win95
                  If Legend_Tog = 0 Then
                         'msg = " Title " + Chr(34) + "RxlevFull观测 " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34) + "ascending off ranges Font (""System"",0,8,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                         'msg = " Title " + Chr(34) + "确定小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                         msg = " Title " + Chr(34) + "确定小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                  Else
                         msg = " Title " + Chr(34) + "确定小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                  End If
                  mapinfo.Do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on " & msg
       mapinfo.Do "close table overlaytemp"
     End If

    If Cell_3.Value = 1 And ARFCN(3) <> 0 Then
       bs_overlay = "overlaytemp"
       Over = Over + 1

       mapinfo.Do "fetch first from  " & tblname
       mapinfo.Do "select Time, Lon, Lat,ci_serv, Rxlev_f From " + tblname + " where CI_SERV=  " + Chr(34) + ci(3) + Chr(34) + " into " & bs_overlay
       Gsm_FileName = Gsm_Path + "\Overlay3.tab"
       mapinfo.Do "commit table overlaytemp as " + Chr(34) + Gsm_FileName + Chr(34)
       mapinfo.Do "open table " + Chr(34) + Gsm_FileName + Chr(34)
       mapinfo.Do "Alter Table ""overlay3"" ( add Serving Smallint)"
       mapinfo.Do "update overlay3 set serving = 1"
       
       For i = 1 To 6
           mapinfo.Do "select * from " + tblname + " where (bsic_n" & Format(i) & "= " + Format(BSIC(3)) + " and bcch_n" & Format(i) & "=" + Format(ARFCN(3)) + ") into overlaytemp"
           row = Val(mapinfo.eval("tableinfo(overlaytemp,8)"))
           If row > 0 Then
              mapinfo.Do "insert into overlay3 (col1,col2,col3,col4,col5) select time,lon,lat,ci_serv,rxlev_n" & Format(i) & " from overlaytemp"
           End If
       Next
       mapinfo.Do "commit table overlay3"
       mapinfo.Do "add map auto layer overlay3"
                  msg = " shade window FrontWindow() overlay3 With RXLEV_F "
                  If Legend_Tog = 0 Then
                       'msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,17: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                       msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 35 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                  Else
                       msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                  End If
                  mapinfo.Do msg
                  If legendid = 0 Then     'win95
                      mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"     'win95
                      mapinfo.Do "Create Legend From Window  Frontwindow()"     'win95
                      legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
                  End If     'win95
                  If Legend_Tog = 0 Then
                         'msg = " Title " + Chr(34) + "RxlevFull观测 " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34) + "ascending off ranges Font (""System"",0,8,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                         'msg = " Title " + Chr(34) + "确定小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                         msg = " Title " + Chr(34) + "确定小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                  Else
                         msg = " Title " + Chr(34) + "确定小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                  End If
                  mapinfo.Do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on " & msg
       mapinfo.Do "close table overlaytemp"
     End If

  Case 461

    If Cell_1.Value = 1 And ARFCN(1) <> 0 Then
        mapinfo.Do "select  *  from cell where ARFCN = " & ARFCN(1) & " into same_arfcn1"
        msg = "Add Map Auto Layer " + Chr(34) + "same_arfcn1" + Chr(34)
        mapinfo.Do msg

        msg = "shade window Frontwindow()  same_arfcn1 with ARFCN values  " + Chr(34) & ARFCN(1) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg

        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_1 同频分析 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   same_arfcn1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
    End If

    If Cell_2.Value = 1 And ARFCN(2) <> 0 Then
        mapinfo.Do "select  *  from cell where ARFCN = " & ARFCN(2) & " into same_arfcn2"
        mapinfo.Do "Add Map Auto Layer same_arfcn2"
        msg = "shade window   Frontwindow()   same_arfcn2 with ARFCN values  " + Chr(34) & ARFCN(2) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg

        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_2 同频分析 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   same_arfcn2"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
    End If

    If Cell_3.Value = 1 And ARFCN(3) <> 0 Then
        mapinfo.Do "select  *  from cell where ARFCN = " & ARFCN(3) & " into same_arfcn3"
        mapinfo.Do "Add Map Auto Layer same_arfcn3"
        msg = "shade window   Frontwindow()  same_arfcn3 with ARFCN values  " + Chr(34) & ARFCN(3) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg

        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_3 同频分析 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()   Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
        mapinfo.Do "browse * from   same_arfcn3"
        mapinfo.Do "set window Frontwindow() Position(0,3) Width 8 Height 1 "
    End If

  Case 462
    If Cell_1.Value = 1 And BSIC(1) <> 0 Then
        mapinfo.Do "select  *  from cell where BSIC = " & BSIC(1) & " into same_BSIC1"
        msg = "Add Map Auto Layer " + Chr(34) + "same_BSIC1" + Chr(34)
        mapinfo.Do msg
         
        msg = "shade window   Frontwindow()  same_BSIC1 with BSIC values  " + Chr(34) & BSIC(1) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg
        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_1 同BSIC分析 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   same_BSIC1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
    End If

    If Cell_2.Value = 1 And BSIC(2) <> 0 Then
        mapinfo.Do "select  *  from cell where BSIC = " & BSIC(2) & " into same_BSIC2"
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

    If Cell_3.Value = 1 And BSIC(3) <> 0 Then
        mapinfo.Do "select  *  from cell where BSIC = " & BSIC(3) & " into same_BSIC3"
        msg = "Add Map Auto Layer " + Chr(34) + "same_BSIC3" + Chr(34)
        mapinfo.Do msg
         
        msg = "shade window   Frontwindow()  same_BSIC3 with BSIC values  " + Chr(34) & BSIC(3) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg
        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_3 同BSIC分析 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "set window Frontwindow() Position(0,3) Width 8 Height 1 "
        mapinfo.Do "browse * from   same_BSIC3"
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
        msg = " Title " + Chr(34) + bs_name + " 同LAC分析 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()   Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   same_lac"
        mapinfo.Do "set window Frontwindow() Position(0,4) Width 8 Height 1 "

  Case 464
    If Cell_1.Value = 1 And ARFCN(1) <> 0 Then

        mapinfo.Do "select  *  from cell where ARFCN = " & ARFCN(1) & " AND BSIC = " & BSIC(1) & " into Bsic_arfc1"
        msg = "Add Map Auto Layer " + Chr(34) + "Bsic_arfc1" + Chr(34)
        mapinfo.Do msg

        msg = "shade window   Frontwindow()  Bsic_arfc1 with ARFCN values  " + Chr(34) & ARFCN(1) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg

        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_1 同频同BSIC 分析 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   Bsic_arfc1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
    End If

    If Cell_2.Value = 1 And ARFCN(2) <> 0 Then
        mapinfo.Do "select  *  from cell where ARFCN = " & ARFCN(2) & " AND BSIC = " & BSIC(2) & " into Bsic_arfc2"
        msg = "Add Map Auto Layer " + Chr(34) + "Bsic_arfc2" + Chr(34)
        mapinfo.Do msg
        msg = "shade window   Frontwindow()  Bsic_arfc2 with ARFCN values  " + Chr(34) & ARFCN(2) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg
        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_2 同频同BSIC 分析 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   Bsic_arfc2"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
    End If
    If Cell_3.Value = 1 And ARFCN(3) <> 0 Then
        mapinfo.Do "select  *  from cell where ARFCN = " & ARFCN(3) & " AND BSIC = " & BSIC(3) & " into Bsic_arfc3"
        msg = "Add Map Auto Layer " + Chr(34) + "Bsic_arfc3" + Chr(34)
        mapinfo.Do msg
        msg = "shade window   Frontwindow()  Bsic_arfc3 with ARFCN values  " + Chr(34) & ARFCN(3) & Chr(34) + " Symbol (58,16711935,12)"
        mapinfo.Do msg
        If legendid = 0 Then
                mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                mapinfo.Do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
        End If
        msg = " Title " + Chr(34) + bs_name + " 小区_3 同频同BSIC 分析 " + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
        mapinfo.Do "set legend window   Frontwindow()  Layer prev " & msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   Bsic_arfc3"
        mapinfo.Do "set window Frontwindow() Position(0,3) Width 8 Height 1 "
    End If

  Case 465
    Dim NCI(16), ci_str As String

    k = 0
    mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\map\ncell.tab" + Chr(34)
' On Error GoTo 0
    If Cell_1.Value = 1 And ARFCN(1) <> 0 Then
        mapinfo.Do "Fetch FIRST from nCELL"
        row = Val(mapinfo.eval("tableinfo(nCELL,8)"))
        ci_str = mapinfo.eval("nCELL.ci")
        i = 0
        While ci_str <> ci(1) And i < row
             mapinfo.Do "Fetch next from ncell"
             ci_str = mapinfo.eval("nCELL.ci")
             i = i + 1
        Wend

        For i = 1 To 16 Step 1
            msg = "ncell.ci_" & i
            NCI(i) = mapinfo.eval(msg)
            If NCI(i) = "0" Or NCI(i) = "" Then
               k = i - 1
               Exit For
            End If
        Next i

        msg = "select  *  from ncell  where ci = " + Chr(34) + NCI(1) + Chr(34)
        For i = 2 To k Step 1
            msg = msg + " or ci = " + Chr(34) + NCI(i) + Chr(34)
        Next i
        msg = msg + "  into ncell1"
        mapinfo.Do msg

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   ncell1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "

        msg = "select  *  from ncell1  where ci_1 <> " + Chr(34) + ci(1) + Chr(34)
        For i = 2 To k Step 1
            msg = msg + " and ci_" & i & " <> " + Chr(34) + ci(1) + Chr(34)
        Next i
        msg = msg + "  into wrong_ncell1"
        mapinfo.Do msg

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   wrong_ncell1"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
    End If

    If Cell_2.Value = 1 And ARFCN(2) <> 0 Then
        mapinfo.Do "Fetch FIRST from nCELL"
        row = Val(mapinfo.eval("tableinfo(nCELL,8)"))
        ci_str = mapinfo.eval("nCELL.ci")
        i = 0
        While ci_str <> ci(2) And i < row
             mapinfo.Do "Fetch next from ncell"
             ci_str = mapinfo.eval("nCELL.ci")
             i = i + 1
        Wend
        
        For i = 1 To 16 Step 1
            msg = "ncell.ci_" & i
            NCI(i) = mapinfo.eval(msg)
            If NCI(i) = "0" Or NCI(i) = "" Then
               k = i - 1
               Exit For
            End If
        Next i

        msg = "select  *  from ncell  where ci = " + Chr(34) + NCI(1) + Chr(34)
        For i = 2 To k Step 1
            msg = msg + " or ci = " + Chr(34) + NCI(i) + Chr(34)
        Next i
        msg = msg + "  into ncell1"
        mapinfo.Do msg

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   ncell1"
        mapinfo.Do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
        msg = "select  *  from ncell1  where ci_1 <> " + Chr(34) + ci(2) + Chr(34)
        For i = 2 To k Step 1
            msg = msg + " and ci_" & i & " <> " + Chr(34) + ci(2) + Chr(34)
        Next i
        msg = msg + "  into wrong_ncell1"
        mapinfo.Do msg

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   wrong_ncell1"
        mapinfo.Do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
    End If

    If Cell_3.Value = 1 And ARFCN(3) <> 0 Then
        mapinfo.Do "Fetch FIRST from nCELL"
        row = Val(mapinfo.eval("tableinfo(nCELL,8)"))
        ci_str = mapinfo.eval("nCELL.ci")
        i = 0
        While ci_str <> ci(3) And i < row
             mapinfo.Do "Fetch next from ncell"
             ci_str = mapinfo.eval("nCELL.ci")
             i = i + 1
        Wend

        For i = 1 To k Step 1
            msg = "ncell.ci_" & i
            NCI(i) = mapinfo.eval(msg)
            If NCI(i) = "0" Or NCI(i) = "" Then
               k = i - 1
               Exit For
            End If
        Next i

        msg = "select  *  from ncell  where ci = " + Chr(34) + NCI(1) + Chr(34)
        For i = 2 To k Step 1
            msg = msg + " or ci = " + Chr(34) + NCI(i) + Chr(34)
        Next i
        msg = msg + "  into ncell1"
        mapinfo.Do msg

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   ncell1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
        msg = "select  *  from ncell1  where ci_1 <> " + Chr(34) + ci(3) + Chr(34)
        For i = 2 To k Step 1
            msg = msg + " and ci_" & i & " <> " + Chr(34) + ci(3) + Chr(34)
        Next i
        msg = msg + "  into wrong_ncell1"
        mapinfo.Do msg

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   wrong_ncell1"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
    End If

  Case 466

     If Cell_1.Value = 1 And ARFCN(1) <> 0 Then
        msg = "select  *  from cell  where ABS(Arfcn - " & ARFCN(1) & ")=1 into neighber_arfcn1"
        mapinfo.Do msg
        msg = "Add Map Auto Layer " + Chr(34) + "neighber_arfcn1" + Chr(34)
        mapinfo.Do msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   neighber_arfcn1"
        mapinfo.Do "set window Frontwindow() Position(0,1) Width 8 Height 1 "
    End If

    If Cell_2.Value = 1 And ARFCN(2) <> 0 Then
        mapinfo.Do "select  *  from cell  where ABS(Arfcn - " & ARFCN(2) & ")=1 into neighber_arfcn2"
        msg = "Add Map Auto Layer " + Chr(34) + "neighber_arfcn2" + Chr(34)
        mapinfo.Do msg
        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   neighber_arfcn2"
        mapinfo.Do "set window Frontwindow() Position(0,2) Width 8 Height 1 "
    End If

    If Cell_3.Value = 1 And ARFCN(3) <> 0 Then
        mapinfo.Do "select  *  from cell  where ABS(Arfcn - " & ARFCN(3) & ")=1 into neighber_arfcn3"
        msg = "Add Map Auto Layer " + Chr(34) + "neighber_arfcn3" + Chr(34)
        mapinfo.Do msg

        mapinfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.Do "browse * from   neighber_arfcn3"
        mapinfo.Do "set window Frontwindow() Position(0,3) Width 8 Height 1 "
    End If
  End Select
 End If
VER_OUT:
 Screen.MousePointer = 0
 Unload Me
End Sub

Private Sub Sub_Over_Click()
   Full_Flag = 0
End Sub

