VERSION 5.00
Begin VB.Form SelCond 
   BackColor       =   &H00C0C0C0&
   Caption         =   "条件选择"
   ClientHeight    =   2205
   ClientLeft      =   7800
   ClientTop       =   990
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Selcond.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2205
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Full/Sub选择"
      Height          =   765
      Left            =   165
      TabIndex        =   9
      Top             =   1305
      Width           =   2430
      Begin VB.OptionButton Option4 
         Caption         =   "Full"
         Height          =   300
         Left            =   390
         TabIndex        =   2
         Top             =   345
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Sub"
         Height          =   315
         Left            =   1425
         TabIndex        =   3
         Top             =   330
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "门限"
      Height          =   1080
      Left            =   165
      TabIndex        =   6
      Top             =   120
      Width           =   2430
      Begin VB.TextBox RxLevValue 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1350
         TabIndex        =   0
         Text            =   "17"
         Top             =   300
         Width           =   495
      End
      Begin VB.TextBox RxQualValue 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1350
         TabIndex        =   1
         Text            =   "3"
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "-dBm"
         Height          =   180
         Left            =   1920
         TabIndex        =   11
         Top             =   690
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "-dBm"
         Height          =   180
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxLev:"
         Height          =   180
         Index           =   6
         Left            =   735
         TabIndex        =   8
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxQual:"
         Height          =   180
         Index           =   7
         Left            =   645
         TabIndex        =   7
         Top             =   690
         Width           =   630
      End
   End
   Begin VB.CommandButton SBSCancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   2700
      TabIndex        =   5
      Top             =   615
      Width           =   1080
   End
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   2700
      TabIndex        =   4
      Top             =   225
      Width           =   1080
   End
End
Attribute VB_Name = "SelCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Full_Sub As Integer

Private Sub Form_Load()
    On Error Resume Next
    Full_Sub = 0
    Select Case Menu_Flag
       Case 914, 915, 916, 917
           If Menu_Flag = 915 Then
               RxLevValue.Text = "83"
           Else
               RxLevValue.Text = "93"
           End If
           RxQualValue.Enabled = 0
       Case 441, 442
           Full_Sub = 0
           Frame1.Caption = "差值"
           RxLevValue.Text = "9"
           RxLevValue.Top = 460
           Label1(6).Top = 500
           'RxQualValue.Text = "1"
           RxQualValue.Visible = False
           Label1(7).Visible = False
       Case 9001
           Label1(6).Caption = "主小区 <"
           Label1(7).Caption = "邻小区 <"
           RxLevValue.Text = "83"
           RxQualValue.Text = "83"
           Label2.Visible = True
           Label3.Visible = True
           Me.Caption = "场强覆盖条件选择"
       Case 9002
           Label1(6).Caption = "RxLev >="
           Label1(7).Caption = "RxQual >"
           RxLevValue.Text = "83"
           RxQualValue.Text = "3"
           Label2.Visible = True
       Case 9003
           Label1(6).Caption = "RxQual >="
           Label1(7).Caption = "FER >"
           RxLevValue.Text = "3"
           RxQualValue.Text = "50"
           Label3.Caption = " %"
           Label3.Visible = True
    End Select
    If Menu_Flag = 442 Then
       RxLevValue.Text = "4"
    End If
End Sub

Private Sub Option4_Click()
    On Error Resume Next
    Full_Sub = 0
End Sub

Private Sub Option5_Click()
    On Error Resume Next
    Full_Sub = 1
End Sub

Private Sub SBSCancel_Click()
   On Error Resume Next
   SelCond.Hide
    Unload SelCond
End Sub

Private Sub SBSOK_Click()
  Dim Rxlev1, Rxqual1 As Integer
  Dim blind As String
  
  Dim i, col_num, Maxbsic, Minbsic As Integer
  Dim Name, str As String
  Dim blind_leg(4), Step As Integer
           
    Dim arfcn_field As String
    Dim my_numtable
    Dim CoverFile As String
    Dim cover_arfcn As Integer, cover_rxlev As Integer
    
    Dim NotBlindflag As Boolean
   
  On Error Resume Next
  If Menu_Flag = 9001 Then
     Rxlev1 = 110 - Val(RxLevValue.Text)
     Rxqual1 = 110 - Val(RxQualValue.Text)
  Else
     Rxlev1 = Val(RxLevValue.Text)
     Rxqual1 = Val(RxQualValue.Text)
  End If
  Unload Me

  If tblname <> "" Then
  Select Case Menu_Flag
  Case 42     '该功能已取消，哈哈哈
  
       On Error Resume Next
       Step = Rxlev1 \ 5
       blind_leg(1) = Step
       blind_leg(2) = Step * 2
       blind_leg(3) = Step * 3
       blind_leg(4) = Step * 4

       blind = "blind"

       mapinfo.do "fetch first from  " & tblname
       If Full_Sub = 0 Then
          If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
             Msg = "select * From " + tblname + " where (RXLEV_F <=  " & Rxlev1 & " ) And (val(RXQUAL_F)<= " & Rxqual1 & " ) into  " & blind
          Else
             Msg = "select * From " + tblname + " where (RXLEV_F <=  " & Rxlev1 & " ) And (RXQUAL_F<= " & Rxqual1 & " ) into  " & blind
          End If
       Else
          If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
             Msg = "select * From " + tblname + " where (RXLEV_S <=  " & Rxlev1 & " ) And (val(RXQUAL_S)<= " & Rxqual1 & " ) into  " & blind
          Else
             Msg = "select * From " + tblname + " where (RXLEV_S <=  " & Rxlev1 & " ) And (RXQUAL_S<= " & Rxqual1 & " ) into  " & blind
          End If
       End If
       mapinfo.do Msg
       mapinfo.do "Add Map window FrontWindow() Layer " & blind
       
       If Full_Sub = 0 Then
            Msg = " shade window Frontwindow()  " + blind + " With RXLEV_F   "
       Else
            Msg = " shade window Frontwindow()  " + blind + " With RXLEV_S   "
       End If

       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  " & Rxlev1 & " : " & blind_leg(4) & " Symbol (39,255,8,""MapInfo Cartographic"",0,0) ," & blind_leg(4) & " : " & blind_leg(3) & " Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ," & blind_leg(3) & " : " & blind_leg(2) & " Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ," & blind_leg(2) & " : " & blind_leg(1) & " Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ," & blind_leg(1) & ": 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
       mapinfo.do Msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If

                  Msg = " Title " + Chr(34) + "盲点观测 (0=-110 dBm) " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev  display on shades off symbols on lines off count on" & Msg
     Case 45
             mapinfo.do "fetch first from  " & tblname
             If Full_Sub = 0 Then
'                msg = "select * From " + tblname + " where (RXLEV_F <=  " & Rxlev1 & " ) And (RXQUAL_F >= " & Rxqual1 & " ) into   RxQual_Bad"
                If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                   Msg = "select * From " + tblname + " where ((RXLEV_F <=  " & Rxlev1 & " ) OR (val(RXQUAL_F) >= " & Rxqual1 & " )) and (ta > """") into RxQual_Bad"
                Else
                   Msg = "select * From " + tblname + " where ((RXLEV_F <=  " & Rxlev1 & " ) OR (RXQUAL_F >= " & Rxqual1 & " )) and (ta > """") into RxQual_Bad"
                End If
             Else
'                msg = "select * From " + tblname + " where (RXLEV_S <=  " & Rxlev1 & " ) And (RXQUAL_S >= " & Rxqual1 & " ) into RxQual_Bad"
                If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                   Msg = "select * From " + tblname + " where ((RXLEV_S <=  " & Rxlev1 & " ) OR (val(RXQUAL_S) >= " & Rxqual1 & " )) and (ta > """") into RxQual_Bad"
                Else
                   Msg = "select * From " + tblname + " where ((RXLEV_S <=  " & Rxlev1 & " ) OR (RXQUAL_S >= " & Rxqual1 & " )) and (ta > """") into RxQual_Bad"
                End If
             End If
             mapinfo.do Msg
             If Val(mapinfo.eval("tableinfo(RxQual_Bad,8)")) = 0 Then
                MsgBox "该路段不存在质差点", 64, "提示"
                mapinfo.do "close table RxQual_Bad"
                Exit Sub
             End If
             mapinfo.do "Add Map window FrontWindow() Layer  RxQual_Bad"
             mapinfo.do "set map redraw off"
             mapinfo.do "Set Map Layer 0 Editable On  "
             mapinfo.do "set map redraw on"
             mapinfo.do "Set Style Symbol MakeSymbol(51,16711680,12)"
            strx = "RxQual_Bad.lon"
            stry = "RxQual_Bad.lat"
            row = Val(mapinfo.eval("tableinfo(RxQual_Bad,8)"))
            mapinfo.do " fetch first from  RxQual_Bad"
            i = 1
            While (i < row)
                 Msg = "Create Point(" & strx & "," & stry & ")"
                 mapinfo.do Msg
                 mapinfo.do "fetch  next from RxQual_Bad"
                 i = i + 1
            Wend
            badname = Right(tblname, 7) + "B.tab"
            mapinfo.do "commit table RxQual_Bad as " + Chr(34) + Gsm_Path + "\normal\" + badname + Chr(34)
            
            'mapinfo.do "set map redraw off"
            'mapinfo.do "close  table " & Left(badname, Len(badname) - 4)
            'mapinfo.do "commit  table cosmetic1  as " + Chr(34) + Gsm_Path + "\normal\" + badname + Chr(34)
            'mapinfo.do "Set Map Layer 0 Editable Off  "
            'mapinfo.do "set map redraw on"
            'mapinfo.do " open Table " + Chr(34) + Gsm_Path + "\normal\" + badname + Chr(34) + " Interactive"
            'mapinfo.do " Add Map Layer " & Left(badname, Len(badname) - 4)
            
            MsgBox "文件 " & tblname & " 的质差点保存为：" + Chr(34) + Gsm_Path + "\normal\" + badname + Chr(34), 64, "提示"
     Case 914
           On Error Resume Next
           Msg = "TableInfo(""" & tblname & """, 4)"
           col_num = Val(mapinfo.eval(Msg))
           mapinfo.do "select Min(BSIC_1) From Base into Temp"
           Msg = "Temp.COL1"
           Minbsic = Val(mapinfo.eval(Msg))
           mapinfo.do "select Max(BSIC_1) From Base into Temp"
           Msg = "Temp.COL1"
           Maxbsic = Val(mapinfo.eval(Msg))
'**************************************************************
           my_numtable = mapinfo.eval("NumTables()")
           For i = 1 To my_numtable
               If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "OUTER" Then
                  mapinfo.do "close table outer"
                  Exit For
               End If
           Next
           CoverFile = Gsm_Path + "\outer.tab"
           mapinfo.do "commit table " & tblname & " as " + Chr(34) + CoverFile + Chr(34)
           mapinfo.do "open table " + Chr(34) + CoverFile + Chr(34)
           mapinfo.do "Alter Table ""outer"" ( add outer_arfcn Decimal(3,0) ) "
           mapinfo.do "Alter Table ""outer"" ( add outer_rxlev Decimal(3,0) ) "
           my_row = mapinfo.eval("tableinfo(outer,8)")
           mapinfo.do "fetch first from outer"
           Screen.MousePointer = 11
           For i = 1 To my_row
               cover_arfcn = 0
               cover_rxlev = 0
               For j = 4 To col_num Step 2
                   If Val(mapinfo.eval("outer.col" & j)) <= Rxlev1 And (Val(mapinfo.eval("outer.col" & (j + 1))) > Maxbsic Or Val(mapinfo.eval("outer.col" & (j + 1))) < Minbsic) And Val(mapinfo.eval("outer.col" & (j + 1))) <> 99 Then
                      arfcn_field = mapinfo.eval("Columninfo(""outer"",""COL" & j & """, 1)")
                      If cover_arfcn = 0 Then
                         cover_arfcn = Val(Right(arfcn_field, Len(arfcn_field) - 6))
                         cover_rxlev = Val(mapinfo.eval("outer.col" & j))
                      Else
                         If cover_rxlev > Val(mapinfo.eval("outer.col" & j)) Then
                            cover_arfcn = Val(Right(arfcn_field, Len(arfcn_field) - 6))
                            cover_rxlev = mapinfo.eval("outer.col" & j)
                         End If
                      End If
                   End If
               Next
               mapinfo.do "update outer set outer_arfcn=" + Format(cover_arfcn) + ",outer_rxlev=" + Format(cover_rxlev) + " where rowid =" & i
               mapinfo.do "fetch next from outer"
           Next
           Screen.MousePointer = 0
           mapinfo.do "commit table outer"
           mapinfo.do "close table temp"
           Msg = "Add Map Auto Layer outer"
           mapinfo.do Msg
           Msg = "shade window Frontwindow() outer with  outer_rxlev "
           If Legend_Tog = 0 Then
              Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0)  20: 83 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,83: 93 Symbol (39,16776960,8,""MapInfo Cartographic"",0,0) ,93: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
           Else
              Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  20: 35 Symbol (39,10551200,8,""MapInfo Cartographic"",0,0) ,35: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 65 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,65: 70 Symbol (39,16777072,8,""MapInfo Cartographic"",0,0) ,70: 75 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,75: 80 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,80: 85 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,85: 90 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,90: 95 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,95: 100 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,100: 105 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,105: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
           End If
           mapinfo.do Msg

          If legendid = 0 Then
                mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                mapinfo.do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
          End If
          Msg = " Title " + Chr(34) + " 非本地覆盖" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev>=" + Format(Rxlev1) + "(dBm)" + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
          mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
     
     Case 915
'           Name = "Jam"
           Msg = "TableInfo(""" & tblname & """, 4)"
           On Error Resume Next
           col_num = Val(mapinfo.eval(Msg))
           
           my_numtable = mapinfo.eval("NumTables()")
           For i = 1 To my_numtable
               If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "JAM" Then
                  mapinfo.do "close table jam"
                  Exit For
               End If
           Next
           CoverFile = Gsm_Path + "\jam.tab"
           mapinfo.do "commit table " & tblname & " as " + Chr(34) + CoverFile + Chr(34)
           mapinfo.do "open table " + Chr(34) + CoverFile + Chr(34)
           mapinfo.do "Alter Table ""jam"" ( add jam_rxlev Decimal(3,0) ) "
           my_row = mapinfo.eval("tableinfo(jam,8)")
           mapinfo.do "fetch first from jam"
           Screen.MousePointer = 11
           For i = 1 To my_row
               cover_rxlev = 0
               For j = 4 To col_num Step 2
                   If Val(mapinfo.eval("jam.col" & j)) <= Rxlev1 And Val(mapinfo.eval("jam.col" & (j + 1))) = 99 Then
                      If cover_rxlev = 0 Then
                         cover_rxlev = Val(mapinfo.eval("jam.col" & j))
                      Else
                         If cover_rxlev > Val(mapinfo.eval("jam.col" & j)) Then
                            cover_rxlev = mapinfo.eval("jam.col" & j)
                         End If
                      End If
                   End If
               Next
               mapinfo.do "update jam set jam_rxlev=" + Format(cover_rxlev) + " where rowid =" & i
               mapinfo.do "fetch next from jam"
           Next
           Screen.MousePointer = 0
           mapinfo.do "commit table jam"
           
           Msg = "Add Map Auto Layer jam"
           On Error Resume Next
           mapinfo.do Msg
     
          On Error Resume Next
          Msg = "shade window Frontwindow() jam with jam_rxlev"
          If Legend_Tog = 0 Then
             Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0)  20: 83 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,83: 93 Symbol (39,16776960,8,""MapInfo Cartographic"",0,0) ,93: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
          Else
             Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  20: 35 Symbol (39,10551200,8,""MapInfo Cartographic"",0,0) ,35: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 65 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,65: 70 Symbol (39,16777072,8,""MapInfo Cartographic"",0,0) ,70: 75 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,75: 80 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,80: 85 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,85: 90 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,90: 95 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,95: 100 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,100: 105 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,105: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
          End If
          mapinfo.do Msg

          If legendid = 0 Then
                mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                On Error Resume Next
                mapinfo.do "Create Legend From Window  Frontwindow()"
                On Error Resume Next
                legendid = mapinfo.eval("windowinfo(1009,12)")
          End If
          Msg = " Title " + Chr(34) + " 干扰点观测" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev>=" + Format(Rxlev1) + "(dBm)" + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
          On Error Resume Next
          mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
     
     Case 916
           Name = "Blind"
           Msg = "TableInfo(""" & tblname & """, 4)"
           col_num = mapinfo.eval(Msg)
           On Error Resume Next
           my_numtable = mapinfo.eval("NumTables()")
           For i = 1 To my_numtable
               If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "BLIND" Then
                  mapinfo.do "close table blind"
                  Exit For
               End If
           Next
           CoverFile = Gsm_Path + "\blind.tab"
           mapinfo.do "commit table " & tblname & " as " + Chr(34) + CoverFile + Chr(34)
           mapinfo.do "open table " + Chr(34) + CoverFile + Chr(34)
           mapinfo.do "Alter Table ""blind"" ( add blind_rxlev Decimal(3,0) ) "
           my_row = mapinfo.eval("tableinfo(blind,8)")
           mapinfo.do "fetch first from blind"
           Screen.MousePointer = 11
           For i = 1 To my_row
               cover_rxlev = 0
               NotBlindflag = False
               For j = 4 To col_num Step 2
                   'If Val(mapinfo.eval("blind.col" & (j + 1))) = 99 Then
                   '   If cover_rxlev = 0 Then
                   '      cover_rxlev = Val(mapinfo.eval("blind.col" & j))
                   '   Else
                   '      If cover_rxlev > Val(mapinfo.eval("blind.col" & j)) Then
                   '         cover_rxlev = mapinfo.eval("blind.col" & j)
                   '      End If
                   '   End If
                   'End If
                   'If Val(mapinfo.eval("blind.col" & (j + 1))) <> 99 Then
                   If Val(mapinfo.eval("blind.col" & (j))) < Rxlev1 Then
                       NotBlindflag = True
                       Exit For
                   End If
               Next
               'mapinfo.do "update blind set blind_rxlev=" + Format(cover_rxlev) + " where rowid =" & i
               If Not NotBlindflag Then
                   mapinfo.do "update blind set blind_rxlev=1 where rowid =" & i
               End If
               mapinfo.do "fetch next from blind"
           Next
           Screen.MousePointer = 0
           mapinfo.do "commit table blind"
           
           Msg = "Add Map Auto Layer blind"
           mapinfo.do Msg
     
          Msg = "shade window Frontwindow() blind with blind_rxlev"
          'If Legend_Tog = 0 Then
          '   msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0)  20: 83 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,83: 93 Symbol (39,16776960,8,""MapInfo Cartographic"",0,0) ,93: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
          'Else
          '   msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  20: 35 Symbol (39,10551200,8,""MapInfo Cartographic"",0,0) ,35: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 65 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,65: 70 Symbol (39,16777072,8,""MapInfo Cartographic"",0,0) ,70: 75 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,75: 80 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,80: 85 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,85: 90 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,90: 95 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,95: 100 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,100: 105 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,105: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
          'End If
          Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  1: 1 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
          mapinfo.do Msg
          If legendid = 0 Then
                mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                mapinfo.do "Create Legend From Window  Frontwindow()"
                legendid = mapinfo.eval("windowinfo(1009,12)")
          End If
          Msg = " Title " + Chr(34) + " 盲点观测 (dBm) " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off,""网络盲点"" display on"
          mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg

     Case 917
           On Error Resume Next
           Msg = "TableInfo(""" & tblname & """, 4)"
           col_num = Val(mapinfo.eval(Msg))
           mapinfo.do "select Min(BSIC) From cell into Temp"
           Msg = "Temp.COL1"
           Minbsic = Val(mapinfo.eval(Msg))
           mapinfo.do "select Max(BSIC) From cell into Temp"
           Msg = "Temp.COL1"
           Maxbsic = Val(mapinfo.eval(Msg))
'**************************************************************
           my_numtable = mapinfo.eval("NumTables()")
           For i = 1 To my_numtable
               If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "LOCAL" Then
                  mapinfo.do "close table local"
                  Exit For
               End If
           Next
           CoverFile = Gsm_Path + "\local.tab"
           mapinfo.do "commit table " & tblname & " as " + Chr(34) + CoverFile + Chr(34)
           mapinfo.do "open table " + Chr(34) + CoverFile + Chr(34)
           mapinfo.do "Alter Table ""local"" ( add my_arfcn Decimal(3,0) ) "
           mapinfo.do "Alter Table ""local"" ( add my_rxlev Decimal(3,0) ) "
           my_row = mapinfo.eval("tableinfo(local,8)")
           mapinfo.do "fetch first from local"
           Screen.MousePointer = 11
           For i = 1 To my_row
               cover_arfcn = 0
               cover_rxlev = 0
               For j = 4 To col_num Step 2
                   If Val(mapinfo.eval("local.col" & j)) <= Rxlev1 And Val(mapinfo.eval("local.col" & (j + 1))) <= Maxbsic And Val(mapinfo.eval("local.col" & (j + 1))) >= Minbsic Then
                      arfcn_field = mapinfo.eval("Columninfo(""local"",""COL" & j & """, 1)")
                      If cover_arfcn = 0 Then
                         cover_arfcn = Val(Right(arfcn_field, Len(arfcn_field) - 6))
                         cover_rxlev = Val(mapinfo.eval("local.col" & j))
                      Else
                         If cover_rxlev > Val(mapinfo.eval("local.col" & j)) Then
                            cover_arfcn = Val(Right(arfcn_field, Len(arfcn_field) - 6))
                            cover_rxlev = mapinfo.eval("local.col" & j)
                         End If
                      End If
                   End If
               Next
               mapinfo.do "update local set my_arfcn=" + Format(cover_arfcn) + ",my_rxlev=" + Format(cover_rxlev) + " where rowid =" & i
               mapinfo.do "fetch next from local"
           Next
           Screen.MousePointer = 0
           mapinfo.do "commit table local"
           Msg = "Add Map Auto Layer local"
           mapinfo.do Msg
           Msg = "shade window Frontwindow() local with  my_rxlev "
           If Legend_Tog = 0 Then
              Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0)  20: 83 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,83: 93 Symbol (39,16776960,8,""MapInfo Cartographic"",0,0) ,93: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
           Else
              Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  20: 35 Symbol (39,10551200,8,""MapInfo Cartographic"",0,0) ,35: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 65 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,65: 70 Symbol (39,16777072,8,""MapInfo Cartographic"",0,0) ,70: 75 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,75: 80 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,80: 85 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,85: 90 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,90: 95 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,95: 100 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,100: 105 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,105: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
           End If
           mapinfo.do Msg
           If legendid = 0 Then
              mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
              mapinfo.do "Create Legend From Window  Frontwindow()"
              legendid = mapinfo.eval("windowinfo(1009,12)")
           End If
           Msg = " Title " + Chr(34) + " 本地覆盖 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev >= " + Format(Rxlev1) + " (dBm)" + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
           mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
           mapinfo.do "close table temp"
       Case 9001
             mapinfo.do "fetch first from  " & tblname
             If Full_Sub = 0 Then
                Msg = "select * From " + tblname + " where ((RXLEV_F < " & Rxlev1 & " ) and rxlev_n1 < " & Rxqual1 & " and rxlev_n2 < " & Rxqual1 & " and rxlev_n3 < " & Rxqual1 & " and rxlev_n4 < " & Rxqual1 & " and rxlev_n5 < " & Rxqual1 & " and rxlev_n6 < " & Rxqual1 & ") into Overlay"
             Else
                Msg = "select * From " + tblname + " where ((RXLEV_s < " & Rxlev1 & " ) and rxlev_n1 < " & Rxqual1 & " and rxlev_n2 < " & Rxqual1 & " and rxlev_n3 < " & Rxqual1 & " and rxlev_n4 < " & Rxqual1 & " and rxlev_n5 < " & Rxqual1 & " and rxlev_n6 < " & Rxqual1 & ") into Overlay"
             End If
             mapinfo.do Msg
             If Val(mapinfo.eval("tableinfo(Overlay,8)")) = 0 Then
                MsgBox "该路段不存在网络覆盖盲区", 64, "提示"
                mapinfo.do "close table Overlay"
                Exit Sub
             End If
             mapinfo.do "Add Map window FrontWindow() Layer Overlay"
             If Full_Sub = 0 Then
                Msg = " shade window FrontWindow() Overlay With RXLEV_F "
             Else
                Msg = " shade window FrontWindow() Overlay With RXLEV_s "
             End If
                  'If Legend_Tog = 0 Then
                     Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 150: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                  'Else
                  '   Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 120: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                  'End If
                  mapinfo.do Msg
                    If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                    End If
                'If Full_Sub = 0 Then
                '   msg = " Title " + Chr(34) + "网络覆盖盲区分布 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "主邻小区场强（FULL）<-" & Format(110 - Rxlev1) & "(dBm)" & Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
                'Else
                '   msg = " Title " + Chr(34) + "网络覆盖盲区分布 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "主邻小区场强（SUB）<-" & Format(110 - Rxlev1) & "(dBm)" & Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
                'End If
                
    'If Legend_Tog = 0 Then
        If Full_Sub = 0 Then
            Msg = " Title " + Chr(34) + "网络覆盖盲区分布 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "主邻小区场强（FULL）<-" & Format(110 - Rxlev1) & "(dBm)" & Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""盲区  [" & Format((mapinfo.eval("tableinfo(Overlay,8)")) / mapinfo.eval("tableinfo(" & tblname & ",8)"), "Percent") & "]"" display on "
        Else
            Msg = " Title " + Chr(34) + "网络覆盖盲区分布 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "主邻小区场强（SUB）<-" & Format(110 - Rxlev1) & "(dBm)" & Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""盲区  [" & Format((mapinfo.eval("tableinfo(Overlay,8)")) / mapinfo.eval("tableinfo(" & tblname & ",8)"), "Percent") & "]"" display on "
        End If
    'Else
    '    If Full_Sub = 0 Then
    '        Msg = " Title " + Chr(34) + "网络覆盖盲区分布 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "主邻小区场强（FULL）<-" & Format(110 - Rxlev1) & "(dBm)" & Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
    '    Else
    '        Msg = " Title " + Chr(34) + "网络覆盖盲区分布 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "主邻小区场强（SUB）<-" & Format(110 - Rxlev1) & "(dBm)" & Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
    '    End If
    'End If
                
                
                 mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
       Case 9002
                  If Full_Sub = 0 Then
                     If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                        mapinfo.do "select * from " & tblname & " where rxlev_f >= " & Format(110 - Rxlev1) & " and val(rxqual_f) > " & Format(Rxqual1) & " into NetWorkDisturb"
                     Else
                        mapinfo.do "select * from " & tblname & " where rxlev_f >= " & Format(110 - Rxlev1) & " and rxqual_f > " & Format(Rxqual1) & " into NetWorkDisturb"
                     End If
                  Else
                     If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                        mapinfo.do "select * from " & tblname & " where rxlev_s >= " & Format(110 - Rxlev1) & " and val(rxqual_s) > " & Format(Rxqual1) & " into NetWorkDisturb"
                     Else
                        mapinfo.do "select * from " & tblname & " where rxlev_s >= " & Format(110 - Rxlev1) & " and rxqual_s > " & Format(Rxqual1) & " into NetWorkDisturb"
                     End If
                  End If
             
             If Val(mapinfo.eval("tableinfo(NetWorkDisturb,8)")) = 0 Then
                MsgBox "该路段不存在网络干扰区", 64, "提示"
                mapinfo.do "close table NetWorkDisturb"
                Exit Sub
             End If
                  
                  For i = 1 To mapinfo.eval("NumWindows()")
                      If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then     'win95
                         WinId = mapinfo.eval("windowid(" & i & ")")
                         If WinId = mapinfo.eval("frontwindow()") Then
                            Exit For
                         End If
                      End If
                  Next
                  mapinfo.do "Add Map window " & WinId & " Layer NetWorkDisturb"
                  If Full_Sub = 0 Then
                     Msg = " shade window " & WinId & " NetWorkDisturb With RXLEV_F "
                  Else
                     Msg = " shade window " & WinId & " NetWorkDisturb With RXLEV_s "
                  End If
                  If Legend_Tog = 0 Then
                       'msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,8,""MapInfo Cartographic"",0,0) "
                       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 35 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) "
                  Else
                       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 17 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) "
                  End If
                  mapinfo.do Msg
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window " & WinId
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  If Legend_Tog = 0 Then
                     'msg = " Title " + Chr(34) + "网络干扰区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                     If Full_Sub = 0 Then
                        Msg = " Title " + Chr(34) + "网络干扰区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev_f>=" & Rxlev1 & "(dBm)且RxQual_f>" & Format(Rxqual1) + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                     Else
                        Msg = " Title " + Chr(34) + "网络干扰区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev_s>=" & Rxlev1 & "(dBm)且RxQual_s>" & Format(Rxqual1) + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                     End If
                  Else
                     If Full_Sub = 0 Then
                        Msg = " Title " + Chr(34) + "网络干扰区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev_f>=" & Rxlev1 & "(dBm)且RxQual_f>" & Format(Rxqual1) + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""17 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                     Else
                        Msg = " Title " + Chr(34) + "网络干扰区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "RxLev_s>=" & Rxlev1 & "(dBm)且RxQual_s> " & Format(Rxqual1) + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""17 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                     End If
                  End If
                  mapinfo.do "set legend window " & WinId & " Layer prev display on shades off symbols on lines off count on " & Msg
       Case 9003
                  If Full_Sub = 0 Then
                     If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                        mapinfo.do "select * from " & tblname & " where val(rxqual_f) > " & Format(Rxqual1) & " and val(fer) > " & Format(Rxqual1) & " into DropCallArea"
                     Else
                        mapinfo.do "select * from " & tblname & " where rxqual_f > " & Format(Rxqual1) & " and val(fer) > " & Format(Rxqual1) & " into DropCallArea"
                     End If
                  Else
                     If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                        mapinfo.do "select * from " & tblname & " where val(rxqual_s) > " & Format(Rxqual1) & " and val(fer) > " & Format(Rxqual1) & " into DropCallArea"
                     Else
                        mapinfo.do "select * from " & tblname & " where rxqual_s > " & Format(Rxqual1) & " and val(fer) > " & Format(Rxqual1) & " into DropCallArea"
                     End If
                  End If
             If Val(mapinfo.eval("tableinfo(DropCallArea,8)")) = 0 Then
                MsgBox "该路段不存在网络掉话区", 64, "提示"
                mapinfo.do "close table DropCallArea"
                Exit Sub
             End If
                  
                  For i = 1 To mapinfo.eval("NumWindows()")
                      If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then     'win95
                         WinId = mapinfo.eval("windowid(" & i & ")")
                         If WinId = mapinfo.eval("frontwindow()") Then
                            Exit For
                         End If
                      End If
                  Next
                  mapinfo.do "Add Map window " & WinId & " Layer DropCallArea"
                  
                  If Full_Sub = 0 Then
                     If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                        Msg = " shade window " & WinId & " DropCallArea With RTrim$(LTrim$(RXQUAL_f)) values """" Symbol (41,14737632,8,""MapInfo Cartographic"",0,0) ,""0"" Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,""1"" Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,""2"" Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,""3"" Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,""4"" Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,""5"" Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,""6"" Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,""7"" Symbol (41,16719904,8,""MapInfo Cartographic"",0,0) "
                     Else
                        Msg = " shade window " & WinId & " DropCallArea With RXQUAL_f values 0 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,7 Symbol (41,16719904,8,""MapInfo Cartographic"",0,0),9 Symbol (41,14737632,8,""MapInfo Cartographic"",0,0)"
                     End If
                  Else
                     If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                        Msg = " shade window " & WinId & " DropCallArea With RTrim$(LTrim$(RXQUAL_s)) values """" Symbol (41,14737632,8,""MapInfo Cartographic"",0,0) ,""0"" Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,""1"" Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,""2"" Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,""3"" Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,""4"" Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,""5"" Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,""6"" Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,""7"" Symbol (41,16719904,8,""MapInfo Cartographic"",0,0) "
                     Else
                        Msg = " shade window " & WinId & " DropCallArea With RXQUAL_s values 0 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,7 Symbol (41,16719904,8,""MapInfo Cartographic"",0,0),9 Symbol (41,14737632,8,""MapInfo Cartographic"",0,0)"
                     End If
                  End If
                  mapinfo.do Msg
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window " & WinId
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  If Full_Sub = 0 Then
                     Msg = " Title " + Chr(34) + "网络掉话区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：RxQual_f>" & Format(Rxqual1) & "且FER>" & Format(Rxqual1) + "%" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off"
                  Else
                     Msg = " Title " + Chr(34) + "网络掉话区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：RxQual_s>" & Format(Rxqual1) & "且FER>" & Format(Rxqual1) + "%" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off"
                  End If
                  mapinfo.do "set legend window " & WinId & " Layer prev display on shades off symbols on lines off count on " & Msg
       
    End Select
 End If
End Sub
