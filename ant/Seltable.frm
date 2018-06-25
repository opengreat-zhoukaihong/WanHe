VERSION 5.00
Begin VB.Form SelTable 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择文件"
   ClientHeight    =   2355
   ClientLeft      =   3690
   ClientTop       =   3525
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Seltable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   3840
   Begin VB.CommandButton Cancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   2640
      TabIndex        =   2
      Top             =   645
      Width           =   1080
   End
   Begin VB.CommandButton OK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1080
   End
   Begin VB.ListBox TblList 
      Height          =   1860
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   2220
   End
End
Attribute VB_Name = "SelTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Item As String
    Dim i As Integer
    On Error Resume Next
    mapinfo.do "jj=1"
    TblList.Clear
    TableNum = mapinfo.eval("NumTables()")
    i = 0
    While i < TableNum
       Item = mapinfo.eval("tableinfo(jj,1)")
       If Menu_Flag <> 35 And Menu_Flag <> 81 And Menu_Flag <> 26 Then
          Select Case UCase(Item)
             Case "GSMCELL", "DCSCELL", "CELL", "BASE", "BASE_ADD", "STREET", "AREA", "LANDMARK", "WATER", "MOUNTAIN", "REPETER", "COMPBASE", "COMPCELL", "PUBLIC", "VIP", "POST", "USER_1", "USER_2", "USER_3", "BLOCK", "TOWN", "DUPLABEL", "DUPLICATE"
             Case Else
                 TblList.AddItem Item
             End Select
       Else
          TblList.AddItem Item
       End If
       mapinfo.do " jj=jj+1 "
       i = i + 1
    Wend
End Sub

Private Sub OK_Click()
   Dim MyField As String
   Dim i As Long, row As Long, k As Integer
   Dim CM As String
   Dim Ci_Serv(1 To 100) As String
   Dim j As Integer, Ci_No As Integer
   Dim My_Color As String
   Dim WinId As Variant
   Dim Is_Smooth As Boolean
   Dim Ci_Name(1 To 100) As String
   Dim Sort_Ci_Serv(1 To 100) As String
   Dim my_msg As String
   Dim finds As Integer
   Dim SortTemp As String
   Dim MyPercent As String
   Dim AllRows As Long
   Dim MyPoint As Long
   Dim PreBcch As Integer
   Dim CauseValue() As Integer
   Dim CVString() As String
   Dim QMark As String
   
   On Error Resume Next
   
   mapinfo.do " reload Custom Symbols From " + Chr(34) + Gsm_Path + "\mysymb" + Chr(34)
   tblname = TblList.Text
   SelTable.Hide

   CM = tblname
   If CM = "" Then
      Exit Sub
   End If
   Msg = "tableinfo(" & CM & ",12)"
   
   row = Val(mapinfo.eval(Msg))

   On Error Resume Next
   If tblname <> "" Then
    Select Case Menu_Flag
          Case 26
               If Left(tblname, 6) <> "street" Then
                  MDIMain.FileDialog.filename = ""
                  MDIMain.FileDialog.Filter = "*.tab Files|*.TAB"
                  MDIMain.FileDialog.DefaultExt = "*.TAB"
                  MDIMain.FileDialog.Flags = &H80000
                  MDIMain.FileDialog.InitDir = Gsm_Path
                  MDIMain.FileDialog.ShowSave
                  If MDIMain.FileDialog.filename <> "" Then
                     mapinfo.do "commit table " & tblname & " as " + Chr(34) + MDIMain.FileDialog.filename + Chr(34)
                  End If
                  MDIMain.FileDialog.filename = ""
               End If
          Case 311
                  Msg = " shade window FrontWindow() " + tblname + " With RXLEV_F "
                  If Legend_Tog = 0 Then
                       'msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,17: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                  Else
                       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 120: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                  End If
                  mapinfo.do Msg
                  
                  For i = 1 To mapinfo.eval("NumWindows()")     'win95
                      If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then     'win95
                         WinId = mapinfo.eval("windowid(" & i & ")")     'win95
                         If WinId = mapinfo.eval("frontwindow()") Then
                            Exit For
                         End If
                      End If     'win95
                  Next     'win95

                  If legendid = 0 Then     'win95
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"     'win95
                      mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
                      legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
                  End If     'win95
                  mapinfo.do "select * from " & tblname & " where rxlev_f >0 into Mytemp"
                  AllRows = mapinfo.eval("tableinfo(mytemp,8)")
                  If Legend_Tog = 0 Then
                         'msg = " Title " + Chr(34) + "RxlevFull观测 " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34) + "ascending off ranges Font (""System"",0,8,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                         'msg = " Title " + Chr(34) + "RxlevFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                         If AllRows > 0 Then
                            mapinfo.do "select * from " & tblname & " where rxlev_f >0 and rxlev_f <15 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = " Title " + Chr(34) + "RxlevFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)  [" & MyPercent & "]"" display on ,"
                            'Change Legend
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=15 and rxlev_f <25 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """15 至 25 (-95至-85dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=25 and rxlev_f <35 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """25 至 35 (-85至-75dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=35 and rxlev_f <120 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """35 以上  (大于-75dBm)  [" & MyPercent & "]"" display on "
                         Else
                            Msg = " Title " + Chr(34) + "RxlevFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                            'Change Legend
                         End If
                  Else
                         If AllRows > 0 Then
                            mapinfo.do "select * from " & tblname & " where rxlev_f >0 and rxlev_f <5 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = " Title " + Chr(34) + "RxlevFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5  (-110至-105dBm)  [" & MyPercent & "]"" display on ,"
                            'Change Legend
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=5 and rxlev_f <10 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """5 至 10 (-105至-100dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=10 and rxlev_f <15 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """10 至 15 (-100至-95dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=15 and rxlev_f <20 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """15 至 20 (-95至-90dBm)  [" & MyPercent & "]"" display on ,"
                            
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=20 and rxlev_f <25 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """20 至 25 (-90至-85dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=25 and rxlev_f <30 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """25 至 30 (-85至-80dBm)  [" & MyPercent & "]"" display on ,"
                            
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=30 and rxlev_f <35 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """30 至 35 (-80至-75dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=35 and rxlev_f <40 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """35 至 40 (-75至-70dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=40 and rxlev_f <45 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """40 至 45 (-70至-65dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=45 and rxlev_f <50 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """45 至 50 (-65至-60dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=50 and rxlev_f <63 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """50 至 63 (-60至-47dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where rxlev_f >=63 and rxlev_f <120 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """63 以上  (大于-47dBm)  [" & MyPercent & "]"" display on "
                         
                         Else
                            Msg = " Title " + Chr(34) + "RxlevFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                            'Change Legend
                         End If
                  End If
                  mapinfo.do "close table mytemp"
                  mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on " & Msg
                mapinfo.do "set map redraw off"
                mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                mapinfo.do "set map redraw on"

          Case 3131, 83131
'                  Msg = " shade window FrontWindow() " + TblName + " With BCCH_SERV  ignore 0  values 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9 ,10 ,11 ,12 ,13 ,14 ,15 ,16 ,17 ,18 ,19 ,20 ,21 ,22 ,23 ,24 ,25 ,26 ,27 ,28 ,29 ,30 ,31 ,32 ,33 ,34 ,35 ,36 ,37 ,38 ,39 ,40 ,41 ,42 ,43 ,44 ,45 ,46 ,47 ,48 ,49 ,50 ,51 ,52 ,53 ,54 ,55 ,56 ,57 ,58 ,59 ,60 ,61 ,62 ,63 ,64 ,65 ,66 ,67 ,68 ,69 ,70 ,71 ,72 ,73 ,74 ,75 ,76 ,77 ,78 ,79 ,80 ,81 ,82 ,83 ,84 ,85 ,86 ,87 ,88 ,89 ,90 ,91 ,92 ,93 ,94 ,95 ,96 ,97 ,98 ,99 ,100 ,101 ,102 ,103 ,104 ,105 ,106 ,107 ,108 ,109 ,110 ,111 ,112 ,113 ,114 ,115 ,116 ,117 ,118 ,119 ,120 ,121 ,122 ,123 ,124 default Symbol Symbol (34,0,12) "
                If GSMDCSBCCH = 0 Then
                    If Menu_Flag = 3131 Then
                       Msg = " shade window FrontWindow() " + tblname + " With BCCH_SERV "
                    Else
                       Msg = " shade window FrontWindow() " + tblname + " With arfcn_2 "
                    End If
                Else
                    If GSMDCSBCCH = 1 Then
                        mapinfo.do "select * from " & tblname & " where BCCH_SERV<125 into SelNet"
                    Else
                        mapinfo.do "select * from " & tblname & " where BCCH_SERV>124 into SelNet"
                    End If
                    mapinfo.do "Add Map window FrontWindow() Layer  SelNet"
                    If Menu_Flag = 3131 Then
                       Msg = " shade window FrontWindow() SelNet With BCCH_SERV "
                    Else
                       Msg = " shade window FrontWindow() SelNet With arfcn_2 "
                    End If
                End If
                   Msg = Msg + "ignore 0 values  1 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "2 Symbol (33,65280,8,""MapInfo Cartographic"",0,0) ,3 Symbol (33,255,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "4 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "6 Symbol (33,65535,8,""MapInfo Cartographic"",0,0) ,7 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "8 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),9 Symbol (33,128,8,""MapInfo Cartographic"",0,0),10 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),11 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "12 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),13 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),14 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),15 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "16 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),17 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),18 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),19 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "20 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),21 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),22 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),23 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "24 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),25 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),26 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),27 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "28 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),29 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),30 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "31 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),32 Symbol (33,6324320,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "33 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),34 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),35 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),36 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "37 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),38 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),39 Symbol (33,255,8,""MapInfo Cartographic"",0,0),40 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "41 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),42 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),43 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),44 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "45 Symbol (33,128,8,""MapInfo Cartographic"",0,0),46 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),47 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),48 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "49 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),50 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),51 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),52 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "53 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),54 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),55 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),56 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "57 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),58 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),59 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),60 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "61 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),62 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),63 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "64 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),65 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),66 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "67 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),68 Symbol (33"
                   Msg = Msg + ",6324320,8,""MapInfo Cartographic"",0,0),69 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "70 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),71 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),72 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),73 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "74 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),75 Symbol (33,255,8,""MapInfo Cartographic"",0,0),76 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),77 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "78 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),79 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),80 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),81 Symbol (33,128,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "82 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),83 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),84 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),85 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "86 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),87 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),88 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),89 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "90 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),91 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),92 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),93 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "94 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),95 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),96 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),97 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "98 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),99 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),100 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),101 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "102 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),103 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),104 Symbol (33,6324320,8,""MapInfo Cartographic"",0,0),105 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "106 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),107 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),108 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),109 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "110 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),111 Symbol (33,255,8,""MapInfo Cartographic"",0,0),112 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),113 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "114 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),115 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),116 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),117 Symbol (33,128,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "118 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),119 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),120 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),121 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "122 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),123 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),124 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),"
                   
                   For i = 512 To 885
                       Msg = Msg & Format(i) & " Symbol (33," & Format(MyRndColor(i - 512)) & ",8,""MapInfo Cartographic"",0,0),"
                   Next
                   Msg = Left(Msg, Len(Msg) - 1)
                   
                   mapinfo.do Msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  If Menu_Flag = 3131 Then
                  'Change Legend
                     Select Case GSMDCSBCCH
                        Case 0
                            Msg = " Title " + Chr(34) + "BCCH/SDCCH观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle ""GSM和DCS网"" Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off"
                        Case 1
                            Msg = " Title " + Chr(34) + "BCCH/SDCCH观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle ""GSM网"" Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off"
                        Case 2
                            Msg = " Title " + Chr(34) + "BCCH/SDCCH观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle ""DCS网"" Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off"
                     End Select
                     'msg = " Title " + Chr(34) + "BCCH/SDCCH观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle " + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
                  Else
                     Msg = " Title " + Chr(34) + "第二手机频率观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle " + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off"
                  End If
                  mapinfo.do "set legend window FrontWindow() Layer prev " & Msg
          Case 3132
                   If GSMDCSBCCH = 0 Then
                      Msg = " shade window FrontWindow() " + tblname + " With num_dch "
                   Else
                        If GSMDCSBCCH = 1 Then
                            mapinfo.do "select * from " & tblname & " where val(num_dch)<125 into SelNet"
                        Else
                            mapinfo.do "select * from " & tblname & " where val(num_dch)>124 into SelNet"
                        End If
                        mapinfo.do "Add Map window FrontWindow() Layer SelNet"
                        Msg = " shade window FrontWindow() SelNet With num_dch "
                   End If
                   Msg = Msg + "ignore """" values 1 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "2 Symbol (33,65280,8,""MapInfo Cartographic"",0,0) ,3 Symbol (33,255,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "4 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "6 Symbol (33,65535,8,""MapInfo Cartographic"",0,0) ,7 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "8 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),9 Symbol (33,128,8,""MapInfo Cartographic"",0,0),10 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),11 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "12 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),13 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),14 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),15 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "16 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),17 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),18 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),19 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "20 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),21 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),22 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),23 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "24 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),25 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),26 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),27 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "28 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),29 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),30 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "31 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),32 Symbol (33,6324320,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "33 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),34 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),35 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),36 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "37 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),38 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),39 Symbol (33,255,8,""MapInfo Cartographic"",0,0),40 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "41 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),42 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),43 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),44 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "45 Symbol (33,128,8,""MapInfo Cartographic"",0,0),46 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),47 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),48 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "49 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),50 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),51 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),52 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "53 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),54 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),55 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),56 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "57 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),58 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),59 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),60 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "61 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),62 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),63 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "64 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),65 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),66 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "67 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),68 Symbol (33"
                   Msg = Msg + ",6324320,8,""MapInfo Cartographic"",0,0),69 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "70 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),71 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),72 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),73 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "74 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),75 Symbol (33,255,8,""MapInfo Cartographic"",0,0),76 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),77 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "78 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),79 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),80 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),81 Symbol (33,128,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "82 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),83 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),84 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),85 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "86 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),87 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),88 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),89 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "90 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),91 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),92 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),93 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "94 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),95 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),96 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),97 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "98 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),99 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),100 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),101 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "102 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),103 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),104 Symbol (33,6324320,8,""MapInfo Cartographic"",0,0),105 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "106 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),107 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),108 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),109 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "110 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),111 Symbol (33,255,8,""MapInfo Cartographic"",0,0),112 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),113 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "114 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),115 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),116 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),117 Symbol (33,128,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "118 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),119 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),120 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),121 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "122 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),123 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),124 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),"
                   
                   For i = 512 To 885
                       Msg = Msg & Format(i) & " Symbol (33," & Format(MyRndColor(i - 512)) & ",8,""MapInfo Cartographic"",0,0),"
                   Next
                   Msg = Left(Msg, Len(Msg) - 1)
                   
                   mapinfo.do Msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                     Select Case GSMDCSBCCH
                        Case 0
                            Msg = " Title " + Chr(34) + "TCH观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle ""GSM和DCS网"" Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off"
                        Case 1
                            Msg = " Title " + Chr(34) + "TCH观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle ""GSM网"" Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off"
                        Case 2
                            Msg = " Title " + Chr(34) + "TCH观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle ""DCS网"" Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off"
                     End Select
                  
                  'msg = " Title " + Chr(34) + "TCH观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle " + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow() Layer prev " & Msg
          Case 3133, 83133
                  mapinfo.do "set map redraw off"
                  If Menu_Flag = 3133 Then
                     If GSMDCSBCCH = 0 Then
                        mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                     ElseIf GSMDCSBCCH = 1 Then
                        mapinfo.do "select * from " & tblname & " where BCCH_SERV<125 into SelNet"
                        mapinfo.do "Add Map window FrontWindow() Layer  SelNet"
                        mapinfo.do "Set Map Layer SelNet Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                     Else
                        mapinfo.do "select * from " & tblname & " where BCCH_SERV>124 into SelNet"
                        mapinfo.do "Add Map window FrontWindow() Layer  SelNet"
                        mapinfo.do "Set Map Layer SelNet Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                     End If
                  Else
                     If GSMDCSBCCH = 0 Then
                        mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With arfcn_2 Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                     ElseIf GSMDCSBCCH = 1 Then
                        mapinfo.do "select * from " & tblname & " where BCCH_SERV<125 into SelNet"
                        mapinfo.do "Add Map window FrontWindow() Layer  SelNet"
                        mapinfo.do "Set Map Layer SelNet Label Visibility Font (""Arial"",257,8,8421376,16777215) With arfcn_2 Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                     Else
                        mapinfo.do "Set Map Layer SelNet Label Visibility Font (""Arial"",257,8,8421376,16777215) With arfcn_2 Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                     End If
                  End If
                  mapinfo.do "set map redraw on"
          Case 3134
                  mapinfo.do "set map redraw off"
                  If GSMDCSBCCH = 0 Then
                     mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8388736,16777215) With num_dch Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
                  ElseIf GSMDCSBCCH = 1 Then
                     mapinfo.do "select * from " & tblname & " where val(num_dch)<125 into SelNet"
                     mapinfo.do "Add Map window FrontWindow() Layer  SelNet"
                     mapinfo.do "Set Map Layer SelNet Label Visibility Font (""Arial"",257,8,8388736,16777215) With num_dch Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
                  Else
                     mapinfo.do "select * from " & tblname & " where val(num_dch)>124 into SelNet"
                     mapinfo.do "Add Map window FrontWindow() Layer  SelNet"
                     mapinfo.do "Set Map Layer SelNet Label Visibility Font (""Arial"",257,8,8388736,16777215) With num_dch Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
                  End If
                  mapinfo.do "set map redraw on"

          Case 314, 88311, 88314
               If Menu_Flag = 314 Then
                  MyField = " rxlev_s "
               ElseIf Menu_Flag = 88311 Then
                  MyField = " rxlev_f_2 "
               Else
                  MyField = " rxlev_s_2 "
               End If
                  
                  Msg = " shade  window FrontWindow() " + tblname + "  With " & MyField
                  If Legend_Tog = 0 Then
                       'msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,17: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                  Else
                       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  120: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                  End If
                  mapinfo.do Msg
                  
                  mapinfo.do "select * from " & tblname & " where " & MyField & ">0 into Mytemp"
                  AllRows = mapinfo.eval("tableinfo(mytemp,8)")

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  If Legend_Tog = 0 Then
                         'msg = " Title " + Chr(34) + "RxlevSub观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                         If AllRows > 0 Then
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >0 and " & MyField & " <15 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            'Change Legend
                            If Menu_Flag = 314 Then
                               Msg = " Title " + Chr(34) + "RxlevSub观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)  [" & MyPercent & "]"" display on ,"
                            ElseIf Menu_Flag = 88314 Then
                               Msg = " Title " + Chr(34) + "第二手机RxlevSub观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)  [" & MyPercent & "]"" display on ,"
                            Else
                               Msg = " Title " + Chr(34) + "第二手机RxlevFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)  [" & MyPercent & "]"" display on ,"
                            End If
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=15 and " & MyField & " <25 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """15 至 25 (-95至-85dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=25 and " & MyField & " <35 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """25 至 35 (-85至-75dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=35 and " & MyField & " <120 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """35 以上  (大于-75dBm)  [" & MyPercent & "]"" display on "
                         Else
                            If Menu_Flag = 314 Then
                               Msg = " Title " + Chr(34) + "RxlevSub观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                            ElseIf Menu_Flag = 88314 Then
                               Msg = " Title " + Chr(34) + "第二手机RxlevSub观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                            Else
                               Msg = " Title " + Chr(34) + "第二手机RxlevFull观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                            End If
                         End If
                  Else
                         If AllRows > 0 Then
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >0 and " & MyField & " <5 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            If Menu_Flag = 314 Then
                               Msg = " Title " + Chr(34) + "RxlevSub观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5  (-110至-105dBm)  [" & MyPercent & "]"" display on ,"
                            ElseIf Menu_Flag = 88314 Then
                               Msg = " Title " + Chr(34) + "第二手机RxlevSub观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5  (-110至-105dBm)  [" & MyPercent & "]"" display on ,"
                            Else
                               Msg = " Title " + Chr(34) + "第二手机RxlevFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5  (-110至-105dBm)  [" & MyPercent & "]"" display on ,"
                            End If
                               
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=5 and " & MyField & " <10 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """5 至 10 (-105至-100dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=10 and " & MyField & " <15 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """10 至 15 (-100至-95dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=15 and " & MyField & " <20 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """15 至 20 (-95至-90dBm)  [" & MyPercent & "]"" display on ,"
                            
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=20 and " & MyField & " <25 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """20 至 25 (-90至-85dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=25 and " & MyField & " <30 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """25 至 30 (-85至-80dBm)  [" & MyPercent & "]"" display on ,"
                            
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=30 and " & MyField & " <35 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """30 至 35 (-80至-75dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=35 and " & MyField & " <40 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """35 至 40 (-75至-70dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=40 and " & MyField & " <45 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """40 至 45 (-70至-65dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=45 and " & MyField & " <50 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """45 至 50 (-65至-60dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=50 and " & MyField & " <63 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """50 至 63 (-60至-47dBm)  [" & MyPercent & "]"" display on ,"
                            mapinfo.do "select * from " & tblname & " where " & MyField & " >=63 and " & MyField & " <120 into Mytemp"
                            MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                            MyPercent = Format(MyPoint / AllRows, "Percent")
                            Msg = Msg + """63 以上  (大于-47dBm)  [" & MyPercent & "]"" display on "
                         Else
                            If Menu_Flag = 314 Then
                               Msg = " Title " + Chr(34) + "RxlevSub观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV  标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                            ElseIf Menu_Flag = 88314 Then
                               Msg = " Title " + Chr(34) + "第二手机RxlevSub观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                            Else
                               Msg = " Title " + Chr(34) + "第二手机RxlevFull观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                            End If
                               
                         End If
                  End If
                  mapinfo.do "close table mytemp"
                  mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on" & Msg
                  If Menu_Flag = 314 Then
                        mapinfo.do "set map redraw off"
                        mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                        mapinfo.do "set map redraw on"
                  End If
          Case 315, 88315, 88312, 312

               If Menu_Flag = 315 Then
                  If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then    'Character
                     MyField = " RTrim$(LTrim$(rxqual_s)) "
                  Else
                     MyField = " rxqual_s "
                  End If
               ElseIf Menu_Flag = 312 Then
                  If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                     MyField = " RTrim$(LTrim$(rxqual_f)) "
                  Else
                     MyField = " rxqual_f "
                  End If
               ElseIf Menu_Flag = 88315 Then
                  If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                     MyField = " RTrim$(LTrim$(rxquql_s_2)) "
                  Else
                     MyField = " rxquql_s_2 "
                  End If
               Else
                  If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                     MyField = " RTrim$(LTrim$(rxquql_f_2)) "
                  Else
                     MyField = " rxquql_f_2 "
                  End If
               End If
                
                If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then       'COL_TYPE_CHAR
                   QMark = """"
                   Msg = " shade window FrontWindow() " + tblname + " With " & MyField & " values """" Symbol (41,14737632,8,""MapInfo Cartographic"",0,0) ,""0"" Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,""1"" Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,""2"" Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,""3"" Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,""4"" Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,""5"" Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,""6"" Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,""7"" Symbol (41,16719904,8,""MapInfo Cartographic"",0,0) "
                Else
                   Msg = " shade window FrontWindow() " + tblname + " With " & MyField & " values 0 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,7 Symbol (41,16719904,8,""MapInfo Cartographic"",0,0),9 Symbol (41,14737632,8,""MapInfo Cartographic"",0,0)"
                End If
                  'msg = " shade window FrontWindow() " + tblname + " With " & MyField & " values 0 Symbol (41,65280,8,""MapInfo Cartographic"",0,0) ,1 Symbol (41,65535,8,""MapInfo Cartographic"",0,0) ,2 Symbol (41,8404992,8,""MapInfo Cartographic"",0,0) ,3 Symbol (41,255,8,""MapInfo Cartographic"",0,0) ,4 Symbol (41,12615935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (41,16756912,8,""MapInfo Cartographic"",0,0) ,6 Symbol (41,16711935,8,""MapInfo Cartographic"",0,0) ,7 Symbol (41,16719904,8,""MapInfo Cartographic"",0,0)"
                  mapinfo.do Msg
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  
                  AllRows = mapinfo.eval("tableinfo(" & tblname & ",8)")
                  If AllRows > 0 Then
                     If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                        mapinfo.do "select * from " & tblname & " where " & MyField & " = """" into Mytemp"
                        MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                        MyPercent = Format(MyPoint / AllRows, "Percent")
                        'Msg = Msg + """IDLE  [" & MyPercent & "]"" display on ,"
                        If Menu_Flag = 315 Then
                           Msg = " Title " + Chr(34) + "RxQualSub观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""IDLE  [" & MyPercent & "]"" display on ,"
                        ElseIf Menu_Flag = 312 Then
                           Msg = " Title " + Chr(34) + "RxQualFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""IDLE  [" & MyPercent & "]"" display on ,"
                        ElseIf Menu_Flag = 88315 Then
                           Msg = " Title " + Chr(34) + "第二手机RxQualSub观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""IDLE  [" & MyPercent & "]"" display on ,"
                        Else
                           Msg = " Title " + Chr(34) + "第二手机RxQualFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""IDLE  [" & MyPercent & "]"" display on ,"
                        End If
                     Else
                        If Menu_Flag = 315 Then
                           Msg = " Title " + Chr(34) + "RxQualSub观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,"
                        ElseIf Menu_Flag = 312 Then
                           Msg = " Title " + Chr(34) + "RxQualFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,"
                        ElseIf Menu_Flag = 88315 Then
                           Msg = " Title " + Chr(34) + "第二手机RxQualSub观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,"
                        Else
                           Msg = " Title " + Chr(34) + "第二手机RxQualFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,"
                        End If
                     End If
                     mapinfo.do "select * from " & tblname & " where " & MyField & " = " & QMark & "0" & QMark & " into Mytemp"
                     MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                     MyPercent = Format(MyPoint / AllRows, "Percent")
                     Msg = Msg + """0  [" & MyPercent & "]"" display on ,"
                     mapinfo.do "select * from " & tblname & " where " & MyField & " = " & QMark & "1" & QMark & " into Mytemp"
                     MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                     MyPercent = Format(MyPoint / AllRows, "Percent")
                     Msg = Msg + """1  [" & MyPercent & "]"" display on ,"
                     mapinfo.do "select * from " & tblname & " where " & MyField & " = " & QMark & "2" & QMark & " into Mytemp"
                     MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                     MyPercent = Format(MyPoint / AllRows, "Percent")
                     Msg = Msg + """2  [" & MyPercent & "]"" display on ,"
                     mapinfo.do "select * from " & tblname & " where " & MyField & " = " & QMark & "3" & QMark & " into Mytemp"
                     MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                     MyPercent = Format(MyPoint / AllRows, "Percent")
                     Msg = Msg + """3  [" & MyPercent & "]"" display on ,"
                     mapinfo.do "select * from " & tblname & " where " & MyField & " = " & QMark & "4" & QMark & " into Mytemp"
                     MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                     MyPercent = Format(MyPoint / AllRows, "Percent")
                     Msg = Msg + """4  [" & MyPercent & "]"" display on ,"
                     mapinfo.do "select * from " & tblname & " where " & MyField & " = " & QMark & "5" & QMark & " into Mytemp"
                     MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                     MyPercent = Format(MyPoint / AllRows, "Percent")
                     Msg = Msg + """5  [" & MyPercent & "]"" display on ,"
                     mapinfo.do "select * from " & tblname & " where " & MyField & " = " & QMark & "6" & QMark & " into Mytemp"
                     MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                     MyPercent = Format(MyPoint / AllRows, "Percent")
                     Msg = Msg + """6  [" & MyPercent & "]"" display on ,"
                     mapinfo.do "select * from " & tblname & " where " & MyField & " = " & QMark & "7" & QMark & " into Mytemp"
                     MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                     MyPercent = Format(MyPoint / AllRows, "Percent")
                     Msg = Msg + """7  [" & MyPercent & "]"" display on "
                     If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") <> 1 Then
                        mapinfo.do "select * from " & tblname & " where " & MyField & " = " & QMark & "9" & QMark & " into Mytemp"
                        MyPoint = mapinfo.eval("tableinfo(mytemp,8)")
                        If MyPoint > 0 Then
                           MyPercent = Format(MyPoint / AllRows, "Percent")
                           Msg = Msg + ",""IDLE  [" & MyPercent & "]"" display on "
                        End If
                     End If
                     mapinfo.do "close table mytemp"
                     mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on " & Msg
                  Else
                     If Menu_Flag = 315 Then
                        Msg = " Title " + Chr(34) + "RxQualSub观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off "
                     ElseIf Menu_Flag = 312 Then
                        Msg = " Title " + Chr(34) + "RxQualFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off "
                     ElseIf Menu_Flag = 88315 Then
                        Msg = " Title " + Chr(34) + "第二手机RxQualSub观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off "
                     Else
                        Msg = " Title " + Chr(34) + "第二手机RxQualFull观测 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,255) ascending off ranges Font (""宋体"",0,9,0) """" display off "
                     End If
                     mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
                  End If
                  If Menu_Flag = 315 Or Menu_Flag = 312 Then
                        mapinfo.do "set map redraw off"
                        mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                        mapinfo.do "set map redraw on"
                  End If
                  'msg = " Title " + Chr(34) + "RxQualSub观测  " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  'mapinfo.do "set legend window FrontWindow()  Layer prev " & msg
          Case 316
               mapinfo.do "select max(bcch_serv) from " & tblname & " into mytemp"
               If Val(mapinfo.eval("mytemp.col1")) > 125 Then
               
                   Msg = " shade window FrontWindow() " + tblname + " With val(Tx_Power) "        ' ignore 0  values 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9 ,10 ,11 ,12 ,13 ,14 ,15 ,16 ,17 ,18 ,19 ,20 ,21 ,22 ,23 ,24 ,25 ,26 ,27 ,28 ,29 ,30  default Symbol (34,0,12) "
                   'msg = msg + "ignore 0 values  1 Symbol (63,16711680,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "ignore 0 values "
                   Msg = Msg + "29 Symbol (63,65280,8,""MapInfo Cartographic"",0,0) ,30 Symbol (63,255,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "31 Symbol (63,16711935,8,""MapInfo Cartographic"",0,0) ,0 Symbol (63,16776960,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "1 Symbol (63,65535,8,""MapInfo Cartographic"",0,0) ,2 Symbol (63,8388608,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "3 Symbol (63,32768,8,""MapInfo Cartographic"",0,0),4 Symbol (63,128,8,""MapInfo Cartographic"",0,0),5 Symbol (63,8388736,8,""MapInfo Cartographic"",0,0),6 Symbol (63,8421376,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "7 Symbol (63,32896,8,""MapInfo Cartographic"",0,0),8 Symbol (63,16744576,8,""MapInfo Cartographic"",0,0),9 Symbol (63,8454016,8,""MapInfo Cartographic"",0,0),10 Symbol (63,8421631,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "11 Symbol (63,16744703,8,""MapInfo Cartographic"",0,0),12 Symbol (63,16777088,8,""MapInfo Cartographic"",0,0),13 Symbol (63,8454143,8,""MapInfo Cartographic"",0,0),14 Symbol (63,8405056,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "15 Symbol (63,4227136,8,""MapInfo Cartographic"",0,0),16 Symbol (63,4210816,8,""MapInfo Cartographic"",0,0),17 Symbol (63,8405120,8,""MapInfo Cartographic"",0,0),18 Symbol (63,8421440,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "19 Symbol (63,4227200,8,""MapInfo Cartographic"",0,0),20 Symbol (63,16761024,8,""MapInfo Cartographic"",0,0),21 Symbol (63,12648384,8,""MapInfo Cartographic"",0,0),22 Symbol (63,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "23 Symbol (63,16761087,8,""MapInfo Cartographic"",0,0),24 Symbol (63,16777152,8,""MapInfo Cartographic"",0,0),25 Symbol (63,12648447,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "26 Symbol (63,26761087,8,""MapInfo Cartographic"",0,0),27 Symbol (63,26777152,8,""MapInfo Cartographic"",0,0),28 Symbol (63,22648447,8,""MapInfo Cartographic"",0,0) "
                   Msg = Msg + " default Symbol (63,16777215,8,""MapInfo Cartographic"",0,0)"
                   mapinfo.do Msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  'msg = " Title " + Chr(34) + "Tx Power 观测  " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                   
                   Msg = " Title " + Chr(34) + "Tx Power 观测  " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""29 (36dBm)"" display on ,""30 (34dBm)"" display on ,""31 (32dBm)"" display on ,""0  (30dBm)"" display on ,"
                   Msg = Msg & """1  (28dBm)"" display on ,""2  (26dBm)"" display on ,""3  (24dBm)"" display on ,""4  (22dBm)"" display on ,"
                   Msg = Msg & """5  (20dBm)"" display on ,""6  (18dBm)"" display on ,""7 (16dBm)"" display on ,""8 (14dBm)"" display on ,"
                   Msg = Msg & """9  (12dBm)"" display on ,""10 (10dBm)"" display on ,""11 (8dBm)"" display on ,""12 (6dBm)"" display on ,"
                   Msg = Msg & """13 (4dBm)"" display on ,""14 (2dBm)"" display on ,""15 (0dBm)"" display on ,""16 (0dBm)"" display on ,"
                   Msg = Msg & """17 (0dBm)"" display on ,""18 (0dBm)"" display on ,""19 (0dBm)"" display on ,""20 (0dBm)"" display on ,"
                   Msg = Msg & """21 (0dBm)"" display on ,""22 (0dBm)"" display on ,""23 (0dBm)"" display on ,""24 (0dBm)"" display on ,""25 (0dBm)"" display on '"
                   Msg = Msg & """26 (0dBm)"" display on ,""27 (0dBm)"" display on ,""28 (0dBm)"" display on "
               
               Else
                   Msg = " shade window FrontWindow() " + tblname + " With val(Tx_Power) "        ' ignore 0  values 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9 ,10 ,11 ,12 ,13 ,14 ,15 ,16 ,17 ,18 ,19 ,20 ,21 ,22 ,23 ,24 ,25 ,26 ,27 ,28 ,29 ,30  default Symbol (34,0,12) "
                   'msg = msg + "ignore 0 values  1 Symbol (63,16711680,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "ignore 0 values "
                   Msg = Msg + "2 Symbol (63,65280,8,""MapInfo Cartographic"",0,0) ,3 Symbol (63,255,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "4 Symbol (63,16711935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (63,16776960,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "6 Symbol (63,65535,8,""MapInfo Cartographic"",0,0) ,7 Symbol (63,8388608,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "8 Symbol (63,32768,8,""MapInfo Cartographic"",0,0),9 Symbol (63,128,8,""MapInfo Cartographic"",0,0),10 Symbol (63,8388736,8,""MapInfo Cartographic"",0,0),11 Symbol (63,8421376,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "12 Symbol (63,32896,8,""MapInfo Cartographic"",0,0),13 Symbol (63,16744576,8,""MapInfo Cartographic"",0,0),14 Symbol (63,8454016,8,""MapInfo Cartographic"",0,0),15 Symbol (63,8421631,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "16 Symbol (63,16744703,8,""MapInfo Cartographic"",0,0),17 Symbol (63,16777088,8,""MapInfo Cartographic"",0,0),18 Symbol (63,8454143,8,""MapInfo Cartographic"",0,0),19 Symbol (63,8405056,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "20 Symbol (63,4227136,8,""MapInfo Cartographic"",0,0),21 Symbol (63,4210816,8,""MapInfo Cartographic"",0,0),22 Symbol (63,8405120,8,""MapInfo Cartographic"",0,0),23 Symbol (63,8421440,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "24 Symbol (63,4227200,8,""MapInfo Cartographic"",0,0),25 Symbol (63,16761024,8,""MapInfo Cartographic"",0,0),26 Symbol (63,12648384,8,""MapInfo Cartographic"",0,0),27 Symbol (63,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "28 Symbol (63,16761087,8,""MapInfo Cartographic"",0,0),29 Symbol (63,16777152,8,""MapInfo Cartographic"",0,0),30 Symbol (63,12648447,8,""MapInfo Cartographic"",0,0) "
                   Msg = Msg + " default Symbol (63,16777215,8,""MapInfo Cartographic"",0,0)"
                   mapinfo.do Msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  'msg = " Title " + Chr(34) + "Tx Power 观测  " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                   
                   Msg = " Title " + Chr(34) + "Tx Power 观测 (DCS)" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""2  (39dBm)"" display on ,""3  (37dBm)"" display on ,""4  (35dBm)"" display on ,""5  (33dBm)"" display on ,"
                   Msg = Msg & """6  (31dBm)"" display on ,""7  (29dBm)"" display on ,""8  (27dBm)"" display on ,""9  (25dBm)"" display on ,"
                   Msg = Msg & """10 (23dBm)"" display on ,""11 (21dBm)"" display on ,""12 (19dBm)"" display on ,""13 (17dBm)"" display on ,"
                   Msg = Msg & """14 (15dBm)"" display on ,""15 (13dBm)"" display on ,""16 (11dBm)"" display on ,""17 (9dBm)"" display on ,"
                   Msg = Msg & """18 (7dBm)"" display on ,""19 (5dBm)"" display on ,""20 (5dBm)"" display on ,""21 (5dBm)"" display on ,"
                   Msg = Msg & """22 (5dBm)"" display on ,""23 (5dBm)"" display on ,""24 (5dBm)"" display on ,""25 (5dBm)"" display on ,"
                   Msg = Msg & """26 (5dBm)"" display on ,""27 (5dBm)"" display on ,""28 (5dBm)"" display on ,""29 (5dBm)"" display on ,""30 (5dBm)"" display on "
               
               End If
                   mapinfo.do "close table mytemp"
                mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
                mapinfo.do "set map redraw off"
                mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                mapinfo.do "set map redraw on"

          'Case 317, 88317
          Case 88317
               If Menu_Flag = 317 Then
                  Msg = " shade window FrontWindow() " + tblname + " With ta  "        ' ignore 0  values 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9 ,10 ,11 ,12 ,13 ,14 ,15 ,16 ,17 ,18 ,19 ,20 ,21 ,22 ,23 ,24 ,25 ,26 ,27 ,28 ,29 ,30  default Symbol (34,0,12) "
               Else
                  Msg = " shade window FrontWindow() " + tblname + " With ta_2 "        ' ignore 0  values 1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9 ,10 ,11 ,12 ,13 ,14 ,15 ,16 ,17 ,18 ,19 ,20 ,21 ,22 ,23 ,24 ,25 ,26 ,27 ,28 ,29 ,30  default Symbol (34,0,12) "
               End If
'白色  16777215
                   'msg = msg + "values """" Symbol (42,0,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "values """" Symbol (63,14737632,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "0 Symbol (63,65280,8,""MapInfo Cartographic"",0,0) , 1 Symbol (63,7585792,8,""MapInfo Cartographic"",0,0) ,"
                   'msg = msg + "values 0 Symbol (63,15257855,8,""MapInfo Cartographic"",0,0) , 1 Symbol (63,16711680,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "2 Symbol (63,8421376,8,""MapInfo Cartographic"",0,0) ,3 Symbol (63,8388736,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "4 Symbol (63,255,8,""MapInfo Cartographic"",0,0) ,5 Symbol (63,8432639,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "6 Symbol (63,65535,8,""MapInfo Cartographic"",0,0) ,7 Symbol (63,16750640,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "8 Symbol (63,16765088,8,""MapInfo Cartographic"",0,0),9 Symbol (63,16711935,8,""MapInfo Cartographic"",0,0),10 Symbol (63,16756952,8,""MapInfo Cartographic"",0,0),11 Symbol (63,8421376,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "12 Symbol (63,32896,8,""MapInfo Cartographic"",0,0),13 Symbol (63,16744576,8,""MapInfo Cartographic"",0,0),14 Symbol (63,8454016,8,""MapInfo Cartographic"",0,0),15 Symbol (63,8421631,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "16 Symbol (63,16744703,8,""MapInfo Cartographic"",0,0),17 Symbol (63,16777088,8,""MapInfo Cartographic"",0,0),18 Symbol (63,8454143,8,""MapInfo Cartographic"",0,0),19 Symbol (63,8405056,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "20 Symbol (63,4227136,8,""MapInfo Cartographic"",0,0),21 Symbol (63,4210816,8,""MapInfo Cartographic"",0,0),22 Symbol (63,8405120,8,""MapInfo Cartographic"",0,0),23 Symbol (63,8421440,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "24 Symbol (63,4227200,8,""MapInfo Cartographic"",0,0),25 Symbol (63,16761024,8,""MapInfo Cartographic"",0,0),26 Symbol (63,12648384,8,""MapInfo Cartographic"",0,0),27 Symbol (63,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "28 Symbol (63,16761087,8,""MapInfo Cartographic"",0,0),29 Symbol (63,16777152,8,""MapInfo Cartographic"",0,0),30 Symbol (63,12648447,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "31 Symbol (63,8413280,8,""MapInfo Cartographic"",0,0),32 Symbol (63,6324320,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "33 Symbol (63,6316160,8,""MapInfo Cartographic"",0,0),34 Symbol (63,8413312,8,""MapInfo Cartographic"",0,0),35 Symbol (63,8421472,8,""MapInfo Cartographic"",0,0),36 Symbol (63,6324352,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "37 Symbol (63,16711680,8,""MapInfo Cartographic"",0,0),38 Symbol (63,65280,8,""MapInfo Cartographic"",0,0),39 Symbol (63,255,8,""MapInfo Cartographic"",0,0),40 Symbol (63,16711935,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "41 Symbol (63,16776960,8,""MapInfo Cartographic"",0,0),42 Symbol (63,65535,8,""MapInfo Cartographic"",0,0),43 Symbol (63,8388608,8,""MapInfo Cartographic"",0,0),44 Symbol (63,32768,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "45 Symbol (63,128,8,""MapInfo Cartographic"",0,0),46 Symbol (63,8388736,8,""MapInfo Cartographic"",0,0),47 Symbol (63,8421376,8,""MapInfo Cartographic"",0,0),48 Symbol (63,32896,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "49 Symbol (63,16744576,8,""MapInfo Cartographic"",0,0),50 Symbol (63,8454016,8,""MapInfo Cartographic"",0,0),51 Symbol (63,8421631,8,""MapInfo Cartographic"",0,0),52 Symbol (63,16744703,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "53 Symbol (63,16777088,8,""MapInfo Cartographic"",0,0),54 Symbol (63,8454143,8,""MapInfo Cartographic"",0,0),55 Symbol (63,8405056,8,""MapInfo Cartographic"",0,0),56 Symbol (63,4227136,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "57 Symbol (63,4210816,8,""MapInfo Cartographic"",0,0),58 Symbol (63,8405120,8,""MapInfo Cartographic"",0,0),59 Symbol (63,8421440,8,""MapInfo Cartographic"",0,0),60 Symbol (63,4227200,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "61 Symbol (63,16761024,8,""MapInfo Cartographic"",0,0),62 Symbol (63,12648384,8,""MapInfo Cartographic"",0,0),63 Symbol (63,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "default Symbol (63,16777215,8,""MapInfo Cartographic"",0,0)"
                   
                   mapinfo.do Msg
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  If Menu_Flag = 317 Then
                     Msg = " Title " + Chr(34) + "覆盖合理性统计 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) """" display off ,""IDLE"" display on"
                  Else
                     Msg = " Title " + Chr(34) + "第二手机 Timing Advance 观测  " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) """" display off,""N/A"" display on"
                  End If
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
                  If Menu_Flag = 317 Then
                        mapinfo.do "set map redraw off"
                        mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                        mapinfo.do "set map redraw on"
                  End If
          Case 318
               Screen.MousePointer = 11
               row = Val(mapinfo.eval("tableinfo(" & tblname & ",8)"))
               mapinfo.do "fetch first from " & tblname
               If UCase(Right(tblname, 1)) = "F" Or UCase(Right(tblname, 1)) = "E" Then
                  Is_Smooth = False
               Else
                  Is_Smooth = True
               End If
               j = 1
               For i = 1 To row
                   If Val(mapinfo.eval(tblname & ".ci_serv")) > 0 Then
                      If j = 1 Then
                         Ci_Serv(j) = mapinfo.eval(tblname & ".ci_serv")
                         j = j + 1
                      Else
                         If Ci_Serv(j - 1) <> mapinfo.eval(tblname & ".ci_serv") Then
                            For k = 1 To j - 1
                                If Ci_Serv(k) = mapinfo.eval(tblname & ".ci_serv") Then
                                   GoTo Next_Point
                                End If
                            Next
                            Ci_Serv(j) = mapinfo.eval(tblname & ".ci_serv")
                            j = j + 1
                         End If
                      End If
                   End If
Next_Point:
                   If Is_Smooth = True Then
                      mapinfo.do "fetch rec " & (i + 50) & " from " & tblname
                      i = i + 50
                   Else
                      mapinfo.do "fetch next from " & tblname
                   End If
               Next
               Screen.MousePointer = 0
               Ci_No = j - 1
               For i = 1 To Ci_No
                   SortTemp = Ci_Serv(1)
                   TempNum = 1
                   For j = 1 To Ci_No
                       If SortTemp < Ci_Serv(j) Then
                          SortTemp = Ci_Serv(j)
                          TempNum = j
                       End If
                   Next
                   Sort_Ci_Serv(Ci_No + 1 - i) = SortTemp
                   Ci_Serv(TempNum) = ""
               Next
               my_msg = "shade window FrontWindow() " + tblname + " With ci_serv values "
               My_Color = "10535167"
               For i = 1 To Ci_No
                   If i = Ci_No Then
                      my_msg = my_msg + Chr(34) + Sort_Ci_Serv(i) + Chr(34) + " Symbol (83," + My_Color + ",10,""Wingdings"",0,0) "
                      my_msg = my_msg + "default Symbol(83,0,10,""Wingdings"",0,0)"
                   Else
                      my_msg = my_msg + Chr(34) + Sort_Ci_Serv(i) + Chr(34) + " Symbol (83," + My_Color + ",10,""Wingdings"",0,0),"
                   End If
                   My_Color = Format(Val(My_Color) + 6000)
                   Ci_Name(i) = Findcell(Sort_Ci_Serv(i))
                   finds = InStr(Ci_Name(i), Chr(0))
                   If finds > 0 Then
                      Ci_Name(i) = Trim(Left(Ci_Name(i), finds - 1))
                   End If
                   If Ci_Name(i) = "" Then
                      Ci_Name(i) = "    "
                   End If
               Next
               mapinfo.do my_msg
               my_msg = ""
               For i = 1 To Ci_No
                   'my_msg = my_msg + "," + Chr(34) + Ci_Name(i) + " [" + Ci_Serv(i) + "]" + Chr(34) + "display on"
                   my_msg = my_msg + "," + Chr(34) + Sort_Ci_Serv(i) + " [" + Ci_Name(i) + "]" + Chr(34) + "display on"
               Next
               'Change Legend
               my_msg = " Title " + Chr(34) + "服务小区分布观察 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + "Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""无服务小区"" display on" + my_msg
               mapinfo.do "set legend window FrontWindow() Layer prev " & my_msg
                mapinfo.do "set map redraw off"
                mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                mapinfo.do "set map redraw on"

          Case 3211, 3221, 3231, 3241, 3251, 3261
                 Select Case Menu_Flag
                     Case 3211
                          Msg = " shade window FrontWindow() " + tblname + " With Rxlev_n1 "
                     Case 3221
                          Msg = " shade window FrontWindow() " + tblname + " With Rxlev_n2 "
                     Case 3231
                          Msg = " shade window FrontWindow() " + tblname + " With Rxlev_n3 "
                     Case 3241
                          Msg = " shade window FrontWindow() " + tblname + " With Rxlev_n4 "
                     Case 3251
                          Msg = " shade window FrontWindow() " + tblname + " With Rxlev_n5 "
                     Case 3261
                          Msg = " shade window FrontWindow() " + tblname + " With Rxlev_n6 "
                  End Select
                  If Legend_Tog = 0 Then
                       'msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,17: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 35 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                  Else
                       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                  End If
                  mapinfo.do Msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If

                  If Legend_Tog = 0 Then
                         'msg = " Title " + Chr(34) + "相邻小区RxlevFull观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on ,""17 至 27 (-93至-83dBm)"" display on ,""27 至 63 (-83至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                         Msg = " Title " + Chr(34) + "相邻小区RxlevFull观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                  Else
                         Msg = " Title " + Chr(34) + "相邻小区RxlevFull观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：RXLEV" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 15 (-100至-95dBm)"" display on ,""15 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                  End If
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg

          Case 3212, 3222, 3232, 3242, 3252, 3262
                 Select Case Menu_Flag
                     Case 3212
                          Msg = " shade window FrontWindow() " + tblname + " With Bcch_N1 ignore 0  values  "
                     Case 3222
                          Msg = " shade window FrontWindow() " + tblname + " With Bcch_N2 ignore 0  values  "
                     Case 3232
                          Msg = " shade window FrontWindow() " + tblname + " With Bcch_N3 ignore 0  values  "
                     Case 3242
                          Msg = " shade window FrontWindow() " + tblname + " With Bcch_N4 ignore 0  values  "
                     Case 3252
                          Msg = " shade window FrontWindow() " + tblname + " With Bcch_N5 ignore 0  values  "
                     Case 3262
                          Msg = " shade window FrontWindow() " + tblname + " With Bcch_N6 ignore 0  values  "
                  End Select
'                  Msg = Msg + "1 ,2 ,3 ,4 ,5 ,6 ,7 ,8 ,9 ,10 ,11 ,12 ,13 ,14 ,15 ,16 ,17 ,18 ,19 ,20 ,21 ,22 ,23 ,24 ,25 ,26 ,27 ,28 ,29 ,30 ,31 ,32 ,33 ,34 ,35 ,36 ,37 ,38 ,39 ,40 ,41 ,42 ,43 ,44 ,45 ,46 ,47 ,48 ,49 ,50 ,51 ,52 ,53 ,54 ,55 ,56 ,57 ,58 ,59 ,60 ,61 ,62 ,63 ,64 ,65 ,66 ,67 ,68 ,69 ,70 ,71 ,72 ,73 ,74 ,75 ,76 ,77 ,78 ,79 ,80 ,81 ,82 ,83 ,84 ,85 ,86 ,87 ,88 ,89 ,90 ,91 ,92 ,93 ,94 ,95 ,96 ,97 ,98 ,99 ,100 ,101 ,102 ,103 ,104 ,105 ,106 ,107 ,108 ,109 ,110 ,111 ,112 ,113 ,114 ,115 ,116 ,117 ,118 ,119 ,120 ,121 ,122 ,123 ,124 default Symbol (34,0,12) "

                   Msg = Msg + " 1 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "2 Symbol (33,65280,8,""MapInfo Cartographic"",0,0) ,3 Symbol (33,255,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "4 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0) ,5 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "6 Symbol (33,65535,8,""MapInfo Cartographic"",0,0) ,7 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "8 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),9 Symbol (33,128,8,""MapInfo Cartographic"",0,0),10 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),11 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "12 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),13 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),14 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),15 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "16 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),17 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),18 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),19 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "20 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),21 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),22 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),23 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "24 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),25 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),26 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),27 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "28 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),29 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),30 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "31 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),32 Symbol (33,6324320,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "33 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),34 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),35 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),36 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "37 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),38 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),39 Symbol (33,255,8,""MapInfo Cartographic"",0,0),40 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "41 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),42 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),43 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),44 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "45 Symbol (33,128,8,""MapInfo Cartographic"",0,0),46 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),47 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),48 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "49 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),50 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),51 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),52 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "53 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),54 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),55 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),56 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "57 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),58 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),59 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),60 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "61 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),62 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),63 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "64 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),65 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),66 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "67 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),68 Symbol (33"
                   Msg = Msg + ",6324320,8,""MapInfo Cartographic"",0,0),69 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "70 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),71 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),72 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),73 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "74 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),75 Symbol (33,255,8,""MapInfo Cartographic"",0,0),76 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),77 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "78 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),79 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),80 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),81 Symbol (33,128,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "82 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),83 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),84 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),85 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "86 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),87 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),88 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0),89 Symbol (33,16777088,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "90 Symbol (33,8454143,8,""MapInfo Cartographic"",0,0),91 Symbol (33,8405056,8,""MapInfo Cartographic"",0,0),92 Symbol (33,4227136,8,""MapInfo Cartographic"",0,0),93 Symbol (33,4210816,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "94 Symbol (33,8405120,8,""MapInfo Cartographic"",0,0),95 Symbol (33,8421440,8,""MapInfo Cartographic"",0,0),96 Symbol (33,4227200,8,""MapInfo Cartographic"",0,0),97 Symbol (33,16761024,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "98 Symbol (33,12648384,8,""MapInfo Cartographic"",0,0),99 Symbol (33,12632319,8,""MapInfo Cartographic"",0,0),100 Symbol (33,16761087,8,""MapInfo Cartographic"",0,0),101 Symbol (33,16777152,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "102 Symbol (33,12648447,8,""MapInfo Cartographic"",0,0),103 Symbol (33,8413280,8,""MapInfo Cartographic"",0,0),104 Symbol (33,6324320,8,""MapInfo Cartographic"",0,0),105 Symbol (33,6316160,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "106 Symbol (33,8413312,8,""MapInfo Cartographic"",0,0),107 Symbol (33,8421472,8,""MapInfo Cartographic"",0,0),108 Symbol (33,6324352,8,""MapInfo Cartographic"",0,0),109 Symbol (33,16711680,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "110 Symbol (33,65280,8,""MapInfo Cartographic"",0,0),111 Symbol (33,255,8,""MapInfo Cartographic"",0,0),112 Symbol (33,16711935,8,""MapInfo Cartographic"",0,0),113 Symbol (33,16776960,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "114 Symbol (33,65535,8,""MapInfo Cartographic"",0,0),115 Symbol (33,8388608,8,""MapInfo Cartographic"",0,0),116 Symbol (33,32768,8,""MapInfo Cartographic"",0,0),117 Symbol (33,128,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "118 Symbol (33,8388736,8,""MapInfo Cartographic"",0,0),119 Symbol (33,8421376,8,""MapInfo Cartographic"",0,0),120 Symbol (33,32896,8,""MapInfo Cartographic"",0,0),121 Symbol (33,16744576,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "122 Symbol (33,8454016,8,""MapInfo Cartographic"",0,0),123 Symbol (33,8421631,8,""MapInfo Cartographic"",0,0),124 Symbol (33,16744703,8,""MapInfo Cartographic"",0,0)"
                   mapinfo.do Msg
 
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  Msg = " Title " + Chr(34) + "相邻小区BCCH_ARFCN观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
          
          Case 330
                  L3_Sel.Show 1
          Case 1657
                  frmMark.Show 1

          Case 337
                  mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + UCase(Trim(Msg_3_Layer)) + Chr(34) + " into Result"
                  mapinfo.do "Add Map window FrontWindow() Layer  Result"

                  mapinfo.do "shade window FrontWindow() Result with MESSAGE values  " + Chr(34) + Msg_3_Layer + Chr(34) + " Symbol (""lay3.bmp"",255,22,0) "
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  Msg = " Title " + Chr(34) + "Other 3 Layer Message 观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg

          Case 34

          Case 35
                  mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 2"
                  mapinfo.do "browse * from  " & tblname
                  mapinfo.do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
          Case 41
                  OverLayFrm.Show 1
          'Case 42, 441, 442
          Case 42
                  SelCond.Show 1
          Case 441, 442
                  FrmDisturb.Show 1
          Case 431
                  Ta_Qual.Show 1
          Case 45
                  SelCond.Show 1
          Case 81
                   mapinfo.do "Set Next Document Parent " & MapForm.hWnd & " Style 1"
                   mapinfo.do "Map From " & tblname
          Case 912, 913
                   ScanSel.Show 1
          Case 914, 915, 916, 917
                  SelCond.Show 1
          Case 918
                 Screen.MousePointer = 11
                 Msg = " shade window FrontWindow() " + tblname + " With c_i_1  ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  0: -110 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0),0: 1 Symbol (39,16711825,8,""MapInfo Cartographic"",0,0) ,1: 2 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) ,2: 3 Symbol (39,16777072,8,""MapInfo Cartographic"",0,0) ,3: 4 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,4: 5 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,5: 6 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,6: 7 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,7: 8 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,8: 9 Symbol (39,0,8,""MapInfo Cartographic"",0,0) ,9: 10 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,10: 110 Symbol (39,10551200,8,""MapInfo Cartographic"",0,0) "
                 mapinfo.do Msg

                 If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                 End If
                 Msg = " Title " + Chr(34) + "载干比1 观测  " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：-dBm" + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 以下"" display on ,""0  "" display on ,""1 "" display on ,""2 "" display on ,""3 "" display on ,""4 "" display on ,""5 "" display on ,""6 "" display on ,""7 "" display on ,""8 "" display on ,""9 "" display on ,""9 以上 "" display on"
                 mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg

                 Screen.MousePointer = 0
          Case 919
                 Screen.MousePointer = 11
                  On Error Resume Next
                 Msg = " shade window FrontWindow() " + tblname + " With c_i_2  ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  0: -110 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0),0: 1 Symbol (39,16711825,8,""MapInfo Cartographic"",0,0) ,1: 2 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) ,2: 3 Symbol (39,16777072,8,""MapInfo Cartographic"",0,0) ,3: 4 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,4: 5 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,5: 6 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,6: 7 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,7: 8 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,8: 9 Symbol (39,0,8,""MapInfo Cartographic"",0,0) ,9: 10 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,10: 110 Symbol (39,10551200,8,""MapInfo Cartographic"",0,0) "
                 mapinfo.do Msg

                 If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                 End If
                 Msg = " Title " + Chr(34) + "载干比2 观测" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：-dBm" + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off ,""0 以下"" display on ,""0  "" display on ,""1 "" display on ,""2 "" display on ,""3 "" display on ,""4 "" display on ,""5 "" display on ,""6 "" display on ,""7 "" display on ,""8 "" display on ,""9 "" display on ,""9 以上 "" display on"
                 mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg

                 Screen.MousePointer = 0

          Case 5004
                 Screen.MousePointer = 11
                 On Error Resume Next
                 Dim ncell_rxlev(6), ncell_bsic(6), ncell_bcch(6), tmp_buf(3), kk, jj, pp As Integer
                 i = 1
                 row = Val(mapinfo.eval("tableinfo(" & tblname & ",8)"))
        '         On Error GoTo 0
                 mapinfo.do "fetch First from " & tblname
                 While i <= row
                     For kk = 1 To 6
                         Msg = tblname + ".rxlev_n" & kk
                         ncell_rxlev(kk) = Val(mapinfo.eval(Msg))

                         Msg = tblname + ".bsic_n" & kk
                         ncell_bsic(kk) = Val(mapinfo.eval(Msg))

                         Msg = tblname + ".bcch_n" & kk
                         ncell_bcch(kk) = Val(mapinfo.eval(Msg))
                     Next kk
                     For kk = 1 To 6
                         tmp_buf(1) = ncell_rxlev(kk)
                         tmp_buf(2) = ncell_bsic(kk)
                         tmp_buf(3) = ncell_bcch(kk)
                         For jj = kk + 1 To 6
                             If tmp_buf(1) < ncell_rxlev(jj) Then pp = jj
                         Next jj
                         If pp <> kk Then
                            ncell_rxlev(kk) = ncell_rxlev(pp)
                            ncell_bsic(kk) = ncell_bsic(pp)
                            ncell_bcch(kk) = ncell_bcch(pp)

                            ncell_rxlev(kk) = tmp_buf(1)
                            ncell_bsic(kk) = tmp_buf(2)
                            ncell_bcch(kk) = tmp_buf(3)
                         End If
                     Next kk

                     Msg = "update  " & tblname & ""
                     For jj = 1 To 6
                       Msg = Msg + " set bcch_n" & jj & "  = " + Chr(34) + str$(ncell_bcch(j)) + Chr(34) + " ,rxlev_n" & jj & "  = " + Chr(34) + str$(ncell_rxlev(j)) + Chr(34) + " ,bsic_n" & jj & "  = " + Chr(34) + str$(ncell_bsic(j)) + Chr(34)
                     Next jj
                     Msg = Msg + "  Where rowid=" & i
'                     MsgBox Msg
                     mapinfo.do Msg
                     mapinfo.do "fetch next from " & tblname
                     i = i + 1
                 Wend
                 Screen.MousePointer = 0
                 mapinfo.do "commit table " & tblname
          Case 7001
                'Call My_Report
          
          Case 9901
                My_ArfcnChanging.Show 1
          Case 888, 885
               Cope_RxLev.Show 1
          Case 887, 884
               Cope_RxQual.Show 1
          Case 886, 883
               Cope_Global.Show 1
          Case 6451, 6452
                  If Menu_Flag = 6451 Then
                     mapinfo.do "select * from " & tblname & " where rxlev_f < 17 into NetWorkBlind"
                  Else
                     mapinfo.do "select * from " & tblname & " where rxlev_s < 17 into NetWorkBlind"
                  End If
                  If Val(mapinfo.eval("tableinfo(NetWorkBlind,8)")) = 0 Then
                      MsgBox "该路段不存在网络覆盖盲区", 64, "提示"
                      mapinfo.do "close talbe NetWorkBlind"
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
                  mapinfo.do "Add Map window " & WinId & " Layer NetWorkBlind"
                  If Menu_Flag = 6451 Then
                     Msg = " shade window " & WinId & " NetWorkBlind With RXLEV_F "
                  Else
                     Msg = " shade window " & WinId & " NetWorkBlind With RXLEV_s "
                  End If
                  If Legend_Tog = 0 Then
                       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 17: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                  Else
                       Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 17: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                  End If
                  mapinfo.do Msg
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window " & WinId
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  If Legend_Tog = 0 Then
                     'msg = " Title " + Chr(34) + "网络覆盖盲区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on "
                     If Menu_Flag = 6451 Then
                         Msg = " Title " + Chr(34) + "网络覆盖盲区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：Rxlev_f<17(RXLEV)" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on "
                     Else
                         Msg = " Title " + Chr(34) + "网络覆盖盲区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：Rxlev_s<17(RXLEV)" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 17 (-110至-93dBm)"" display on "
                     End If
                  Else
                     If Menu_Flag = 6451 Then
                         Msg = " Title " + Chr(34) + "网络覆盖盲区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：Rxlev_f<17(RXLEV)" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 17 (-100至-93dBm)"" display on "
                     Else
                         Msg = " Title " + Chr(34) + "网络覆盖盲区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：Rxlev_s<17(RXLEV)" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 17 (-100至-93dBm)"" display on "
                     End If
                     'msg = " Title " + Chr(34) + "网络覆盖盲区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 5 (-110至-105dBm)"" display on ,""5 至 10 (-105至-100dBm)"" display on ,""10 至 17 (-100至-93dBm)"" display on "
                  End If
                  mapinfo.do "set legend window " & WinId & " Layer prev display on shades off symbols on lines off count on " & Msg
          Case 7451, 7452
                  If Menu_Flag = 7451 Then
                     If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                        mapinfo.do "select * from " & tblname & " where rxlev_f > 17 and val(rxqual_f) > 3 into NetWorkDisturb"
                     Else
                        mapinfo.do "select * from " & tblname & " where rxlev_f > 17 and rxqual_f > 3 into NetWorkDisturb"
                     End If
                  Else
                     If mapinfo.eval("ColumnInfo(" & tblname & ", ""rxqual_s"", 3)") = 1 Then
                        mapinfo.do "select * from " & tblname & " where rxlev_s > 17 and val(rxqual_s) > 3 into NetWorkDisturb"
                     Else
                        mapinfo.do "select * from " & tblname & " where rxlev_s > 17 and rxqual_s > 3 into NetWorkDisturb"
                     End If
                  End If
                  If Val(mapinfo.eval("tableinfo(NetWorkDisturb,8)")) = 0 Then
                      MsgBox "该路段不存在网络干扰区", 64, "提示"
                      mapinfo.do "close talbe NetWorkDisturb"
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
                  If Menu_Flag = 7451 Then
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
                      If Menu_Flag = 7451 Then
                           Msg = " Title " + Chr(34) + "网络干扰区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：RxLev_f>17(RXLEV)且RxQual_f>3" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                      Else
                           Msg = " Title " + Chr(34) + "网络干扰区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：RxLev_s>17(RXLEV)且RxQual_s>3" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                      End If
                        'msg = " Title " + Chr(34) + "网络干扰区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
                  Else
                      If Menu_Flag = 7451 Then
                           Msg = " Title " + Chr(34) + "网络干扰区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：RxLev_f>17(RXLEV)且RxQual_f>3" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""17 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                      Else
                           Msg = " Title " + Chr(34) + "网络干扰区 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：RxLev_s>17(RXLEV)且RxQual_s>3" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""17 至 20 (-95至-90dBm)"" display on ,""20 至 25 (-90至-85dBm)"" display on ,""25 至 30 (-85至-80dBm)"" display on ,""30 至 35 (-80至-75dBm)"" display on ,""35 至 40 (-75至-70dBm)"" display on ,""40 至 45 (-70至-65dBm)"" display on ,""45 至 50 (-65至-60dBm)"" display on ,""50 至 63 (-60至-47dBm)"" display on ,""63 以上 (大于-47dBm)"" display on"
                      End If
                  
                  End If
                  mapinfo.do "set legend window " & WinId & " Layer prev display on shades off symbols on lines off count on " & Msg
          Case 1201
                  If Val(mapinfo.eval("tableinfo(" & tblname & ",4)")) = 59 Then
                     mapinfo.do "select * from " & tblname & " where act_rlink <> """" into RadioLink"
                     mapinfo.do "Add Map window FrontWindow() Layer RadioLink"
                     Msg = " shade window FrontWindow() RadioLink with int((val(max_rlink)-val(act_rlink))/val(max_rlink)*100) ignore 0 ranges apply all use all Symbol (39,4219999,10,""MapInfo Cartographic"",0,0) "
                    'msg = " shade window FrontWindow() RadioLink with (val(max_rlink)-val(act_rlink)) ignore 0 values 1 Symbol (39,16711680,10,""MapInfo Cartographic"",0,0),2 Symbol (39,65280,10,""MapInfo Cartographic"",0,0),3 Symbol (39,255,10,""MapInfo Cartographic"",0,0),4 Symbol (39,16711935,10,""MapInfo Cartographic"",0,0),5 Symbol (39,16776960,10,""MapInfo Cartographic"",0,0),6 Symbol (39,65535,10,""MapInfo Cartographic"",0,0),7 Symbol (39,8388608,10,""MapInfo Cartographic"",0,0),8 Symbol (39,32768,10,""MapInfo Cartographic"",0,0),9 Symbol (39,128,10,""MapInfo Cartographic"",0,0),10 Symbol (39,8388736,10,""MapInfo Cartographic"",0,0) ,11 Symbol (39,8421376,10,""MapInfo Cartographic"",0,0) ,12 Symbol (39,32896,10,""MapInfo Cartographic"",0,0),13 Symbol (39,16744576,10,""MapInfo Cartographic"",0,0),"
                    'msg = msg + "14 Symbol (39,8454016,10,""MapInfo Cartographic"",0,0),15 Symbol (39,8421631,10,""MapInfo Cartographic"",0,0) ,16 Symbol (39,16744703,10,""MapInfo Cartographic"",0,0) ,17 Symbol (39,16777088,10,""MapInfo Cartographic"",0,0) ,18 Symbol (39,8454143,10,""MapInfo Cartographic"",0,0) ,19 Symbol (39,8405056,10,""MapInfo Cartographic"",0,0),20 Symbol (39,4227136,10,""MapInfo Cartographic"",0,0) ,21 Symbol (39,4210816,10,""MapInfo Cartographic"",0,0),22 Symbol (39,4111116,10,""MapInfo Cartographic"",0,0),23 Symbol (39,2222816,10,""MapInfo Cartographic"",0,0),24 Symbol (39,4219999,10,""MapInfo Cartographic"",0,0)"
                  ElseIf Val(mapinfo.eval("tableinfo(" & tblname & ",4)")) = 76 Or Val(mapinfo.eval("tableinfo(" & tblname & ",4)")) = 88 Or Val(mapinfo.eval("tableinfo(" & tblname & ",4)")) = 150 Then
                     Msg = " shade window FrontWindow() " & tblname & " with val(fer) ignore 0 ranges apply all use all Symbol (39,4219999,10,""MapInfo Cartographic"",0,0) "
                  End If
                  Msg = Msg + "1: 5 Symbol (39,65280,10,""MapInfo Cartographic"",0,0),5: 10 Symbol (39,255,10,""MapInfo Cartographic"",0,0),"
                  Msg = Msg + "10: 15 Symbol (39,16711935,10,""MapInfo Cartographic"",0,0),15: 20 Symbol (39,16776960,10,""MapInfo Cartographic"",0,0),20: 25 Symbol (39,65535,10,""MapInfo Cartographic"",0,0),"
                  Msg = Msg + "25: 30 Symbol (39,8388608,10,""MapInfo Cartographic"",0,0),30: 35 Symbol (39,32768,10,""MapInfo Cartographic"",0,0),35: 40 Symbol (39,128,10,""MapInfo Cartographic"",0,0),"
                  Msg = Msg + "40: 45 Symbol (39,8388736,10,""MapInfo Cartographic"",0,0),45: 50 Symbol (39,8421376,10,""MapInfo Cartographic"",0,0),50: 100 Symbol (39,16711680,10,""MapInfo Cartographic"",0,0)"
                  mapinfo.do Msg
                  
                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                  'msg = " Title " + Chr(34) + "无线链路丢失状况（RadioLink) " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending on ranges Font(""宋体"",0,9,0) """" display off"
                  Msg = " Title " + Chr(34) + "无线链路丢失状况(FER) " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "单位：%   标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) """" display off"
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
                mapinfo.do "set map redraw off"
                mapinfo.do "Set Map Layer " & tblname & " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                mapinfo.do "set map redraw on"

          Case 1202
                  mapinfo.do "select * from " & tblname & " where ucase$(hopping) = ""YES"" into Hopping"
                  If Val(mapinfo.eval("tableinfo(Hopping,8)")) = 0 Then
                     MsgBox "该路段不存在跳频", 64, "提示"
                     mapinfo.do "close table Hopping"
                  Else
                     mapinfo.do "Add Map window FrontWindow() Layer Hopping"
                     mapinfo.do "shade window frontwindow() Hopping with Ucase$(Hopping) values ""YES"" Symbol (87,16711680,10,""MapInfo Cartographic"",0,210)" ' default Symbol(32,16777215,0,""MapInfo Cartographic"",0,0)"
                     Msg = " Title " + Chr(34) + "跳频状态（Hopping) " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) """" display off"
                     mapinfo.do "set legend window FrontWindow() Layer prev " & Msg
                  End If
          Case 1203
               HoppingFrm.Show 1
          Case 1204
               TaLabel.Show 1
               
          Case 1116, 1118, 1113, 1114
                  AllRows = Val(mapinfo.eval("tableinfo(" & tblname & ",4)"))
                  If AllRows <> 88 And AllRows <> 150 Then
                     MsgBox "该文件不能观测 C/A 值", 64, "提示"
                     Exit Sub
                  End If
                  mapinfo.do "select * from " & tblname & " where rxle_neig2 >0 into Mytemp"
                  AllRows = mapinfo.eval("tableinfo(mytemp,8)")
                  mapinfo.do "close table mytemp"
                  If AllRows = 0 Then
                     MsgBox "该文件不能观测 C/A 值", 64, "提示"
                     Exit Sub
                  End If
          
               If Menu_Flag = 1116 Then
                  Msg = " shade window FrontWindow() " + tblname + " With RXLEV_F - rxle_neig2"
               ElseIf Menu_Flag = 1118 Then
                  Msg = " shade window FrontWindow() " + tblname + " With RXLEV_s - rxle_neig2"
               ElseIf Menu_Flag = 1113 Then
                  Msg = " shade window FrontWindow() " + tblname + " With RXLEV_f - rxle_neig1"
               Else
                  Msg = " shade window FrontWindow() " + tblname + " With RXLEV_s - rxle_neig1"
               End If
                  'msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) -63: -18 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) ,-18: -15 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,-15: -12 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,-12: -9 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,-9: 0 Symbol (39,8454016,8,""MapInfo Cartographic"",0,0) ,0: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) "
                  Msg = Msg + " ignore 0 ranges apply all use all Symbol (42,0,6,""MapInfo Weather"",0,0) -63: -18 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) ,-18: -15 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,-15: -12 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,-12: -9 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,-9: 0 Symbol (39,8454016,8,""MapInfo Cartographic"",0,0),0: 63 Symbol(42,0,6,""MapInfo Weather"",0,0) default Symbol (42,0,6,""MapInfo Weather"",0,0) "
                  mapinfo.do Msg
                  For i = 1 To mapinfo.eval("NumWindows()")     'win95
                      If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then     'win95
                         WinId = mapinfo.eval("windowid(" & i & ")")     'win95
                         If WinId = mapinfo.eval("frontwindow()") Then
                            Exit For
                         End If
                      End If     'win95
                  Next     'win95

                  If legendid = 0 Then     'win95
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"     'win95
                      mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
                      legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
                  End If     'win95
                            
                 If Menu_Flag = 1116 Then
                    Msg = " Title " + Chr(34) + "实时 C/A -1 分布图 (Full)" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "C/A=0(dBm)时忽略显示   标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
                ElseIf Menu_Flag = 1118 Then
                    Msg = " Title " + Chr(34) + "实时 C/A -1 分布图 (Sub)" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "C/A=0(dBm)时忽略显示   标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
                ElseIf Menu_Flag = 1113 Then
                
                   Msg = " Title " + Chr(34) + "邻频 Full C/A -2 观测" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
                Else
                   Msg = " Title " + Chr(34) + "邻频 Sub C/A -2 观测" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
                End If
                 mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
          
                mapinfo.do "set map redraw off"
                mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                mapinfo.do "set map redraw on"
                
                mapinfo.do "close table selection"
          Case 1117, 1119, 1112, 1115
                  AllRows = Val(mapinfo.eval("tableinfo(" & tblname & ",4)"))
                  If AllRows <> 88 And AllRows <> 150 Then
                     MsgBox "该文件不能观测 C/A 值", 64, "提示"
                     Exit Sub
                  End If
                              
                  mapinfo.do "select * from " & tblname & " where rxle_neig2 >0 into Mytemp"
                  AllRows = mapinfo.eval("tableinfo(mytemp,8)")
                  mapinfo.do "close table mytemp"
                  If AllRows = 0 Then
                     MsgBox "该文件不能观测 C/A 值", 64, "提示"
                     Exit Sub
                  End If
                  
               
               If Menu_Flag = 1117 Then
                  Msg = " shade window FrontWindow() " + tblname + " With RXLEV_f - rxle_neig3"
               ElseIf Menu_Flag = 1119 Then
                  Msg = " shade window FrontWindow() " + tblname + " With RXLEV_s - rxle_neig3"
               ElseIf Menu_Flag = 1112 Then
                  Msg = " shade window FrontWindow() " + tblname + " With RXLEV_f - rxle_neig4"
               Else
                  Msg = " shade window FrontWindow() " + tblname + " With RXLEV_s - rxle_neig4"
               End If
                  'msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) -120: -30 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,-30: -20 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,-20: -15 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,-15: -10 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,-10: 5 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,-5: 0 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,0: 5 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,5: 10 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,10: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 20 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,20: 30 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,30: 120 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                  'msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) -120: -30 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) ,-30: -20 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,-20: -15 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,-15: -10 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,-10: 5 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,-5: 0 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,0: 5 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,5: 10 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,10: 15 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,15: 20 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,20: 30 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,30: 120 Symbol (39,65280,8,""MapInfo Cartographic"",0,0)"
                  Msg = Msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) -63: -18 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) ,-18: -15 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,-15: -12 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,-12: -9 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,-9: 0 Symbol (39,8454016,8,""MapInfo Cartographic"",0,0) ,0: 63 Symbol(42,0,6,""MapInfo Weather"",0,0)"
                  mapinfo.do Msg
                  
                  For i = 1 To mapinfo.eval("NumWindows()")     'win95
                      If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then     'win95
                         WinId = mapinfo.eval("windowid(" & i & ")")     'win95
                         If WinId = mapinfo.eval("frontwindow()") Then
                            Exit For
                         End If
                      End If     'win95
                  Next     'win95

                  If legendid = 0 Then     'win95
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"     'win95
                      mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
                      legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
                  End If     'win95
                 If Menu_Flag = 1117 Then
                    Msg = " Title " + Chr(34) + "实时 C/A +1 分布图 (Full)" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "C/A=0(dBm)时忽略显示   标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
                 ElseIf Menu_Flag = 1119 Then
                    Msg = " Title " + Chr(34) + "实时 C/A +1 分布图 (Sub)" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "C/A=0(dBm)时忽略显示   标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
                 ElseIf Menu_Flag = 1112 Then
                    Msg = " Title " + Chr(34) + "邻频 Full C/A +2 观测" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
                 Else
                    Msg = " Title " + Chr(34) + "邻频 Sub C/A +2 观测" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "标注：BCCH" + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) ""其余全部"" display off "
                 End If
                 mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
          
                 mapinfo.do "set map redraw off"
                 mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                 mapinfo.do "set map redraw on"
                 mapinfo.do "close table selection"
          Case 1120
                  AllRows = Val(mapinfo.eval("tableinfo(" & tblname & ",4)"))
                  If AllRows = 59 Then
                     MsgBox "该文件不能观测 SQL 值", 64, "提示"
                     Exit Sub
                  End If

                  Msg = " shade window FrontWindow() " + tblname + " With val(SQI) ignore 0 "
                   
                   Msg = Msg + "values 22 Symbol (63,65280,8,""MapInfo Cartographic"",0,0) ,21 Symbol (63,8454016,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "20 Symbol (63,12648384,8,""MapInfo Cartographic"",0,0) ,19 Symbol (63,12632319,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "18 Symbol (63,255,8,""MapInfo Cartographic"",0,0) ,17 Symbol (63,16776960,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "16 Symbol (63,65535,8,""MapInfo Cartographic"",0,0) ,15 Symbol (63,8388608,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "14 Symbol (63,12648447,8,""MapInfo Cartographic"",0,0),13 Symbol (63,128,8,""MapInfo Cartographic"",0,0),12 Symbol (63,8388736,8,""MapInfo Cartographic"",0,0),11 Symbol (63,8421376,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "10 Symbol (63,16711935,8,""MapInfo Cartographic"",0,0),9 Symbol (63,16744576,8,""MapInfo Cartographic"",0,0),8 Symbol (63,15257855,8,""MapInfo Cartographic"",0,0),7 Symbol (63,8421631,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "6 Symbol (63,16744703,8,""MapInfo Cartographic"",0,0),5 Symbol (63,16777088,8,""MapInfo Cartographic"",0,0),4 Symbol (63,8454143,8,""MapInfo Cartographic"",0,0),3 Symbol (63,8405056,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "2 Symbol (63,4227136,8,""MapInfo Cartographic"",0,0),1 Symbol (63,4210816,8,""MapInfo Cartographic"",0,0),0 Symbol (63,8405120,8,""MapInfo Cartographic"",0,0),-1 Symbol (63,8421440,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "-2 Symbol (63,4227200,8,""MapInfo Cartographic"",0,0),-3 Symbol (63,16761024,8,""MapInfo Cartographic"",0,0),-4 Symbol (63,32896,8,""MapInfo Cartographic"",0,0),-5 Symbol (63,12632319,8,""MapInfo Cartographic"",0,0),"
                   Msg = Msg + "-6 Symbol (63,16761087,8,""MapInfo Cartographic"",0,0),-7 Symbol (63,16777152,8,""MapInfo Cartographic"",0,0),-8 Symbol (63,16711680,8,""MapInfo Cartographic"",0,0) ,"
                   Msg = Msg + "-9 Symbol (63,16711680,8,""MapInfo Cartographic"",0,0) ,-10 Symbol (63,16711680,8,""MapInfo Cartographic"",0,0), "
                   Msg = Msg + "-11 Symbol (63,16711680,8,""MapInfo Cartographic"",0,0) ,-12 Symbol (63,16711680,8,""MapInfo Cartographic"",0,0) "
                   Msg = Msg + "default Symbol (63,16777215,8,""MapInfo Cartographic"",0,0)"
                   mapinfo.do Msg

                  If legendid = 0 Then
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                      mapinfo.do "Create Legend From Window  Frontwindow()"
                      legendid = mapinfo.eval("windowinfo(1009,12)")
                  End If
                     
                  Msg = " Title " + Chr(34) + "SQL 下行话音质量观测 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) """" display off"
                                   
                  mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
        Case 1999
           frmForecast.Show 1
        Case 1111
             frmLabel.Show 1
        Case 55550
            For i = 1 To mapinfo.eval("NumTables()")
                If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "SWITCHPOINT" Then
                   mapinfo.do "close table SwitchPoint"
                   Exit For
                End If
            Next
            If Dir(Gsm_Path + "\user\SwitchPoint.tab", 0) <> "" Then
               Kill Gsm_Path + "\user\SwitchPoint.*"
            End If
            mapinfo.do "Select * from " & tblname & " group by Bcch_serv into mytemp"
            mapinfo.do "commit table mytemp as " + Chr(34) + Gsm_Path + "\user\SwitchPoint.tab" + Chr(34)
            mapinfo.do "close table mytemp"
            mapinfo.do "open table " + Chr(34) + Gsm_Path + "\user\SwitchPoint.tab" + Chr(34)
            mapinfo.do "fetch first from SwitchPoint"
            PreBcch = mapinfo.eval("SwitchPoint.bcch_serv")
            For i = 1 To mapinfo.eval("tableinfo(SwitchPoint,8)")
                If mapinfo.eval("SwitchPoint.bcch_serv") = 0 Or PreBcch = 0 Then
                   mapinfo.do "delete from SwitchPoint where rowid= " & Format(i)
                Else
                   If Not (mapinfo.eval("SwitchPoint.bcch_serv") > 124 And PreBcch <= 124 Or mapinfo.eval("SwitchPoint.bcch_serv") <= 124 And PreBcch > 124) Then
                      mapinfo.do "delete from SwitchPoint where rowid= " & Format(i)
                   End If
                End If
                PreBcch = mapinfo.eval("SwitchPoint.bcch_serv")
                mapinfo.do "fetch next from SwitchPoint"
            Next
                  If Val(mapinfo.eval("tableinfo(SwitchPoint,8)")) = 0 Then
                      MsgBox "该路段不存在GSM/DCS切换点", 64, "提示"
                      mapinfo.do "close talbe SwitchPoint"
                      Exit Sub
                  End If
            mapinfo.do "Create Map For SwitchPoint CoordSys Earth Projection 1, 0"
            mapinfo.do "Set Style Symbol MakeSymbol(33,0,2)" '
            mapinfo.do "update SwitchPoint set Obj= CreatePoint(Lon, Lat)"
            mapinfo.do "Add Map window FrontWindow() Layer SwitchPoint"
            'mapinfo.do "shade window frontwindow() SwitchPoint with bcch_serv values ""YES"" Symbol (87,16711680,10,""MapInfo Cartographic"",0,210)"
            'mapinfo.do "shade window frontwindow() SwitchPoint with bcch_serv values  Symbol(54,16711935,12,""MapInfo Symbols"",0,0)"
            mapinfo.do "shade window frontwindow() SwitchPoint with Bcch_serv ranges apply all use color Symbol(54,16711935,12,""MapInfo Symbols"",0,0) 0: 999 Symbol(54,16711935,12,""MapInfo Symbols"",0,0) default Symbol(54,16711935,12,""MapInfo Symbols"",0,0)"
            'msg = " Title " + Chr(34) + "GSM/DCS切换点 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending on ranges Font(""宋体"",0,9,0) """" display off"
            Msg = " Title " + Chr(34) + "GSM/DCS切换点 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,255) ascending on ranges Font(""宋体"",0,9,0) auto display off ,""切换点"" display on"
            mapinfo.do "set legend window FrontWindow() Layer prev " & Msg
            mapinfo.do "commit table SwitchPoint"
         Case 9001, 9003, 9004
              SelCond.Show 1
                       
         Case 9005, 9006, 9007, 9008
              
              If Menu_Flag = 9005 Then
                 mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "HANDOVER FAILURE" + Chr(34) + " and Rxle_same1>0 group by RXLE_SAME1 order by Rxle_same1 into Selection1"
              ElseIf Menu_Flag = 9006 Then
                 mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "DISCONNECT" + Chr(34) + " and Rxle_same1>0 group by RXLE_SAME1 order by Rxle_same1 into Selection1"
              ElseIf Menu_Flag = 9007 Then
                 mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "RELEASE" + Chr(34) + " and Rxle_same1>0 group by RXLE_SAME1 order by Rxle_same1 into Selection1"
              ElseIf Menu_Flag = 9008 Then
                 mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "CHANNEL RELEASE" + Chr(34) + " and Rxle_same1>0 group by RXLE_SAME1 order by Rxle_same1 into Selection1"
              End If
              row = Val(mapinfo.eval("tableinfo(Selection1,8)"))
              If row = 0 Then
                 Exit Sub
              End If
              ReDim CauseValue(1 To row) As Integer
              ReDim CVString(1 To row) As String
              For i = 1 To row
                  CauseValue(i) = mapinfo.eval("Selection1.Rxle_same1")
                  mapinfo.do "fetch next from Selection1"
              Next
              mapinfo.do "close table Selection1"
              If Menu_Flag = 9005 Then
                 mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "HANDOVER FAILURE" + Chr(34) + " into Result"
              ElseIf Menu_Flag = 9006 Then
                 mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "DISCONNECT" + Chr(34) + " into Result"
              ElseIf Menu_Flag = 9007 Then
                 mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "RELEASE" + Chr(34) + " into Result"
              ElseIf Menu_Flag = 9008 Then
                 mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "CHANNEL RELEASE" + Chr(34) + " into Result"
              End If
              If Menu_Flag = 9005 Or Menu_Flag = 9008 Then
                 For i = 1 To row
                     Select Case CauseValue(i)
                         Case 0
                             CVString(i) = "Normal event"
                         Case 1
                             CVString(i) = "Abnormal release,unspecified"
                         Case 2
                             CVString(i) = "Abnormal release,channel unacceptable"
                         Case 3
                             CVString(i) = "Abnormal release,timer expired"
                         Case 4
                             CVString(i) = "Abnormal release,no activity on the radio path"
                         Case 5
                             CVString(i) = "preemptive release"
                         Case 8
                             CVString(i) = "handover impossible,timing advance out of range"
                         Case 9
                             CVString(i) = "Channel mode unacceptable"
                         Case 10
                             CVString(i) = "Frequency not implemented"
                         Case 65
                             CVString(i) = "call already cleared"
                         Case 95
                             CVString(i) = "semantically incorrect message"
                         Case 96
                             CVString(i) = "invalid mandatory information"
                         Case 97
                             CVString(i) = "Message type non-existent or not implenmented"
                         Case 98
                             CVString(i) = "Message type not compatible with protocol state"
                         Case 100
                             CVString(i) = "Conditional IE error"
                         Case 101
                             CVString(i) = "No cell allocation available"
                         Case 111
                             CVString(i) = "protocol error unspecified"
                     End Select
                 Next
              Else
                 For i = 1 To row
                     Select Case CauseValue(i)
                         Case 1
                             CVString(i) = "Unassiagned number"
                         Case 3
                             CVString(i) = "No route to destination"
                         Case 6
                             CVString(i) = "Channel unacceptable"
                         Case 16
                             CVString(i) = "Normal clearing"
                         Case 17
                             CVString(i) = "User busy"
                         Case 18
                             CVString(i) = "No user responding"
                         Case 19
                             CVString(i) = "User alerting,no answer"
                         Case 21
                             CVString(i) = "Call rejected"
                         Case 22
                             CVString(i) = "Number changed"
                         Case 26
                             CVString(i) = "Non selected user clearing"
                         Case 27
                             CVString(i) = "Destination out of order "
                         Case 28
                             CVString(i) = "Incomplete number"
                         Case 29
                             CVString(i) = "Facility rejected"
                         Case 30
                             CVString(i) = "Response to status enquiry"
                         Case 31
                             CVString(i) = "Normal,unspecified"
                         Case 34
                             CVString(i) = "34,No circuit/channel available"
                         Case 38
                             CVString(i) = "38,Network out of order"
                         Case 41
                             CVString(i) = "41,Temporary failure"
                         Case 42
                             CVString(i) = "42,Switching equipment congestion"
                         Case 43
                             CVString(i) = "43,Access information discarded"
                         Case 44
                             CVString(i) = "44,Requested circuit/channel not available"
                         Case 47
                             CVString(i) = "47,Resources unavailable,unspecified"
                         Case 49
                             CVString(i) = "49,Quality of service unavailable"
                         Case 50
                             CVString(i) = "50,Requested facility not subscribed"
                         Case 55
                             CVString(i) = "55,Incoming calls barred within the CUG"
                         Case 57
                             CVString(i) = "57,Bearer capability not authorized"
                         Case 58
                             CVString(i) = "58,Bearer capability not presently available"
                         Case 63
                             CVString(i) = "63,Service or option not available,unspecified"
                         Case 65
                             CVString(i) = "65,Bearer service not implemented"
                         Case 68
                             CVString(i) = "68,ACM equal to or greater than ACMmax"
                         Case 69
                             CVString(i) = "69,Requested facility not implemented"
                         Case 70
                             CVString(i) = "70,Only restricted digital information bearer"
                         Case 79
                             CVString(i) = "79,Service or option not implemented"
                         Case 81
                             CVString(i) = "81,Invalid transaction identrfier value"
                         Case 87
                             CVString(i) = "87,User not member of CUG"
                         Case 88
                             CVString(i) = "88,Incompatible destination"
                         Case 91
                             CVString(i) = "91,Invalid mandatory information"
                         Case 95
                             CVString(i) = "95,Semantically incorrect message"
                         Case 96
                             CVString(i) = "96,Invalid mandatory information"
                         Case 97
                             CVString(i) = "97,Message type non-existent or not implemented"
                         Case 98
                             CVString(i) = "98,Message type not compatible with protocol state"
                         Case 99
                             CVString(i) = "99,Information element non-existent or not implemented"
                         Case 100
                             CVString(i) = "100,Conditional IE error "
                         Case 101
                             CVString(i) = "101,Message not compatible with protocol state"
                         Case 102
                             CVString(i) = "102,Recovery on timer expiry"
                         Case 111
                             CVString(i) = "111,Protocol error,unspecified"
                         Case 127
                             CVString(i) = "127,Interworking,unspecified"
                     End Select
                 Next
              End If
              mapinfo.do "Add Map window FrontWindow() Layer  Result"
              Msg = "shade window FrontWindow() Result with RXLE_SAME1 values "
              For i = 1 To row
                  Msg = Msg & Format(CauseValue(i)) & " Symbol (41," & Format(MyRndColor(i)) & " ,8,""MapInfo Cartographic"",0,0),"
              Next
              Msg = Left(Msg, Len(Msg) - 1)
              mapinfo.do Msg
              
              If legendid = 0 Then
                 mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                 mapinfo.do "Create Legend From Window  Frontwindow()"
                 legendid = mapinfo.eval("windowinfo(1009,12)")
              End If
              
              If Menu_Flag = 9005 Then
                 Msg = " Title " + Chr(34) + "HANDOVER FAILURE 事件原因" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off ,"
              ElseIf Menu_Flag = 9006 Then
                 Msg = " Title " + Chr(34) + "DISCONNECT 事件原因" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off ,"
              ElseIf Menu_Flag = 9007 Then
                 Msg = " Title " + Chr(34) + "RELEASE 事件原因" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off ,"
              ElseIf Menu_Flag = 9008 Then
                 Msg = " Title " + Chr(34) + "CHANNEL RELEASE 事件原因" + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off ,"
              End If
              For i = 1 To row
                  Msg = Msg + Chr(34) + Format(CauseValue(i)) & ": " & CVString(i) + Chr(34) + " display on,"
              Next
              Msg = Left(Msg, Len(Msg) - 1)
              mapinfo.do "set legend window FrontWindow()  Layer prev " & Msg
        Case 9010
             frmStrongNcell.Show 1
        Case 9020
             frmHandoverCause.Show 1
        Case 9002
             frmNewDisturb.Show 1
        Case 991028      '——今天我生日！
             frmCallAnalyse.Show 1
        Case 991103
             mapinfo.do "select * from " & tblname & " where rxlev_s>" & Format(110 - 70) & " into BestCover"
             MyPoint = mapinfo.eval("tableinfo(BestCover,8)")
             If MyPoint = 0 Then
                MsgBox "该路段不存在最佳主小区覆盖", 64, "提示"
                mapinfo.do "close table BestCover"
                GoTo VER_OUT
             End If
             
             mapinfo.do "Add Map window FrontWindow() Layer BestCover"
             mapinfo.do " shade window FrontWindow() BestCover With RXLEV_s ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0)"
                  
                  For i = 1 To mapinfo.eval("NumWindows()")     'win95
                      If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then     'win95
                         WinId = mapinfo.eval("windowid(" & i & ")")     'win95
                         If WinId = mapinfo.eval("frontwindow()") Then
                            Exit For
                         End If
                      End If     'win95
                  Next     'win95

                  If legendid = 0 Then     'win95
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"     'win95
                      mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
                      legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
                  End If     'win95
                  AllRows = mapinfo.eval("tableinfo(" & tblname & ",8)")
                  MyPercent = Format(MyPoint / AllRows, "Percent")
                  Msg = " Title " + Chr(34) + "最佳主小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：RxLev>-70dBm  标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""40 以上 (大于-70dBm)  [" & MyPercent & "]"" display on"
                  mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on " & Msg
                mapinfo.do "close table selection"
                mapinfo.do "set map redraw off"
                mapinfo.do "Set Map Layer ""BestCover"" Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                mapinfo.do "set map redraw on"
        
        Case 991104
             mapinfo.do "select * from " & tblname & " where rxlev_s<=" & Format(110 - 70) & " and rxlev_s>" & Format(110 - 83) & " into BetterCover"
             MyPoint = mapinfo.eval("tableinfo(BetterCover,8)")
             If MyPoint = 0 Then
                MsgBox "该路段不存在次佳主小区覆盖", 64, "提示"
                mapinfo.do "close table BetterCover"
                GoTo VER_OUT
             End If
             
             mapinfo.do "Add Map window FrontWindow() Layer BetterCover"
             mapinfo.do " shade window FrontWindow() BetterCover With RXLEV_s ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 40: 27 Symbol (39,255,8,""MapInfo Cartographic"",0,0)"
                  
                  For i = 1 To mapinfo.eval("NumWindows()")     'win95
                      If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then     'win95
                         WinId = mapinfo.eval("windowid(" & i & ")")     'win95
                         If WinId = mapinfo.eval("frontwindow()") Then
                            Exit For
                         End If
                      End If     'win95
                  Next     'win95

                  If legendid = 0 Then     'win95
                      mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"     'win95
                      mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
                      legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
                  End If     'win95
                  AllRows = mapinfo.eval("tableinfo(" & tblname & ",8)")
                  MyPercent = Format(MyPoint / AllRows, "Percent")
                  Msg = " Title " + Chr(34) + "次佳主小区覆盖 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "条件：-70dBm≥RxLev>-83dBm  标注：BCCH" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""27 至 40 (-83至-70dBm)  [" & MyPercent & "]"" display on"
                  mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on " & Msg
                mapinfo.do "close table selection"
                mapinfo.do "set map redraw off"
                mapinfo.do "Set Map Layer ""BetterCover"" Label Visibility Font (""Arial"",257,8,8421376,16777215) With bcch_serv Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
                mapinfo.do "set map redraw on"
        Case 991105
                frmGDCover.Show 1
        Case 991106
                frmBYCover.Show 1
        Case 317, 991121
                frmTACover.Show 1
        Case 991107
                frmBcchAdjust.Show 1
        Case 5003
                frmGDNcell.Show 1
        Case 991108
                'mapinfo.do "select * from " & tblname & " where ci_serv= """ & MyNRSelCellCI & """ and bcch_serv= " & Format(MyNRSelCellBcch) & " into NRSelection"
                frmNcellA.Show 1
                If MyNRSelCellCI = "" Or MyNRSelCellBcch = 0 Or MyNRSelCellLac = "" Then
                    GoTo VER_OUT
                End If
            Dim LinCell1 As Boolean, LinCell2 As Boolean, LinCell3 As Boolean
                'mapinfo.do "select * from " & tblname & " where ci_serv= """ & MyNRSelCellCI & """ and  bcch_serv= " & Format(MyNRSelCellBcch) & " and bsic_serv= " & Format(MyNRSelCellBsic) & " and  lac_serv= """ & Format(MyNRSelCellLac) & """ into NRSelection"
                mapinfo.do "select * from " & tblname & " where ci_serv= """ & MyNRSelCellCI & """ and  bcch_serv= " & Format(MyNRSelCellBcch) & " and bsic_serv= " & Format(MyNRSelCellBsic) & " and  lac_serv= """ & Format(MyNRSelCellLac) & """ and (not(bcch_n1=" & Format(MyNRSelCellBcch) & " and bsic_n1=" & Format(MyNRSelCellBsic) & ") and not(bcch_n2=" & Format(MyNRSelCellBcch) & " and bsic_n2=" & Format(MyNRSelCellBsic) & ")  and not(bcch_n3=" & Format(MyNRSelCellBcch) & " and bsic_n3=" & Format(MyNRSelCellBsic) & ")  and not (bcch_n4=" & Format(MyNRSelCellBcch) & " and bsic_n4=" & Format(MyNRSelCellBsic) & ")  and not (bcch_n5=" & Format(MyNRSelCellBcch) & " and bsic_n5=" & Format(MyNRSelCellBsic) & ")  and not (bcch_n6=" & Format(MyNRSelCellBcch) & " and bsic_n6=" & Format(MyNRSelCellBsic) & ")) into NRSelection_1"
                If mapinfo.eval("tableinfo(NRSelection_1,8)") = 0 Then
                    'MsgBox "不存在该小区覆盖，无法进行邻小区合理性统计", 64, "提示"
                    LinCell1 = True
                    mapinfo.do "close table NRSelection_1"
                Else
                    CurrentNcellRS = 1
                    frmNcellRS.Tag = 1
                    frmNcellRS.Show
                    SetWindowPos frmNcellRS.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                End If
                If MyNRSelCellCI_2 <> "" Or MyNRSelCellBcch_2 <> 0 Then
                    MyNRSelCellCI = MyNRSelCellCI_2
                    MyNRSelCellBcch = MyNRSelCellBcch_2
                    MyNRSelCellBsic = MyNRSelCellBsic_2
                    mapinfo.do "select * from " & tblname & " where ci_serv= """ & MyNRSelCellCI & """ and  bcch_serv= " & Format(MyNRSelCellBcch) & " and bsic_serv= " & Format(MyNRSelCellBsic) & " and  lac_serv= """ & Format(MyNRSelCellLac) & """ and (not(bcch_n1=" & Format(MyNRSelCellBcch) & " and bsic_n1=" & Format(MyNRSelCellBsic) & ") and not(bcch_n2=" & Format(MyNRSelCellBcch) & " and bsic_n2=" & Format(MyNRSelCellBsic) & ")  and not(bcch_n3=" & Format(MyNRSelCellBcch) & " and bsic_n3=" & Format(MyNRSelCellBsic) & ")  and not (bcch_n4=" & Format(MyNRSelCellBcch) & " and bsic_n4=" & Format(MyNRSelCellBsic) & ")  and not (bcch_n5=" & Format(MyNRSelCellBcch) & " and bsic_n5=" & Format(MyNRSelCellBsic) & ")  and not (bcch_n6=" & Format(MyNRSelCellBcch) & " and bsic_n6=" & Format(MyNRSelCellBsic) & ")) into NRSelection_2"
                    If mapinfo.eval("tableinfo(NRSelection_2,8)") = 0 Then
                        'MsgBox "不存在该小区覆盖，无法进行邻小区合理性统计", 64, "提示"
                        LinCell2 = True
                        mapinfo.do "close table NRSelection_2"
                    Else
                        CurrentNcellRS = 2
                        frmNcellRS_2.Tag = 2
                        frmNcellRS_2.Show
                        frmNcellRS_2.Top = frmNcellRS_2.Top + 200
                        frmNcellRS_2.Left = frmNcellRS_2.Left - 200
                        SetWindowPos frmNcellRS_2.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                    End If
                End If
                If MyNRSelCellCI_3 <> "" Or MyNRSelCellBcch_3 <> 0 Then
                    MyNRSelCellCI = MyNRSelCellCI_3
                    MyNRSelCellBcch = MyNRSelCellBcch_3
                    MyNRSelCellBsic = MyNRSelCellBsic_3
                    mapinfo.do "select * from " & tblname & " where ci_serv= """ & MyNRSelCellCI & """ and  bcch_serv= " & Format(MyNRSelCellBcch) & " and bsic_serv= " & Format(MyNRSelCellBsic) & " and  lac_serv= """ & Format(MyNRSelCellLac) & """ and (not(bcch_n1=" & Format(MyNRSelCellBcch) & " and bsic_n1=" & Format(MyNRSelCellBsic) & ") and not(bcch_n2=" & Format(MyNRSelCellBcch) & " and bsic_n2=" & Format(MyNRSelCellBsic) & ")  and not(bcch_n3=" & Format(MyNRSelCellBcch) & " and bsic_n3=" & Format(MyNRSelCellBsic) & ")  and not (bcch_n4=" & Format(MyNRSelCellBcch) & " and bsic_n4=" & Format(MyNRSelCellBsic) & ")  and not (bcch_n5=" & Format(MyNRSelCellBcch) & " and bsic_n5=" & Format(MyNRSelCellBsic) & ")  and not (bcch_n6=" & Format(MyNRSelCellBcch) & " and bsic_n6=" & Format(MyNRSelCellBsic) & ")) into NRSelection_3"
                    If mapinfo.eval("tableinfo(NRSelection_3,8)") = 0 Then
                        'MsgBox "不存在该小区覆盖，无法进行邻小区合理性统计", 64, "提示"
                        LinCell3 = True
                        mapinfo.do "close table NRSelection_3"
                    Else
                        CurrentNcellRS = 3
                        frmNcellRS_3.Tag = 3
                        frmNcellRS_3.Top = frmNcellRS_3.Top + 400
                        frmNcellRS_3.Left = frmNcellRS_3.Left - 400
                        frmNcellRS_3.Show
                        SetWindowPos frmNcellRS_3.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                    End If
                End If
                If LinCell1 Then
                    MsgBox "不存在小区" & MyNRSelCellCI & "覆盖，无法进行邻小区合理性统计", 64, "提示"
                End If
                If LinCell2 Then
                    MsgBox "不存在小区" & MyNRSelCellCI_2 & "覆盖，无法进行邻小区合理性统计", 64, "提示"
                End If
                If LinCell3 Then
                    MsgBox "不存在小区" & MyNRSelCellCI_3 & "覆盖，无法进行邻小区合理性统计", 64, "提示"
                End If
        Case 991109
                If mapinfo.eval("tableinfo(" & tblname & ",4)") <> 150 Then
                    MsgBox "该数据是旧数据，无法显示切换前后参数", 64, "提示"
                    GoTo VER_OUT
                End If
                HOParaFlag = 0
                frmGDNcell.Show 1
                If HOParaFlag > 0 Then
                    frmHOPara.Show
                    SetWindowPos frmHOPara.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
                End If
        Case 991110
                If mapinfo.eval("tableinfo(" & tblname & ",4)") <> 150 Then
                    MsgBox "该数据是旧数据，无法进行切换统计", 64, "提示"
                Else
                    mapinfo.do "select * from " & tblname & " where Left$(mark1,3)=""HOA"" or Left$(mark1,3)=""HOS"" or Left$(mark1,3)=""HOF"" or Left$(mark1,3)=""HSC"" or Left$(mark1,3)=""HFC"" into HOStat"
                    If mapinfo.eval("tableinfo(HOStat,8)") = 0 Then
                        MsgBox "不存在切换过程或该数据是旧数据，无法进行切换统计", 64, "提示"
                    Else
                        frmHOStat.Show
                    End If
                End If
        Case 991112
                frmDurbSelBcch.Show 1
        Case 991021
                If mapinfo.eval("tableinfo(" & tblname & ",4)") <> 150 Then
                    MsgBox "该数据是旧数据，无法进行通话过程统计", 64, "提示"
                Else
                    frmCallAnalyse.Show 1
                End If
        Case 991120
                If UCase(Right(tblname, 1)) = "F" Or UCase(Right(tblname, 1)) = "E" Then
                    MsgBox "该数据不是全转换数据，第三层信令无法完整显示", 64, "提示"
                End If
                If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
                    MapForm.WindowState = 0
                End If
                MapForm.Move 0, 10, 7000, 4050
                Load Graph_Call
                Graph_Call.Move 0, 4050, 7000, 3495
                Load FrmLayer3
                FrmLayer3.Move 7000, 10, 4800, 4050
                Load frmCallEvent
                frmCallEvent.Move 7000, 4050, 4800, 3495

      End Select
   End If
VER_OUT:
   Unload SelTable
End Sub

