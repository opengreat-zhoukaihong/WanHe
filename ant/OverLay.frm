VERSION 5.00
Begin VB.Form OverLay 
   Caption         =   "ȷ��С�����Ƿ�Χ"
   ClientHeight    =   3480
   ClientLeft      =   6480
   ClientTop       =   450
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OverLay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "����ѡ��"
      Height          =   1170
      Left            =   210
      TabIndex        =   14
      Top             =   2235
      Width           =   2955
      Begin VB.CheckBox Check1 
         Caption         =   "���ǰ��� BSIC=99 �ĵ�"
         Height          =   240
         Left            =   405
         TabIndex        =   17
         Top             =   765
         Value           =   1  'Checked
         Width           =   2205
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Sub"
         Height          =   240
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   375
         Width           =   750
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Full"
         Height          =   240
         Index           =   0
         Left            =   375
         TabIndex        =   15
         Top             =   375
         Value           =   -1  'True
         Width           =   660
      End
   End
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O ȷ��"
      Height          =   320
      Left            =   3345
      TabIndex        =   9
      Top             =   690
      Width           =   1080
   End
   Begin VB.CommandButton SBSCancel 
      Caption         =   "&C ȡ��"
      Height          =   320
      Left            =   3345
      TabIndex        =   8
      Top             =   1110
      Width           =   1080
   End
   Begin VB.ComboBox Combo1 
      DataField       =   " "
      DataSource      =   " "
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   150
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      Caption         =   "С��ѡ��"
      Height          =   1590
      Left            =   210
      TabIndex        =   0
      Top             =   585
      Width           =   2955
      Begin VB.OptionButton Cell_3 
         Caption         =   "С��-1"
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   1185
         Width           =   885
      End
      Begin VB.OptionButton Cell_2 
         Caption         =   "С��-1"
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   795
         Width           =   885
      End
      Begin VB.OptionButton Cell_1 
         Caption         =   "С��-1"
         Height          =   240
         Left            =   360
         TabIndex        =   11
         Top             =   390
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.TextBox Arfcn_1 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Height          =   240
         Left            =   2085
         TabIndex        =   3
         Text            =   "  "
         Top             =   405
         Width           =   420
      End
      Begin VB.TextBox Arfcn_2 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Height          =   240
         Left            =   2085
         TabIndex        =   2
         Text            =   " "
         Top             =   780
         Width           =   420
      End
      Begin VB.TextBox Arfcn_3 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Height          =   240
         Left            =   2085
         TabIndex        =   1
         Top             =   1170
         Width           =   420
      End
      Begin VB.Label Cell2Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1485
         TabIndex        =   6
         Top             =   825
         Width           =   540
      End
      Begin VB.Label Cell3Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1485
         TabIndex        =   5
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Cell1Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1485
         TabIndex        =   4
         Top             =   420
         Width           =   540
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "��վѡ��"
      Height          =   180
      Index           =   0
      Left            =   195
      TabIndex        =   10
      Top             =   210
      Width           =   900
   End
End
Attribute VB_Name = "OverLay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BaseArfcn As String, BaseBsic As String, BaseCi As String

Private Sub Form_Load()
    Dim i As Integer, BaseRows As Integer

    On Error Resume Next
    BaseRows = Val(mapinfo.eval("tableinfo(base,8)"))
    mapinfo.do "fetch First from base"
    For i = 1 To BaseRows
        Combo1.AddItem mapinfo.eval("base.bs_NAME")
        mapinfo.do "fetch next from base"
    Next
    mapinfo.do "fetch First from base"
    Combo1.ListIndex = 0

End Sub

Private Sub SBSCancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Combo1_Click()
    Dim finds As Integer

    On Error Resume Next
    If Combo1.Text <> "" Then
       mapinfo.do "select * from base where bs_name = " & Chr(34) & Trim(Combo1.Text) & Chr(34) & "into temp"
       Arfcn_1.Text = mapinfo.eval("temp.BCCH_1")
       Arfcn_2.Text = mapinfo.eval("temp.BCCH_2")
       Arfcn_3.Text = mapinfo.eval("temp.BCCH_3")
       mapinfo.do "close table temp"
    End If

End Sub

Private Sub SBSOK_Click()
    Dim i As Integer, TempRows As Integer
    Dim MyMsg As String
    Dim OpenTableNum As Integer
    Dim CloseTableName(1 To 6) As String
    Dim j As Integer
        
    On Error Resume Next
    Screen.MousePointer = 11
    If Combo1.Text <> "" Then
       mapinfo.do "select * from base where bs_name = " & Chr(34) & Trim(Combo1.Text) & Chr(34) & "into basetemp"
       OverLay.Hide
       OpenTableNum = Val(mapinfo.eval("NumTables()"))
       j = 0
       For i = 1 To OpenTableNum
           If InStr(UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")), "SERVINGLAY") > 0 Or InStr(UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")), "NEIGHBORLAY") > 0 Then
              CloseTableName(j + 1) = mapinfo.eval("tableinfo(" & Format(i) & ",1)")
              j = j + 1
              'mapinfo.do "close table " & mapinfo.eval("tableinfo(" & Format(i) & ",1)")
           End If
       Next
       For i = 1 To j
           mapinfo.do "close table " & CloseTableName(i)
       Next
       If Cell_1.Value = True Then
          BaseArfcn = mapinfo.eval("basetemp.BCCH_1")
          If BaseArfcn = 0 Then
             GoTo ExitMySub
          End If
          BaseBsic = mapinfo.eval("basetemp.BSIC_1")
          BaseCi = mapinfo.eval("basetemp.ci_1")
          MakeThematic ("1")
       ElseIf Cell_2.Value = True Then
          BaseArfcn = mapinfo.eval("basetemp.BCCH_2")
          If BaseArfcn = 0 Then
             GoTo ExitMySub
          End If
          BaseBsic = mapinfo.eval("basetemp.BSIC_2")
          BaseCi = mapinfo.eval("basetemp.ci_2")
          MakeThematic ("2")
       Else
          BaseArfcn = mapinfo.eval("basetemp.BCCH_3")
          If BaseArfcn = 0 Then
             GoTo ExitMySub
          End If
          BaseBsic = mapinfo.eval("basetemp.BSIC_3")
          BaseCi = mapinfo.eval("basetemp.ci_3")
          MakeThematic ("3")
       End If
       mapinfo.do "close table basetemp"
    End If
    
ExitMySub:
    Screen.MousePointer = 0
    Unload Me
End Sub

Sub MakeThematic(SelectCell As String)
    Dim MyMsg As String
    On Error Resume Next
    
    Gsm_FileName = Gsm_Path + "\NeighborLay" & SelectCell & ".tab"
    mapinfo.do "fetch first from  " & tblname
    mapinfo.do "select Time, Lon, Lat,Ci_Serv,bcch_serv,Bsic_serv,Rxlev_f,Rxlev_s From " + tblname + " where CI_SERV=  " + Chr(34) + BaseCi + Chr(34) + " into ServingLay" & SelectCell
    mapinfo.do "commit table ServingLay" & SelectCell & " as " + Chr(34) + Gsm_FileName + Chr(34)
    mapinfo.do "add map auto layer ServingLay" & SelectCell
    If Option4(0).Value = True Then
       MyMsg = "shade window FrontWindow() ServingLay" & SelectCell & " With RXLEV_F "
    Else
       MyMsg = "shade window FrontWindow() ServingLay" & SelectCell & " With RXLEV_s "
    End If
    If Legend_Tog = 0 Then
       MyMsg = MyMsg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,17: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
    Else
       MyMsg = MyMsg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  90: 63 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
    End If
    mapinfo.do MyMsg
    If legendid = 0 Then     'win95
       mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"     'win95
       mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
       legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
    End If     'win95
    If Legend_Tog = 0 Then
       MyMsg = " Title " + Chr(34) + "ȷ��С������ " + tblname + " ServingLay" + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""����"",0,9,0) ascending on ranges Font (""����"",0,9,0) ""����ȫ��"" display off ,""0 �� 17 (-110��-93dBm)"" display on ,""17 �� 27 (-93��-83dBm)"" display on ,""27 �� 63 (-83��-47dBm)"" display on ,""63 ���� (����-47dBm)"" display on"
    Else
       MyMsg = " Title " + Chr(34) + "ȷ��С������ " + tblname + " ServingLay" + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""����"",0,9,0) ascending on ranges Font (""����"",0,9,0) ""����ȫ��"" display off ,""0 �� 5 (-110��-105dBm)"" display on ,""5 �� 10 (-105��-100dBm)"" display on ,""10 �� 15 (-100��-95dBm)"" display on ,""15 �� 20 (-95��-90dBm)"" display on ,""20 �� 25 (-90��-85dBm)"" display on ,""25 �� 30 (-85��-80dBm)"" display on ,""30 �� 35 (-80��-75dBm)"" display on ,""35 �� 40 (-75��-70dBm)"" display on ,""40 �� 45 (-70��-65dBm)"" display on ,""45 �� 50 (-65��-60dBm)"" display on ,""50 �� 63 (-60��-47dBm)"" display on ,""63 ���� (����-47dBm)"" display on"
    End If
    mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on " & MyMsg
          
    mapinfo.do "open table " + Chr(34) + Gsm_FileName + Chr(34)
    mapinfo.do "Alter Table ""NeighborLay" & SelectCell & """ ( drop Rxlev_s rename Bcch_Serv Bcch,Bsic_serv Bsic,Rxlev_f Rxlev ) Interactive"
    mapinfo.do "delete from NeighborLay" & SelectCell
    mapinfo.do "commit table NeighborLay" & SelectCell
    mapinfo.do "pack table NeighborLay" & SelectCell & " Graphic Data Data Interactive "
    For i = 1 To 6
        If Check1.Value = 1 Then
           mapinfo.do "select * from " + tblname + " where (bsic_n" & Format(i) & "= " + Format(BaseBsic) + " and bcch_n" & Format(i) & "=" + Format(BaseArfcn) + ") or (bsic_n" & Format(i) & " = 99 and bcch_n" & Format(i) & "=" + Format(BaseArfcn) + ") into overlaytemp"
        Else
           mapinfo.do "select * from " + tblname + " where (bsic_n" & Format(i) & "= " + Format(BaseBsic) + " and bcch_n" & Format(i) & "=" + Format(BaseArfcn) + ") into overlaytemp"
        End If
        TempRows = Val(mapinfo.eval("tableinfo(overlaytemp,8)"))
        If TempRows > 0 Then
           mapinfo.do "insert into NeighborLay" & SelectCell & " (col1,col2,col3,col4,col5,col6,col7) select time,lon,lat,ci_serv,bcch_n" & Format(i) & ",bsic_n" & Format(i) & ",rxlev_n" & Format(i) & " from overlaytemp"
        End If
    Next
    mapinfo.do "commit table NeighborLay" & SelectCell
    mapinfo.do "close table overlaytemp"

    mapinfo.do "add map auto layer NeighborLay" & SelectCell
    MyMsg = "shade window FrontWindow() NeighborLay" & SelectCell & " With RXLEV "
    If Legend_Tog = 0 Then
       MyMsg = MyMsg + " ignore 0 ranges apply all use all Symbol (39,65280,4,""MapInfo Cartographic"",0,0) 90: 63 Symbol (39,8388736,4,""MapInfo Cartographic"",0,0) , 63: 27 Symbol (39,65280,4,""MapInfo Cartographic"",0,0) ,27: 17 Symbol (39,255,4,""MapInfo Cartographic"",0,0) ,17: 0 Symbol (39,16711680,4,""MapInfo Cartographic"",0,0) "
    Else
       MyMsg = MyMsg + " ignore 0 ranges apply all use all Symbol (39,16711680,4,""MapInfo Cartographic"",0,0)  90: 63 Symbol (39,8388736,4,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,65280,4,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,7585792,4,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,4,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,4,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,4,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,4,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,4,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,4,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,4,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,4,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,4,""MapInfo Cartographic"",0,0)"
    End If
    mapinfo.do MyMsg
    If legendid = 0 Then     'win95
       mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"     'win95
       mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
       legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
    End If     'win95
    If Legend_Tog = 0 Then
       MyMsg = " Title " + Chr(34) + "ȷ��С������ " + tblname + " NeighborLay" + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""����"",0,9,0) ascending on ranges Font (""����"",0,9,0) ""����ȫ��"" display off ,""0 �� 17 (-110��-93dBm)"" display on ,""17 �� 27 (-93��-83dBm)"" display on ,""27 �� 63 (-83��-47dBm)"" display on ,""63 ���� (����-47dBm)"" display on"
    Else
       MyMsg = " Title " + Chr(34) + "ȷ��С������ " + tblname + " NeighborLay" + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""����"",0,9,0) ascending on ranges Font (""����"",0,9,0) ""����ȫ��"" display off ,""0 �� 5 (-110��-105dBm)"" display on ,""5 �� 10 (-105��-100dBm)"" display on ,""10 �� 15 (-100��-95dBm)"" display on ,""15 �� 20 (-95��-90dBm)"" display on ,""20 �� 25 (-90��-85dBm)"" display on ,""25 �� 30 (-85��-80dBm)"" display on ,""30 �� 35 (-80��-75dBm)"" display on ,""35 �� 40 (-75��-70dBm)"" display on ,""40 �� 45 (-70��-65dBm)"" display on ,""45 �� 50 (-65��-60dBm)"" display on ,""50 �� 63 (-60��-47dBm)"" display on ,""63 ���� (����-47dBm)"" display on"
    End If
    mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on " & MyMsg

End Sub
