VERSION 5.00
Begin VB.Form frmForecast 
   Caption         =   "ѡƵ��ǿ�ֲ�"
   ClientHeight    =   2745
   ClientLeft      =   1305
   ClientTop       =   1140
   ClientWidth     =   3390
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmForecast.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   3390
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OK 
      Caption         =   "&O ȷ��"
      Height          =   320
      Left            =   2145
      TabIndex        =   2
      Top             =   435
      Width           =   1080
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&C ȡ��"
      Height          =   320
      Left            =   2145
      TabIndex        =   1
      Top             =   825
      Width           =   1080
   End
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   210
      TabIndex        =   0
      Top             =   390
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ѡƵ�ʣ�"
      Height          =   180
      Left            =   210
      TabIndex        =   3
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmForecast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ScanArfcn() As String
Dim ScanFileName() As String

Private Sub Cancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Mymsg As String
    Dim col_num As Integer, j As Integer
    Dim arfcn_field As String, cover_arfcn As String
    
    On Error Resume Next
    Mymsg = "TableInfo(""" & tblname & """, 4)"
    col_num = Val(mapinfo.eval(Mymsg))
    For j = 4 To col_num Step 2
        arfcn_field = mapinfo.eval("Columninfo(""" & tblname & """,""COL" & j & """, 1)")
        cover_arfcn = Right(arfcn_field, Len(arfcn_field) - 6)
        List1.AddItem cover_arfcn
    Next

End Sub

Private Sub OK_Click()
    Dim i As Integer, j As Integer
    Dim WinId As Long
    
    On Error Resume Next
    If List1.ListIndex > -1 Then
                  msg = " shade window FrontWindow() " + tblname + " With arfcn_" & Trim(List1.List(List1.ListIndex))
                  If Legend_Tog = 0 Then
                     'msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                     'msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 0: 75 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 75: 85 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,85: 95 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,95: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                     msg = msg + " ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 75: 0 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 85: 75 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,95: 85 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,110: 95 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
                  Else
                     'msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 120: 63 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,63: 50 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,50: 45 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,45: 40 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,40: 35 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,35: 30 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,30: 25 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,25: 20 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,20: 15 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,15: 10 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,10: 5 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,5: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                     'msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 0: 47 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,47: 60 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,60: 65 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,65: 70 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,70: 75 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,75: 80 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,80: 85 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,85: 90 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,90: 95 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,95: 100 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,100: 105 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,105:110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                     msg = msg + " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) 47: 0 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,60: 47 Symbol (39,7585792,8,""MapInfo Cartographic"",0,0) ,65: 60 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,70: 65 Symbol (39,16750640,8,""MapInfo Cartographic"",0,0) ,75: 70 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,80: 75 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,85: 80 Symbol (39,8432639,8,""MapInfo Cartographic"",0,0) ,90: 85 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,95: 90 Symbol (39,9584,8,""MapInfo Cartographic"",0,0) ,100: 95 Symbol (39,16744576,8,""MapInfo Cartographic"",0,0) ,105: 100 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,110:105 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
                  End If
                  mapinfo.do msg
                  
                  For i = 1 To mapinfo.eval("NumWindows()")     'win95
                      If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then     'win95
                         WinId = mapinfo.eval("windowid(" & i & ")")     'win95
                         If WinId = mapinfo.eval("frontwindow()") Then
                            Exit For
                         End If
                      End If     'win95
                  Next     'win95

                  If legendid = 0 Then     'win95
                      mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"     'win95
                      mapinfo.do "Create Legend From Window  Frontwindow()"     'win95
                      legendid = mapinfo.eval("windowinfo(1009,12)")     'win95
                  End If     'win95
                  If Legend_Tog = 0 Then
                         'msg = " Title " + Chr(34) + "��ǿ�ֲ��۲� " + tblname + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + "��ѡƵ�ʣ�" & Trim(List1.List(List1.ListIndex)) + Chr(34) + " Font (""����"",0,9,0) ascending on ranges Font (""����"",0,9,0) ""����ȫ��"" display off ,""0 �� 15 (-110��-95dBm)"" display on ,""15 �� 25 (-95��-85dBm)"" display on ,""25 �� 35 (-85��-75dBm)"" display on ,""35 ���� (����-75dBm)"" display on"
                         'msg = " Title " + Chr(34) + "��ǿ�ֲ��۲� " + tblname + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + "��ѡƵ�ʣ�" & Trim(List1.List(List1.ListIndex)) + Chr(34) + " Font (""����"",0,9,0) ascending on ranges Font (""����"",0,9,0) ""����ȫ��"" display on ,""0 �� 15 (-110��-95dBm)"" display on ,""15 �� 25 (-95��-85dBm)"" display on ,""25 �� 35 (-85��-75dBm)"" display on ,""35 ���� (����-75dBm)"" display on"
                         msg = " Title " + Chr(34) + "��ǿ�ֲ��۲� " + tblname + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + "��ѡƵ�ʣ�" & Trim(List1.List(List1.ListIndex)) + Chr(34) + " Font (""����"",0,9,255) ascending off ranges Font (""����"",0,9,0) ""����ȫ��"" display off ,""35 ���� (����-75dBm)"" display on ,""25 �� 35 (-85��-75dBm)"" display on ,""15 �� 25 (-95��-85dBm)"" display on ,""0 �� 15 (-110��-95dBm)"" display on "
                         
                  Else
                         'msg = " Title " + Chr(34) + "��ǿ�ֲ��۲� " + tblname + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + "��ѡƵ�ʣ�" & Trim(List1.List(List1.ListIndex)) + Chr(34) + " Font (""����"",0,9,0) ascending on ranges Font (""����"",0,9,0) ""����ȫ��"" display off ,""0 �� 5 (-110��-105dBm)"" display on ,""5 �� 10 (-105��-100dBm)"" display on ,""10 �� 15 (-100��-95dBm)"" display on ,""15 �� 20 (-95��-90dBm)"" display on ,""20 �� 25 (-90��-85dBm)"" display on ,""25 �� 30 (-85��-80dBm)"" display on ,""30 �� 35 (-80��-75dBm)"" display on ,""35 �� 40 (-75��-70dBm)"" display on ,""40 �� 45 (-70��-65dBm)"" display on ,""45 �� 50 (-65��-60dBm)"" display on ,""50 �� 63 (-60��-47dBm)"" display on ,""63 ���� (����-47dBm)"" display on"
                         msg = " Title " + Chr(34) + "��ǿ�ֲ��۲� " + tblname + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + "��ѡƵ�ʣ�" & Trim(List1.List(List1.ListIndex)) + Chr(34) + " Font (""����"",0,9,255) ascending off ranges Font (""����"",0,9,0) ""����ȫ��"" display off ,""63 ���� (����-47dBm)"" display on,""50 �� 63 (-60��-47dBm)"" display on ,""45 �� 50 (-65��-60dBm)"" display on ,""40 �� 45 (-70��-65dBm)"" display on ,""35 �� 40 (-75��-70dBm)"" display on ,""30 �� 35 (-80��-75dBm)"" display on ,""25 �� 30 (-85��-80dBm)"" display on ,""20 �� 25 (-90��-85dBm)"" display on ,""15 �� 20 (-95��-90dBm)"" display on ,""10 �� 15 (-100��-95dBm)"" display on ,""5 �� 10 (-105��-100dBm)"" display on ,""0 �� 5 (-110��-105dBm)"" display on "
                  End If
                  mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on " & msg
                  
                  mapinfo.do "set map redraw off"
                  mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With """ & Trim(List1.List(List1.ListIndex)) & """ Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
                  mapinfo.do "set map redraw on"
    
    
       Unload Me
    Else
       MsgBox "��ѡ��һ��Ƶ��", 64, "��ʾ"
    End If
End Sub
