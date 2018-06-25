VERSION 5.00
Begin VB.Form Tch_mmap_choice 
   BackColor       =   &H00C0C0C0&
   Caption         =   "TCH ��ͼ��ʾѡ��"
   ClientHeight    =   4155
   ClientLeft      =   2805
   ClientTop       =   2670
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tch_mmap.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4155
   ScaleWidth      =   4650
   Begin VB.CheckBox Check2 
      Caption         =   "�������֮һ"
      Height          =   240
      Left            =   2220
      TabIndex        =   11
      Top             =   225
      Width           =   1380
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ÿ�߻�����"
      Height          =   240
      Left            =   570
      TabIndex        =   10
      Top             =   210
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ѡ��"
      Height          =   2970
      Left            =   210
      TabIndex        =   2
      Top             =   585
      Width           =   4230
      Begin VB.OptionButton Option7 
         Caption         =   "�л�,�л�������,�л�ʧ����,�л��ɹ���"
         Enabled         =   0   'False
         Height          =   240
         Left            =   330
         TabIndex        =   9
         Top             =   2610
         Width           =   3630
      End
      Begin VB.OptionButton Option6 
         Caption         =   "�����ؽ���,���н�����"
         Enabled         =   0   'False
         Height          =   240
         Left            =   330
         TabIndex        =   8
         Top             =   2235
         Width           =   2190
      End
      Begin VB.OptionButton Option5 
         Caption         =   "ռ��ʱ��,����������,������Ӧ��"
         Enabled         =   0   'False
         Height          =   240
         Left            =   330
         TabIndex        =   7
         Top             =   1860
         Width           =   3000
      End
      Begin VB.OptionButton Option4 
         Caption         =   "������,��ͨ����"
         Enabled         =   0   'False
         Height          =   240
         Left            =   330
         TabIndex        =   6
         Top             =   1515
         Width           =   1650
      End
      Begin VB.OptionButton Option3 
         Caption         =   "�����ŵ���,�ŵ���,�����"
         Enabled         =   0   'False
         Height          =   240
         Left            =   330
         TabIndex        =   5
         Top             =   1140
         Width           =   2475
      End
      Begin VB.OptionButton Option2 
         Caption         =   "������,�������,�������"
         Enabled         =   0   'False
         Height          =   240
         Left            =   330
         TabIndex        =   4
         Top             =   780
         Width           =   2460
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ӵ����,������,���гɹ���"
         Enabled         =   0   'False
         Height          =   240
         Left            =   330
         TabIndex        =   3
         Top             =   405
         Value           =   -1  'True
         Width           =   2460
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   320
      Left            =   1080
      TabIndex        =   1
      Top             =   3765
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   320
      Left            =   2325
      TabIndex        =   0
      Top             =   3765
      Width           =   1080
   End
End
Attribute VB_Name = "Tch_mmap_choice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check2_Click()
    On Error Resume Next
    If Check2.Value = 1 Then
       Option1.Enabled = True
       Option2.Enabled = True
       Option3.Enabled = True
       Option4.Enabled = True
       Option5.Enabled = True
       Option6.Enabled = True
       Option7.Enabled = True
    Else
       Option1.Enabled = False
       Option2.Enabled = False
       Option3.Enabled = False
       Option4.Enabled = False
       Option5.Enabled = False
       Option6.Enabled = False
       Option7.Enabled = False
    End If

End Sub

Private Sub Command1_Click()
    Dim center_point, center_lon, center_lat
    Dim i As Integer
    Dim WinId
    
    On Error Resume Next
    If Check1.Value = 0 And Check2.Value = 0 Then
       Unload Me
       Exit Sub
    End If
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\tch_sts.tab" + Chr(34)
    mapinfo.do "set next document parent " & MapForm.hwnd & "style 1"
    TableNum = Val(mapinfo.eval("NumTables()"))
    If TableNum > 1 Then
       msg = "Add Map Auto Layer" + Chr(34) + "tch_sts" + Chr(34)
       mapinfo.do msg
       msg = Chr(34) + "km" + Chr(34)
       mapinfo.do "set map zoom 6 units " & msg
    Else
       msg = "Map from " + Chr(34) + "tch_sts" + Chr(34)
       mapinfo.do msg
       thereIsAMap = True
       mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    MapForm.Caption = MapForm.Caption + ",TCH"
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
           WinId = mapinfo.eval("windowid(" & i & ")")
           If WinId = mapinfo.eval("frontwindow()") Then
              Exit For
           End If
        End If
    Next
    If Check1.Value = 1 Then
       mapinfo.do "Set Style Pen MakePen(1,60,0)"
       mapinfo.do "set style brush  makebrush(2,7585792,7585792) "
       mapinfo.do "shade window " + WinId + " tch_sts with col7 pie Angle 180 Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " At Value 1 vary size by " + Chr(34) + "SQRT" + Chr(34) + " border Pen (1,2,8245248)  position center center style Brush (2,8245248,16777215)  # max 1 color 0 #"
       mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " TCH��״ͼ" + Chr(34) + " Font(""����"",0,9,0) ascending on ranges Font(""����"",0,9,0) """" display off ," + Chr(34) + "ÿ�߻�����(erl)" + Chr(34) + " display on"
    End If
    If Check2.Value = 1 Then
       If Option1.Value = True Then
          mapinfo.do "shade window " + WinId + " tch_sts with col16,col19,col17 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.505 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215) ,Brush (2,255,16777215)  # max 100 color 0 #"
          mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " TCHֱ��ͼ" + Chr(34) + " Font(""����"",0,9,0) ascending on ranges Font(""����"",0,9,0) """" display off ," + Chr(34) + "ӵ���� %" + Chr(34) + " display on ," + Chr(34) + "������ %" + Chr(34) + " display on ," + Chr(34) + "���гɹ��� %" + Chr(34) + " display on"
       End If
       If Option2.Value = True Then
          mapinfo.do "shade window " + WinId + " tch_sts with col18,col13,col14 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.505 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215) ,Brush (2,255,16777215)  # max 100 color 0 #"
          mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " TCHֱ��ͼ" + Chr(34) + " Font(""����"",0,9,0) ascending on ranges Font(""����"",0,9,0) """" display off ," + Chr(34) + "������" + Chr(34) + " display on ," + Chr(34) + "�������" + Chr(34) + " display on ," + Chr(34) + "�������" + Chr(34) + " display on"
       End If
       If Option3.Value = True Then
          mapinfo.do "shade window " + WinId + " tch_sts with col4,col3,col5 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.505 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215) ,Brush (2,255,16777215)  # max 100 color 0 #"
          mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " TCHֱ��ͼ" + Chr(34) + " Font(""����"",0,9,0) ascending on ranges Font(""����"",0,9,0) """" display off ," + Chr(34) + "�����ŵ���" + Chr(34) + " display on ," + Chr(34) + "�ŵ���" + Chr(34) + " display on ," + Chr(34) + "����� %" + Chr(34) + " display on"
       End If
       If Option4.Value = True Then
          mapinfo.do "shade window " + WinId + " tch_sts with col6,col8 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.340 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215) # max 100 color 0 #"
          mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " TCHֱ��ͼ" + Chr(34) + " Font(""����"",0,9,0) ascending on ranges Font(""����"",0,9,0) """" display off ," + Chr(34) + "������" + Chr(34) + " display on ," + Chr(34) + "��ͨ����" + Chr(34) + " display on "
       End If
       If Option5.Value = True Then
          mapinfo.do "shade window " + WinId + " tch_sts with col9,col10,col11 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.505 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215) ,Brush (2,255,16777215)  # max 100 color 0 #"
          mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " TCHֱ��ͼ" + Chr(34) + " Font(""����"",0,9,0) ascending on ranges Font(""����"",0,9,0) """" display off ," + Chr(34) + "ռ��ʱ��" + Chr(34) + " display on ," + Chr(34) + "����������" + Chr(34) + " display on ," + Chr(34) + "������Ӧ��" + Chr(34) + " display on"
       End If
       If Option6.Value = True Then
          mapinfo.do "shade window " + WinId + " tch_sts with col12,col15 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.340 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215)  # max 100 color 0 #"
          mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " TCHֱ��ͼ" + Chr(34) + " Font(""����"",0,9,0) ascending on ranges Font(""����"",0,9,0) """" display off ," + Chr(34) + "�����ؽ���" + Chr(34) + " display on ," + Chr(34) + "���н�����" + Chr(34) + " display on "
       End If
       If Option7.Value = True Then
          mapinfo.do "shade window " + WinId + " tch_sts with col20,col21,col22,col23 bar normalized Max Size 0.635 Units " + Chr(34) + "cm" + Chr(34) + " border Pen (1,2,0) Width 0.685 Units " + Chr(34) + "cm" + Chr(34) + " position center above style Brush (2,16711680,16777215) ,Brush (2,15790080,16777215) ,Brush (2,255,16777215) ,Brush (2,16711935,16777215) # max 100 color 0 #"
          mapinfo.do "set legend window " + WinId + " layer prev display on shades on symbols off lines off count off title " + Chr(34) + " TCHֱ��ͼ" + Chr(34) + " Font(""����"",0,9,0) ascending on ranges Font(""����"",0,9,0) """" display off ," + Chr(34) + "�л�" + Chr(34) + " display on ," + Chr(34) + "�л�������" + Chr(34) + " display on ," + Chr(34) + "�л�ʧ���� %" + Chr(34) + " display on," + Chr(34) + "�л��ɹ��� %" + Chr(34) + " display on"
       End If
       
    End If
    thereIsAMap = True
    If mapid = 0 Then
       mapid = Val(mapinfo.eval("FrontWindow()"))
    End If
    MDIMain.SUB_23.Enabled = 1
    MDIMain.SUB_24.Enabled = 1
    MDIMain.SUB_25.Enabled = 1
    MDIMain.SUB_26.Enabled = 1
    MDIMain.SUB_28.Enabled = 1
    
    center_point = mapinfo.eval("tableinfo(tch_sts,8)")
    mapinfo.do "fetch first from tch_sts"
    For i = 1 To center_point
       center_lon = mapinfo.eval("tch_sts.lon")
       center_lat = mapinfo.eval("tch_sts.lat")
       If center_lon <> 0 And center_lat <> 0 Then
          Exit For
       Else
          mapinfo.do "fetch next from tch_sts"
       End If
    Next
    mapinfo.do "set map Center(" & center_lon & "," & center_lat & ") "
    mapinfo.runmenucommand 610
    
    Unload Me

End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub
