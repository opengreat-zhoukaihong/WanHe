VERSION 5.00
Begin VB.Form frmBcchAdjust 
   Caption         =   "��Ƶ����ͳ��"
   ClientHeight    =   2475
   ClientLeft      =   4125
   ClientTop       =   3000
   ClientWidth     =   3450
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBcchAdjust.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3450
   Begin VB.CommandButton SBSOK 
      Caption         =   "ȷ��"
      Height          =   320
      Left            =   585
      TabIndex        =   4
      Top             =   2040
      Width           =   1080
   End
   Begin VB.CommandButton SBSCancel 
      Caption         =   "ȡ��"
      Height          =   320
      Left            =   1740
      TabIndex        =   3
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Height          =   1650
      Left            =   270
      TabIndex        =   0
      Top             =   135
      Width           =   2910
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   2100
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "-9"
         Top             =   795
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "��Ƶ��ײRxLev��ֵ��"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   825
         Width           =   1710
      End
   End
End
Attribute VB_Name = "frmBcchAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SBSCancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub SBSOK_Click()
    Dim MyRows As Long
    Dim DoNotExist As Boolean
    
    On Error Resume Next
    Me.Hide
    mapinfo.do "select * from " & tblname & " where rxle_neig2 >0 into Mytemp"
    MyRows = mapinfo.eval("tableinfo(mytemp,8)")
    mapinfo.do "close table mytemp"
    If MyRows = 0 Then
        MsgBox "���ļ���������Ƶ����ͳ��", 64, "��ʾ"
        Exit Sub
    End If
    'mapinfo.do "select * from " & tblname & " where rxle_neig2-rxlev_s>" & Format(Abs(Val(Text1.Text))) & " or rxle_neig3-rxlev_s>" & Format(Abs(Val(Text1.Text))) & " into BcchAdjust"
    'mapinfo.do "select * from " & tblname & " where rxle_neig2-rxlev_s>" & Format(Abs(Val(Text1.Text))) & " into BcchAdjust1"
    mapinfo.do "select * from " & tblname & " where rxlev_s-rxle_neig2<" & Text1.Text & " into BcchAdjust1"
    MyRows = mapinfo.eval("tableinfo(BcchAdjust1,8)")
    If MyRows = 0 Then
        DoNotExist = True
        mapinfo.do "close table BcchAdjust1"
        GoTo NextOne
    End If
    mapinfo.do "Add Map window Frontwindow() Layer BcchAdjust1"
    mapinfo.do "shade window FrontWindow() BcchAdjust1 With RXLEV_s ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
    mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "��Ƶ����ͳ����Ƶ��ײ��ʾ��-1�� " + tblname + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + "��λ��RXLEV  ��ע��BCCH����ɫ��" + Chr(34) + " Font (""����"",0,9,255) ascending on ranges Font (""����"",0,9,0) ""����ȫ��"" display off ,""0 �� 15 (-110��-95dBm)"" display on ,""15 �� 25 (-95��-85dBm)"" display on ,""25 �� 35 (-85��-75dBm)"" display on ,""35 ���� (����-75dBm)"" display on"
    
    mapinfo.do "set map redraw off"
    mapinfo.do "Set Map Layer BcchAdjust1 Label Visibility Font (""Arial"",257,8,16711680,16777215) With rxle_neig2-rxlev_s Auto On Overlap Off Duplicates On Position above Auto On Offset 2"
    mapinfo.do "set map redraw on"
    
NextOne:
    mapinfo.do "select * from " & tblname & " where rxlev_s-rxle_neig3<" & Text1.Text & " into BcchAdjust2"
    MyRows = mapinfo.eval("tableinfo(BcchAdjust2,8)")
    If MyRows = 0 Then
        If DoNotExist Then
            MsgBox "��������ƵƵ����ײ", 64, "��ʾ"
            mapinfo.do "close table BcchAdjust2"
            Unload Me
            Exit Sub
        End If
    End If
    mapinfo.do "Add Map window Frontwindow() Layer BcchAdjust2"
    mapinfo.do "shade window FrontWindow() BcchAdjust2 With RXLEV_s ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
    mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "��Ƶ����ͳ����Ƶ��ײ��ʾ��+1�� " + tblname + Chr(34) + " Font (""����"",0,9,0) Subtitle" + Chr(34) + "��λ��RXLEV  ��ע��BCCH���ӻ�ɫ��" + Chr(34) + " Font (""����"",0,9,255) ascending on ranges Font (""����"",0,9,0) ""����ȫ��"" display off ,""0 �� 15 (-110��-95dBm)"" display on ,""15 �� 25 (-95��-85dBm)"" display on ,""25 �� 35 (-85��-75dBm)"" display on ,""35 ���� (����-75dBm)"" display on"
    
    mapinfo.do "set map redraw off"
    mapinfo.do "Set Map Layer BcchAdjust2 Label Visibility Font (""Arial"",257,8,8421376,16777215) With rxle_neig3-rxlev_s Auto On Overlap Off Duplicates On Position Below Auto On Offset 2"
    mapinfo.do "set map redraw on"
         
    Unload Me
End Sub
