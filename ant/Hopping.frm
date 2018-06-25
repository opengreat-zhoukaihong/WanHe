VERSION 5.00
Begin VB.Form HoppingFrm 
   Caption         =   "标注跳频参数"
   ClientHeight    =   2280
   ClientLeft      =   2850
   ClientTop       =   2145
   ClientWidth     =   3315
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Hopping.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   510
      TabIndex        =   7
      Top             =   1890
      Width           =   1080
   End
   Begin VB.CommandButton SBSCancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   1725
      TabIndex        =   6
      Top             =   1890
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      Caption         =   "标注选择"
      Height          =   1440
      Left            =   1725
      TabIndex        =   1
      Top             =   195
      Width           =   1290
      Begin VB.OptionButton Option2 
         Caption         =   "竖"
         Enabled         =   0   'False
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   945
         Width           =   510
      End
      Begin VB.OptionButton Option1 
         Caption         =   "横"
         Enabled         =   0   'False
         Height          =   240
         Left            =   360
         TabIndex        =   2
         Top             =   495
         Value           =   -1  'True
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "参数选择"
      Height          =   1440
      Left            =   300
      TabIndex        =   0
      Top             =   195
      Width           =   1290
      Begin VB.CheckBox Check2 
         Caption         =   "HSN"
         Height          =   240
         Left            =   345
         TabIndex        =   5
         Top             =   945
         Width           =   675
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Maio"
         Height          =   240
         Left            =   345
         TabIndex        =   4
         Top             =   480
         Value           =   1  'Checked
         Width           =   705
      End
   End
End
Attribute VB_Name = "HoppingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    On Error Resume Next
    If Check1.Value = 0 Or Check2.Value = 0 Then
       Option1.Enabled = False
       Option2.Enabled = False
    End If
    If Check1.Value = 1 And Check2.Value = 1 Then
       Option1.Enabled = True
       Option2.Enabled = True
    End If

End Sub

Private Sub Check2_Click()
    On Error Resume Next
    If Check1.Value = 0 Or Check2.Value = 0 Then
       Option1.Enabled = False
       Option2.Enabled = False
    End If
    If Check1.Value = 1 And Check2.Value = 1 Then
       Option1.Enabled = True
       Option2.Enabled = True
    End If
    
End Sub

Private Sub SBSCancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub SBSOK_Click()
    Dim Layers As Integer, i As Integer
    Dim Mymsg As String
    
    On Error Resume Next
    
    If Check1.Value = 1 And Check2.Value = 1 Then
       mapinfo.do "select * from " & tblname & " into Duplicate"
       mapinfo.do "Add Map window FrontWindow() Layer Duplicate"
       
       Layers = mapinfo.eval("mapperinfo(frontwindow(),9)")
       Mymsg = "set map order "
       For i = 2 To Layers
           Mymsg = Mymsg + Format(i) + ","
       Next
       Mymsg = Mymsg + "1"
       mapinfo.do Mymsg
       
       If Option2.Value = True Then
          Unload Me
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,255,16777215) With maio_dch Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
          mapinfo.do "set map redraw on"
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer ""Duplicate"" Label Visibility Font (""Arial"",257,8,16711935,16777215) With hsn_dch_ Auto On Overlap Off Duplicates On Position Below Auto On Offset 10"
          mapinfo.do "set map redraw on"
       Else
          Unload Me
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,255,16777215) With maio_dch Auto On Overlap Off Duplicates On Position Left Auto On Offset 10"
          mapinfo.do "set map redraw on"
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer ""Duplicate"" Label Visibility Font (""Arial"",257,8,16711935,16777215) With hsn_dch_ Auto On Overlap Off Duplicates On Position Right Auto On Offset 10"
          mapinfo.do "set map redraw on"
       End If
    Else
       If Check1.Value = 1 Then
          Unload Me
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,255,16777215) With maio_dch Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
          mapinfo.do "set map redraw on"
       Else
          Unload Me
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,16711935,16777215) With hsn_dch_ Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
          mapinfo.do "set map redraw on"
       End If
    End If
    
End Sub
