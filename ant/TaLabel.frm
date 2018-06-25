VERSION 5.00
Begin VB.Form TaLabel 
   Caption         =   "双网TA标注"
   ClientHeight    =   1905
   ClientLeft      =   2130
   ClientTop       =   2520
   ClientWidth     =   3135
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TaLabel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   3135
   Begin VB.CommandButton SBSCancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   1665
      TabIndex        =   4
      Top             =   1485
      Width           =   1080
   End
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   450
      TabIndex        =   3
      Top             =   1485
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      Caption         =   "标注选择"
      Height          =   1005
      Left            =   330
      TabIndex        =   0
      Top             =   225
      Width           =   2490
      Begin VB.OptionButton Option1 
         Caption         =   "横"
         Height          =   240
         Left            =   405
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   525
      End
      Begin VB.OptionButton Option2 
         Caption         =   "竖"
         Height          =   240
         Left            =   1515
         TabIndex        =   1
         Top             =   495
         Width           =   510
      End
   End
End
Attribute VB_Name = "TaLabel"
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
    Dim Layers As Integer, i As Integer
    Dim Mymsg As String
    
    On Error Resume Next
       
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
       mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,255,16777215) With ta Auto On Overlap Off Duplicates on Position Above Auto On Offset 10"
       mapinfo.do "set map redraw on"
       mapinfo.do "set map redraw off"
       mapinfo.do "Set Map Layer ""Duplicate"" Label Visibility Font (""Arial"",257,8,16711935,16777215) With ta_2 Auto On Overlap Off Duplicates on Position Below Auto On Offset 10"
       mapinfo.do "set map redraw on"
    Else
       Unload Me
       mapinfo.do "set map redraw off"
       mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,255,16777215) With ta Auto On Overlap Off Duplicates On Position Left Auto On Offset 10"
       mapinfo.do "set map redraw on"
       mapinfo.do "set map redraw off"
       mapinfo.do "Set Map Layer ""Duplicate"" Label Visibility Font (""Arial"",257,8,16711935,16777215) With ta_2 Auto On Overlap Off Duplicates On Position Right Auto On Offset 10"
       mapinfo.do "set map redraw on"
    End If

End Sub
