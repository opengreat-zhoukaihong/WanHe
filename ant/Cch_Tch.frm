VERSION 5.00
Begin VB.Form CchTch_Frm 
   Caption         =   "同频邻频观测"
   ClientHeight    =   2985
   ClientLeft      =   3990
   ClientTop       =   2415
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
   Icon            =   "Cch_Tch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3315
   Begin VB.Frame Frame1 
      Caption         =   "观测距离"
      Height          =   1320
      Index           =   1
      Left            =   225
      TabIndex        =   5
      Top             =   1080
      Width           =   2835
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Cch_Tch.frx":030A
         Left            =   1455
         List            =   "Cch_Tch.frx":0314
         TabIndex        =   8
         Text            =   "5"
         Top             =   405
         Width           =   720
      End
      Begin VB.OptionButton Option2 
         Caption         =   "全网"
         Height          =   240
         Index           =   2
         Left            =   255
         TabIndex        =   7
         Top             =   825
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "距离限制"
         Height          =   240
         Index           =   1
         Left            =   255
         TabIndex        =   6
         Top             =   450
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公里"
         Height          =   180
         Left            =   2250
         TabIndex        =   9
         Top             =   465
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "信道类型选择"
      Height          =   900
      Index           =   0
      Left            =   225
      TabIndex        =   2
      Top             =   120
      Width           =   2835
      Begin VB.OptionButton Option2 
         Caption         =   "TCH"
         Height          =   240
         Index           =   0
         Left            =   1725
         TabIndex        =   4
         Top             =   420
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         Caption         =   "BCCH"
         Height          =   240
         Index           =   0
         Left            =   450
         TabIndex        =   3
         Top             =   420
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "&C 取消"
      Height          =   300
      Left            =   1725
      TabIndex        =   1
      Top             =   2595
      Width           =   1080
   End
   Begin VB.CommandButton OK 
      Caption         =   "&O 确认"
      Height          =   300
      Left            =   525
      TabIndex        =   0
      Top             =   2595
      Width           =   1080
   End
End
Attribute VB_Name = "CchTch_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
    On Error Resume Next
    SearchDistance = 19999
    Unload Me
End Sub

Private Sub OK_Click()
    On Error Resume Next
    If Option1(0).Value = True Then
       CELL_CCH = 1
    Else
       CELL_CCH = 0
    End If
    If Option2(2).Value = True Then
       SearchDistance = 0
    Else
       If Val(Combo1.Text) = 0 Then
          SearchDistance = 5
       Else
          SearchDistance = Val(Combo1.Text)
       End If
    End If
    Unload Me
    
End Sub

Private Sub Option2_Click(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        Combo1.Enabled = True
    Else
        Combo1.Enabled = False
    End If
End Sub
