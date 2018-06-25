VERSION 5.00
Begin VB.Form cvChoice 
   Caption         =   "数据转换"
   ClientHeight    =   2775
   ClientLeft      =   4125
   ClientTop       =   1965
   ClientWidth     =   3660
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Convert_Choice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   1020
      TabIndex        =   4
      Top             =   1350
      Width           =   1920
      Begin VB.OptionButton Option4 
         Caption         =   "抽取点(150米/点)"
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   465
         Width           =   1770
      End
      Begin VB.OptionButton Option3 
         Caption         =   "滤除相同经纬度"
         Height          =   240
         Left            =   105
         TabIndex        =   5
         Top             =   75
         Value           =   -1  'True
         Width           =   1560
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "转换选择"
      Height          =   1995
      Left            =   360
      TabIndex        =   1
      Top             =   180
      Width           =   2910
      Begin VB.OptionButton Option2 
         Caption         =   "地理点滤除处理"
         Height          =   285
         Left            =   525
         TabIndex        =   3
         Top             =   780
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton Option1 
         Caption         =   "地理点平滑处理"
         Height          =   240
         Left            =   525
         TabIndex        =   2
         Top             =   405
         Width           =   1560
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   320
      Left            =   1305
      TabIndex        =   0
      Top             =   2385
      Width           =   1080
   End
End
Attribute VB_Name = "cvChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error Resume Next
    If Option1.Value = True Then
       tran_del = 1
    Else
       If Option3.Value Then
          tran_del = 2
       Else
          tran_del = 3
       End If
    End If
    Unload Me
    DocManager.Show 1
End Sub

Private Sub Option1_Click()
    On Error Resume Next
    If Option2.Value Then
       Option3.Enabled = True
       Option4.Enabled = True
    Else
       Option3.Enabled = False
       Option4.Enabled = False
    End If

End Sub

Private Sub Option2_Click()
    On Error Resume Next
    If Option2.Value Then
       Option3.Enabled = True
       Option4.Enabled = True
    Else
       Option3.Enabled = False
       Option4.Enabled = False
    End If
End Sub
