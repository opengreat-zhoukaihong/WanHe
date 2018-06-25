VERSION 5.00
Begin VB.Form cv_Choice 
   Caption         =   "数据转换"
   ClientHeight    =   2385
   ClientLeft      =   4125
   ClientTop       =   1965
   ClientWidth     =   3450
   BeginProperty Font 
      Name            =   "System"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   3450
   Begin VB.Frame Frame1 
      Caption         =   "转换选择"
      Height          =   1500
      Left            =   345
      TabIndex        =   1
      Top             =   150
      Width           =   2805
      Begin VB.OptionButton Option2 
         Caption         =   "滤除相同经纬度"
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   990
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         Caption         =   "地理点平滑处理"
         Height          =   240
         Left            =   360
         TabIndex        =   2
         Top             =   540
         Width           =   2190
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确 定"
      Height          =   450
      Left            =   1170
      TabIndex        =   0
      Top             =   1845
      Width           =   1215
   End
End
Attribute VB_Name = "cv_choice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error Resume Next
    If Option1.Value = True Then
       tran_del = False
    Else
       tran_del = True
    End If
    Unload Me
    DocManager.Show 1
End Sub
