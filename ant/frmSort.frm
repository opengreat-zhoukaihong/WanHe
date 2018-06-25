VERSION 5.00
Begin VB.Form frmSort 
   ClientHeight    =   3105
   ClientLeft      =   4080
   ClientTop       =   3885
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   3750
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   360
      Left            =   1245
      TabIndex        =   1
      Top             =   2625
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "排序选择"
      Height          =   2235
      Left            =   345
      TabIndex        =   0
      Top             =   225
      Width           =   3030
      Begin VB.OptionButton Option1 
         Caption         =   "按更新时间排序"
         Height          =   300
         Index           =   4
         Left            =   690
         TabIndex        =   6
         Top             =   1725
         Width           =   1620
      End
      Begin VB.OptionButton Option1 
         Caption         =   "按小区类型排序"
         Height          =   300
         Index           =   3
         Left            =   690
         TabIndex        =   5
         Top             =   1365
         Width           =   1620
      End
      Begin VB.OptionButton Option1 
         Caption         =   "按 Lac 排序"
         Height          =   300
         Index           =   1
         Left            =   690
         TabIndex        =   4
         Top             =   660
         Width           =   1320
      End
      Begin VB.OptionButton Option1 
         Caption         =   "按 Ci 排序"
         Height          =   300
         Index           =   2
         Left            =   690
         TabIndex        =   3
         Top             =   1005
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         Caption         =   "按小区名排序"
         Height          =   300
         Index           =   0
         Left            =   690
         TabIndex        =   2
         Top             =   330
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Integer
    On Error Resume Next
    For i = 0 To 4
        If Option1(i).Value Then
           SortType = i
           Exit For
        End If
    Next
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Option1(SortType).Value = True
End Sub
