VERSION 5.00
Begin VB.Form frmBcchRetrieve 
   Caption         =   "频率复用规律条件选择"
   ClientHeight    =   2415
   ClientLeft      =   4215
   ClientTop       =   2895
   ClientWidth     =   3300
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBcchRetrieve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3300
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   330
      Left            =   1770
      TabIndex        =   4
      Top             =   2010
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   345
      Left            =   600
      TabIndex        =   3
      Top             =   2010
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   255
      TabIndex        =   0
      Top             =   105
      Width           =   2760
      Begin VB.CheckBox Check1 
         Caption         =   "三个频率复用"
         Height          =   405
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   435
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Caption         =   "两个频率复用"
         Height          =   405
         Index           =   1
         Left            =   720
         TabIndex        =   1
         Top             =   930
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmBcchRetrieve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error Resume Next
    If Check1(0).Value = 1 And Check1(1).Value = 1 Then
        SelBcchGroup = 1
    ElseIf Check1(0).Value = 1 Then
        SelBcchGroup = 2
    ElseIf Check1(1).Value = 1 Then
        SelBcchGroup = 3
    Else
        SelBcchGroup = 0
    End If
    Unload Me

End Sub

Private Sub Command2_Click()
    On Error Resume Next
    SelBcchGroup = 0
    Unload Me
End Sub
