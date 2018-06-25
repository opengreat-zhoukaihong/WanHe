VERSION 5.00
Begin VB.Form frmDelete 
   Caption         =   "删除基站选择"
   ClientHeight    =   2595
   ClientLeft      =   3165
   ClientTop       =   2250
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   3720
   Begin VB.Frame Frame1 
      Height          =   1995
      Left            =   315
      TabIndex        =   2
      Top             =   15
      Width           =   2940
      Begin VB.CheckBox Check1 
         Caption         =   "删除第三小区"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   720
         TabIndex        =   6
         Top             =   1410
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Caption         =   "删除第二小区"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   720
         TabIndex        =   5
         Top             =   1035
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Caption         =   "删除第一小区"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   645
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Caption         =   "删除整个基站"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   270
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   660
      TabIndex        =   1
      Top             =   2175
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1830
      TabIndex        =   0
      Top             =   2175
      Width           =   1065
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click(Index As Integer)
    On Error Resume Next
    If Check1(Index).Value = 1 Then
       If Index = 0 Then
          Check1(1).Value = 0
          Check1(2).Value = 0
          Check1(3).Value = 0
       Else
          Check1(0).Value = 0
       End If
    End If
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    On Error Resume Next
    If Check1(0).Value = 1 Then
       CheckValue(0) = False
       CheckValue(1) = False
       CheckValue(2) = False
    Else
       For i = 0 To 2
           If Check1(i + 1).Value = 1 Then
              CheckValue(i) = False
           End If
       Next
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    On Error Resume Next
    Check1(0).Value = 1
    For i = 0 To 2
        If Not CheckValue(i) Then
           Check1(i + 1).Enabled = False
        End If
    Next
End Sub
