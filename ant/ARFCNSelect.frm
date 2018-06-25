VERSION 5.00
Begin VB.Form ARFCNSelect 
   Caption         =   "频率分析"
   ClientHeight    =   2595
   ClientLeft      =   3540
   ClientTop       =   3375
   ClientWidth     =   4020
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ARFCNSelect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   4020
   Begin VB.CheckBox Check2 
      Caption         =   "DCS"
      Height          =   300
      Left            =   1350
      TabIndex        =   9
      Top             =   1755
      Value           =   1  'Checked
      Width           =   600
   End
   Begin VB.CheckBox Check1 
      Caption         =   "GSM"
      Height          =   300
      Left            =   450
      TabIndex        =   8
      Top             =   1755
      Value           =   1  'Checked
      Width           =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   320
      Left            =   2040
      TabIndex        =   7
      Top             =   2205
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   320
      Left            =   840
      TabIndex        =   6
      Top             =   2205
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      Caption         =   "分析手段"
      Height          =   1410
      Left            =   2280
      TabIndex        =   3
      Top             =   180
      Width           =   1515
      Begin VB.OptionButton Option4 
         Caption         =   "标注"
         Height          =   300
         Left            =   375
         TabIndex        =   5
         Top             =   840
         Width           =   750
      End
      Begin VB.OptionButton Option3 
         Caption         =   "专题图"
         Height          =   300
         Left            =   360
         TabIndex        =   4
         Top             =   435
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "参数选择"
      Height          =   1410
      Left            =   255
      TabIndex        =   0
      Top             =   180
      Width           =   1890
      Begin VB.OptionButton Option2 
         Caption         =   "SDCCH/TCH"
         Height          =   300
         Left            =   390
         TabIndex        =   2
         Top             =   855
         Width           =   1125
      End
      Begin VB.OptionButton Option1 
         Caption         =   "BCCH"
         Height          =   300
         Left            =   375
         TabIndex        =   1
         Top             =   405
         Value           =   -1  'True
         Width           =   705
      End
   End
End
Attribute VB_Name = "ARFCNSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error Resume Next
    If Option1.Value = True Then
       If Option3.Value = True Then
          If Menu_Flag = 88388 Then
             Menu_Flag = 83131
          Else
             Menu_Flag = 3131
          End If
       Else
          If Menu_Flag = 88388 Then
             Menu_Flag = 83133
          Else
             Menu_Flag = 3133
          End If
       End If
    Else
       If Option3.Value = True Then
          Menu_Flag = 3132
       Else
          Menu_Flag = 3134
       End If
    End If
    If Check1.Value = 1 And Check2.Value = 1 Then
       GSMDCSBCCH = 0
    Else
       If Check1.Value = 1 Then
          GSMDCSBCCH = 1
       ElseIf Check2.Value = 1 Then
          GSMDCSBCCH = 2
       Else
          GSMDCSBCCH = 0
       End If
    End If
    Unload Me
    SelTable.Show 1
    
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    If Menu_Flag = 88388 Then
        Option2.Enabled = False
    End If
End Sub
