VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Frmrepot 
   BackColor       =   &H80000004&
   Caption         =   "生成测试报告"
   ClientHeight    =   1380
   ClientLeft      =   825
   ClientTop       =   6975
   ClientWidth     =   4560
   Icon            =   "Frmrepot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   405
      TabIndex        =   1
      Top             =   855
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   476
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   405
      TabIndex        =   2
      Top             =   165
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Left            =   405
      TabIndex        =   0
      Top             =   435
      Width           =   90
   End
End
Attribute VB_Name = "Frmrepot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Me.Top = 6500
    Me.Left = 800

End Sub
