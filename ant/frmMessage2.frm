VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMessage2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "邻频干扰查找"
   ClientHeight    =   2730
   ClientLeft      =   7785
   ClientTop       =   6060
   ClientWidth     =   3420
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMessage2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ListView ListView1 
      Height          =   1860
      Left            =   105
      TabIndex        =   4
      Top             =   765
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   3281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "干扰小区中文名"
         Object.Width           =   2751
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ARFCN"
         Object.Width           =   599
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "BSIC"
         Object.Width           =   546
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   180
      Index           =   4
      Left            =   2595
      TabIndex        =   3
      Top             =   465
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   180
      Index           =   3
      Left            =   1770
      TabIndex        =   2
      Top             =   450
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   180
      Index           =   2
      Left            =   165
      TabIndex        =   1
      Top             =   450
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "主小区中文名      ARFCN    BSIC"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   2790
   End
End
Attribute VB_Name = "frmMessage2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    MessageId2 = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    MessageId2 = 0

End Sub

