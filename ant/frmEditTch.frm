VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmUpgradeCell 
   BorderStyle     =   0  'None
   Caption         =   "小区库存升级"
   ClientHeight    =   3270
   ClientLeft      =   2115
   ClientTop       =   1305
   ClientWidth     =   4590
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3250.877
   ScaleMode       =   0  'User
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   2940
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   9798
            MinWidth        =   9798
            Text            =   "本次操作将对所选目录下的所有小区库进行升级"
            TextSave        =   "本次操作将对所选目录下的所有小区库进行升级"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   225
      Width           =   2790
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   3
      Top             =   2550
      Width           =   2805
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   345
      Left            =   3255
      TabIndex        =   2
      Top             =   1020
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   345
      Left            =   3255
      TabIndex        =   1
      Top             =   600
      Width           =   1065
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   270
      TabIndex        =   0
      Top             =   585
      Width           =   2790
   End
End
Attribute VB_Name = "frmUpgradeCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTemp As String

Private Sub Form_Load()
    On Error Resume Next
    'Me.Width = Text1.Width
    'Me.Height = Text1.Height
    'Me.Left = TchNcellLeft
    'Me.Top = TchNcellTop
    
    Text1.Text = strTchNcell(EditFrmFlag - 1)
    strTemp = strTchNcell(EditFrmFlag - 1)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    EditFrmFlag = 0

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Or KeyAscii = &H1B Then
       strTchNcell(EditFrmFlag - 1) = Trim(Text1.Text)
       If strTchNcell(EditFrmFlag - 1) <> strTemp Then
          EditFlag(Int((EditFrmFlag - 1) / 2)) = True
       End If
       Unload Me
    End If
End Sub
