VERSION 5.00
Begin VB.Form PassWord 
   BackColor       =   &H00C0C0C0&
   Caption         =   "用户检查"
   ClientHeight    =   1590
   ClientLeft      =   2565
   ClientTop       =   3090
   ClientWidth     =   2925
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Password.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1590
   ScaleWidth      =   2925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   195
      TabIndex        =   4
      Top             =   105
      Width           =   2565
      Begin VB.TextBox PASSTEXT 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1290
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   345
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "请输入口令："
         Height          =   180
         Left            =   210
         TabIndex        =   3
         Top             =   390
         Width           =   1080
      End
   End
   Begin VB.CommandButton PASSCANCEL 
      Cancel          =   -1  'True
      Caption         =   "&C 取消"
      Height          =   300
      Left            =   1530
      TabIndex        =   2
      Top             =   1200
      Width           =   1080
   End
   Begin VB.CommandButton PASSOK 
      Caption         =   "&O 确定"
      Height          =   300
      Left            =   345
      TabIndex        =   1
      Top             =   1200
      Width           =   1080
   End
End
Attribute VB_Name = "PassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub PASSCANCEL_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub PASSOK_Click()
  Dim MyRecord As Record  ' Declare variable.

  On Error Resume Next
  If Len(PASSTEXT.Text) <> 6 Then
     MsgBox "密码须输满6位!", 64, "提示"
     Exit Sub
  End If
  Gsm_FileName = Gsm_Path + "\gsm.dat"
  Open Gsm_FileName For Random As #1 Len = Len(MyRecord)
  Get #1, 1, MyRecord  ' Read third record.
  Close #1
   
    If PASSTEXT.Text = MyRecord.Pass Then
       Unload Me
       If Menu_Flag = 61 Then
          Load New_Base
          New_Base.Move 700, 200, 10440, 7320
       Else
          SYSTEM.Show 1
       End If
    Else
       MsgBox "口令不对,请再试一次！", 64, "提示"
    End If
End Sub

Private Sub PASSTEXT_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
       KeyCode = 0
       PASSOK_Click
    End If
End Sub
