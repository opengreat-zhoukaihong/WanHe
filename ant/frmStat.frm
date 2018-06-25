VERSION 5.00
Begin VB.Form frmStat 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "统计结果"
   ClientHeight    =   4080
   ClientLeft      =   2115
   ClientTop       =   1305
   ClientWidth     =   4440
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
   ScaleHeight     =   4080
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "关闭"
      Height          =   375
      Left            =   3045
      TabIndex        =   1
      Top             =   3375
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   4035
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4395
   End
End
Attribute VB_Name = "frmStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim hFreefile As Integer
    Dim myFilename As String
    Dim i As Integer
    
    On Error Resume Next
    myFilename = CellFileName & ".log"
    MDIMain.FileDialog.CancelError = True
    MDIMain.FileDialog.filename = myFilename
    MDIMain.FileDialog.Filter = "*.log"
    MDIMain.FileDialog.DefaultExt = "*.log"
    MDIMain.FileDialog.Flags = &H80000
    MDIMain.FileDialog.InitDir = Gsm_Path & "\User"
    Err = 0
open_again:
    MDIMain.FileDialog.ShowSave
    If Err Then
       GoTo ExitWindow
    End If
    If MDIMain.FileDialog.filename <> "" Then
       myFilename = MDIMain.FileDialog.filename
       If Dir(myFilename, 0) <> "" Then
          i = MsgBox(myFilename & " 已存在，是否将它覆盖？", 49, "保存文件")
          If i = 2 Then
             MDIMain.FileDialog.filename = ""
             GoTo open_again
          Else
             Kill myFilename
          End If
       End If
    Else
       GoTo ExitWindow
    End If
    hFreefile = FreeFile
    Open myFilename For Binary As #hFreefile
    StatString = Trim(Text1.Text)
    Put #hFreefile, , StatString
    Close #hFreefile
    
ExitWindow:
    MDIMain.FileDialog.filename = ""
    MDIMain.FileDialog.CancelError = False
    StatString = ""
    Unload Me

End Sub

Private Sub Form_Load()
    On Error Resume Next
    Text1.Text = StatString
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Width = ScaleWidth - 10
    Text1.Height = ScaleHeight
    Command1.Left = Me.Width - (4560 - 3045)
    Command1.Top = Me.Height - (4485 - 3372)
    
End Sub
