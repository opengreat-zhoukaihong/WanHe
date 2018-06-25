VERSION 5.00
Begin VB.Form SYSTEM 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "系统管理"
   ClientHeight    =   3180
   ClientLeft      =   2805
   ClientTop       =   1830
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "System.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "交换机类型选择"
      Height          =   1350
      Left            =   180
      TabIndex        =   7
      Top             =   1320
      Width           =   3495
      Begin VB.OptionButton Option1 
         Caption         =   "NORTEL"
         Height          =   240
         Index           =   5
         Left            =   1965
         TabIndex        =   13
         Top             =   990
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ITALTEL"
         Height          =   240
         Index           =   4
         Left            =   375
         TabIndex        =   12
         Top             =   990
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SIEMENS"
         Height          =   240
         Index           =   3
         Left            =   1950
         TabIndex        =   11
         Top             =   675
         Width           =   960
      End
      Begin VB.OptionButton Option1 
         Caption         =   "NOKIA"
         Height          =   240
         Index           =   2
         Left            =   375
         TabIndex        =   10
         Top             =   675
         Width           =   810
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MOTOROLA"
         Height          =   240
         Index           =   1
         Left            =   1950
         TabIndex        =   9
         Top             =   345
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ERICSSON"
         Height          =   240
         Index           =   0
         Left            =   375
         TabIndex        =   8
         Top             =   345
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "用户名和口令修改"
      Height          =   1095
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   3495
      Begin VB.TextBox PASSTEXT 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   1
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox USERNAME 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1170
         TabIndex        =   0
         Top             =   315
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "用户口令："
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   705
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "用户名称："
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.CommandButton PASSOK 
      Caption         =   "&O 确定"
      Height          =   320
      Left            =   765
      TabIndex        =   2
      Top             =   2805
      Width           =   1080
   End
   Begin VB.CommandButton PASSCANCEL 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   1950
      TabIndex        =   3
      Top             =   2805
      Width           =   1080
   End
End
Attribute VB_Name = "SYSTEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyRecord As Record  ' Declare variable.

Private Sub Form_Load()
    On Error Resume Next
    
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    Open Gsm_FileName For Random As #1 Len = Len(MyRecord)
    Get #1, 1, MyRecord  ' Read third record.
    Close #1
    USERNAME.Text = Trim(MyRecord.Name)
    PASSTEXT.Text = Trim(MyRecord.Pass)
    Option1(Val(MyRecord.exchange)).Value = True
    'If Val(MyRecord.Antenna) > 0 Then
    '   Text2.Text = Trim(MyRecord.Antenna)
    'End If
End Sub

Private Sub PASSCANCEL_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub PASSOK_Click()
    Dim i As Integer
    On Error Resume Next
    If Len(PASSTEXT.Text) <> 6 Then
       MsgBox "密码须输满6位!", 64, "提示"
       Exit Sub
    End If
    USERNAME = Trim(USERNAME.Text)
    MyRecord.Name = Trim(USERNAME.Text)
    MyRecord.Pass = Trim(PASSTEXT.Text)
    'If Val(Trim(Text2.Text)) <= 0 Then
    '   MyRecord.Antenna = "200"
    'Else
    '   MyRecord.Antenna = Trim(Text2.Text)
    'End If
    For i = 0 To 5
        If Option1(i).Value = True Then
           MyRecord.exchange = i
           Exit For
        End If
    Next
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    Open Gsm_FileName For Random As #1 Len = Len(MyRecord)
    Put #1, 1, MyRecord  ' Read third record.
    Close #1
    Unload Me
End Sub
