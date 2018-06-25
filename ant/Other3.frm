VERSION 5.00
Begin VB.Form OTHER3 
   BackColor       =   &H00C0C0C0&
   Caption         =   "其它信令分析"
   ClientHeight    =   1845
   ClientLeft      =   3315
   ClientTop       =   3825
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Other3.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1845
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox SEL_3_MSG 
      Height          =   300
      Left            =   1125
      TabIndex        =   4
      Text            =   "Setup"
      Top             =   345
      Width           =   2865
   End
   Begin VB.CommandButton CANCEL 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   2145
      TabIndex        =   2
      Top             =   1395
      Width           =   1080
   End
   Begin VB.CommandButton OK 
      Caption         =   "&O 确定"
      Height          =   320
      Left            =   930
      TabIndex        =   1
      Top             =   1395
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "信令临时表名称：RESULT"
      Height          =   180
      Left            =   195
      TabIndex        =   3
      Top             =   885
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "所选信令："
      Height          =   180
      Left            =   195
      TabIndex        =   0
      Top             =   405
      Width           =   900
   End
End
Attribute VB_Name = "OTHER3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  On Error Resume Next
  mapinfo.do "open table " + Chr(34) + Gsm_Path + "\tems_msg" + Chr(34)
  i = 0
  row = Val(mapinfo.eval("tableinfo(tems_msg,8)"))
  mapinfo.do "fetch First from tems_msg"
  While i < row
       SEL_3_MSG.AddItem mapinfo.eval("tems_msg.message")
       mapinfo.do "fetch next from tems_msg"
       i = i + 1
  Wend
  mapinfo.do "close table tems_msg"
End Sub


Private Sub OK_Click()
    On Error Resume Next
    Msg_3_Layer = SEL_3_MSG.Text
    Unload Me
    SelTable.Show 1
End Sub
