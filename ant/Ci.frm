VERSION 5.00
Begin VB.Form CI_Cell 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " 按CI查找小区"
   ClientHeight    =   1755
   ClientLeft      =   2970
   ClientTop       =   2790
   ClientWidth     =   3540
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Ci.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1755
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   270
      TabIndex        =   3
      Top             =   105
      Width           =   2970
      Begin VB.TextBox ARFCN 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1110
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   435
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "请输入CI："
         DragMode        =   1  'Automatic
         Height          =   180
         Index           =   1
         Left            =   255
         TabIndex        =   5
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "（十进制）"
         DragMode        =   1  'Automatic
         Height          =   180
         Index           =   0
         Left            =   1905
         TabIndex        =   4
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.CommandButton OK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   615
      TabIndex        =   1
      Top             =   1335
      Width           =   1080
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   1830
      TabIndex        =   2
      Top             =   1335
      Width           =   1080
   End
End
Attribute VB_Name = "CI_Cell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ARFCN_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
       KeyCode = 0
       OK_Click
    End If

End Sub

Private Sub Cancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If Menu_Flag = 4700 Or Menu_Flag = 4788 Then
       If Menu_Flag = 4700 Then
          CI_Cell.Caption = "按 BaseNo 查找小区"
          Label1(1).Caption = "请输入BaseNo:"
       Else
          CI_Cell.Caption = "按 BSIC 查找小区"
          Label1(1).Caption = "请输入 BSIC:"
       End If
       Label1(0).Caption = ""
       Label1(1).Left = 360
       ARFCN.MaxLength = 20
       ARFCN.Left = 1550
       ARFCN.Width = 1000
    End If
End Sub

Private Sub OK_Click()
    Dim YouCI As String
    On Error Resume Next
    YouCI = Trim(ARFCN.Text)
    Unload Me
    If Menu_Flag = 4700 Then
       mapinfo.do "select  *  from cell where col2 = " + Chr(34) + YouCI + Chr(34) + " into My_BaseNo"
       row = Val(mapinfo.eval("tableinfo(My_BaseNo,8)"))
       If row < 1 Then
          MsgBox "所查找的小区不存在！", 64, "提示"
       Else
          mapinfo.do "Add Map Auto Layer " + Chr(34) + "My_BaseNo" + Chr(34)
          mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
          mapinfo.do "browse * from My_BaseNo"
          mapinfo.do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
       End If
    Else
       If Menu_Flag = 4788 Then
          mapinfo.do "select  *  from cell where bsic = " & YouCI & " into My_BSIC"
          row = Val(mapinfo.eval("tableinfo(My_BSIC,8)"))
          If row < 1 Then
             MsgBox "所查找的小区不存在！", 64, "提示"
          Else
             mapinfo.do "Add Map Auto Layer " + Chr(34) + "My_BSIC" + Chr(34)
             mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
             mapinfo.do "browse * from My_BSIC"
             mapinfo.do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
          End If
       Else
          mapinfo.do "select  *  from cell where CI = " + Chr(34) + YouCI + Chr(34) + " into My_cell"
          row = Val(mapinfo.eval("tableinfo(my_cell,8)"))
          If row < 1 Then
             MsgBox "所查找的小区不存在！", 64, "提示"
          Else
             msg = "Add Map Auto Layer " + Chr(34) + "My_cell" + Chr(34)
             mapinfo.do msg
             mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
             mapinfo.do "browse * from   My_cell"
             mapinfo.do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
          End If
       End If
    End If
End Sub
