VERSION 5.00
Begin VB.Form Find_Lac 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " 按LAC查找小区"
   ClientHeight    =   1875
   ClientLeft      =   2610
   ClientTop       =   3165
   ClientWidth     =   3240
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lac.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1875
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   255
      TabIndex        =   3
      Top             =   120
      Width           =   2760
      Begin VB.TextBox ARFCN 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "请输入LAC："
         DragMode        =   1  'Automatic
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   5
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "(DEC)"
         DragMode        =   1  'Automatic
         Height          =   180
         Index           =   0
         Left            =   2070
         TabIndex        =   4
         Top             =   525
         Width           =   465
      End
   End
   Begin VB.CommandButton OK 
      Caption         =   "&O 确认"
      Height          =   300
      Left            =   435
      TabIndex        =   1
      Top             =   1455
      Width           =   1080
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&C 取消"
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   1455
      Width           =   1080
   End
End
Attribute VB_Name = "Find_Lac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub OK_Click()
Dim YouLAC As String
Dim row As Integer

On Error Resume Next
YouLAC = Trim(ARFCN.Text)
Unload Me

        mapinfo.do "select  *  from cell where LAC = " & YouLAC & " into My_LAC"

        row = Val(mapinfo.eval("tableinfo(my_lac,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
           msg = "Add Map Auto Layer " + Chr(34) + "My_LAC" + Chr(34)
           mapinfo.do msg
                     
           mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
           mapinfo.do "browse * from   My_LAC"
           mapinfo.do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
        End If
End Sub
