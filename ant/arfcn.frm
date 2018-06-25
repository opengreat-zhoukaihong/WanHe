VERSION 5.00
Begin VB.Form ARFCN 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " 按ARFCN查找同频小区"
   ClientHeight    =   1830
   ClientLeft      =   4470
   ClientTop       =   4020
   ClientWidth     =   4455
   Icon            =   "arfcn.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1830
   ScaleWidth      =   4455
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   225
      TabIndex        =   5
      Top             =   135
      Width           =   2670
      Begin VB.TextBox ARFCN 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1635
         MaxLength       =   4
         TabIndex        =   0
         Text            =   " "
         Top             =   420
         Width           =   615
      End
      Begin VB.OptionButton SSOption1 
         Caption         =   "TCH"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1575
         TabIndex        =   2
         Top             =   1110
         Width           =   570
      End
      Begin VB.OptionButton SSOption1 
         Caption         =   "CCH"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   1095
         Value           =   -1  'True
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "请输入ARFCN："
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   435
         TabIndex        =   6
         Top             =   465
         Width           =   1170
      End
   End
   Begin VB.CommandButton OK 
      Caption         =   "&O 确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   3210
      TabIndex        =   3
      Top             =   315
      Width           =   1080
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "&C 取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   3210
      TabIndex        =   4
      Top             =   690
      Width           =   1080
   End
End
Attribute VB_Name = "ARFCN"
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

Private Sub OK_Click()
    Dim YouARF As Integer
    On Error Resume Next
    YouARF = Val(ARFCN.Text)
    If SSOption1(0).Value = True Then
       CELL_CCH = 1
    Else
       CELL_CCH = 0
    End If
    Unload Me

       On Error Resume Next
       If CELL_CCH = 1 Then
         mapinfo.do "select  *  from cell where ARFCN = " & YouARF & " into same_arfcn"
       Else
         'mapinfo.do "select  *  from cell where Non_bcch_1 = " & YouARF & " or Non_bcch_2 = " & YouARF & "or Non_bcch_3 = " & YouARF & "or Non_bcch_4 = " & YouARF & "or Non_bcch_5 = " & YouARF & "or Non_bcch_6 = " & YouARF & " into same_arfcn"
         mapinfo.do "Select * from cell where Like(Non_bcch,""%" & Trim(YouARF) & "%"","""") = 1 into same_arfcn"
       End If
        row = Val(mapinfo.eval("tableinfo(same_arfcn,8)"))
        If row < 1 Then
           MsgBox "所查找的小区不存在！", 64, "提示"
        Else
        msg = "Add Map Auto Layer " + Chr(34) + "same_arfcn" + Chr(34)
        mapinfo.do msg
        
        mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
        mapinfo.do "browse * from   same_arfcn"
        mapinfo.do "set window Frontwindow() Position(0,4) Width 8 Height 1 "
        End If
End Sub

Private Sub SSOption1_Click(Index As Integer)
        If Value = True Then
           If Index = 0 Then
              CELL_CCH = 1
           Else
              CELL_CCH = 0
           End If
        End If
End Sub
