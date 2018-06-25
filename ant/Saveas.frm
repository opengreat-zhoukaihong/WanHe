VERSION 5.00
Begin VB.Form SaveAs 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "另存文件"
   ClientHeight    =   3555
   ClientLeft      =   3075
   ClientTop       =   3960
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   7665
   Begin VB.CommandButton C_cancel 
      Caption         =   " 取  消"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6105
      TabIndex        =   6
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton C_ok 
      Caption         =   "  确  定 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6135
      TabIndex        =   5
      Top             =   375
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   420
      Width           =   2400
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   510
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2910
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3255
      TabIndex        =   2
      Top             =   2910
      Width           =   2430
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   495
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   810
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   3240
      TabIndex        =   0
      Top             =   780
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drives:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3270
      TabIndex        =   11
      Top             =   2670
      Width           =   1665
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "List Files of Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   510
      TabIndex        =   10
      Top             =   2670
      Width           =   1650
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   510
      TabIndex        =   9
      Top             =   150
      Width           =   2085
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Directories:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3225
      TabIndex        =   8
      Top             =   165
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3255
      TabIndex        =   7
      Top             =   465
      Width           =   2400
   End
End
Attribute VB_Name = "SaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub C_cancel_Click()
    Unload Me
End Sub

Private Sub C_ok_Click()
   On Error Resume Next
   Dim new_file As String

    msg = Label1.Caption
    If Right(msg, 1) = "\" Then
       msg = Left(msg, Len(msg) - 1)
    End If
   new_file = msg + "\" + Text1.Text
   MapInfo.do "commit table " & tblname & " as " + Chr(34) + new_file + Chr(34)
   Unload Me
End Sub

Private Sub Combo1_Click()
    On Error Resume Next
    File1.Pattern = Combo1.Text
End Sub

Sub Drive1_Change()

    On Error GoTo fixit
    Dir1.path = Drive1.Drive
    Exit Sub
fixit:
    dd = MsgBox("无法读取磁盘驱动器 " + Drive1.Drive + Chr(10) + Chr(10) + "请检查驱动器的门是否已关好！   ", 64, "Open File")
    Drive1.Drive = "c:"
End Sub

Sub Dir1_Change()
    On Error Resume Next
    File1.path = Dir1.path
    Label1.Caption = Dir1.path
End Sub

Private Sub File1_Click()
    On Error Resume Next
    Text1.Text = Trim(File1.filename)

    msg = Label1.Caption
    If Right(msg, 1) = "\" Then
       msg = Left(msg, Len(msg) - 1)
    End If
    ncell_file = msg + "\" + File1.filename
End Sub


Private Sub Form_Load()
    On Error Resume Next
    Text1.Text = "*.txt"
    Combo1.ListIndex = 0
    Label1.Caption = Dir1.path
End Sub

