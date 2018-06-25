VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Full_Sub 
   BackColor       =   &H00C0C0C0&
   Caption         =   "FULL数据转为SUB ----- 文件选择"
   ClientHeight    =   4575
   ClientLeft      =   2100
   ClientTop       =   2355
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "System"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4575
   ScaleWidth      =   7245
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
      Height          =   495
      Left            =   5580
      TabIndex        =   1
      Top             =   420
      Width           =   1455
   End
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
      Height          =   495
      Left            =   5580
      TabIndex        =   0
      Top             =   1155
      Width           =   1455
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3870
      Index           =   0
      Left            =   300
      TabIndex        =   2
      Top             =   420
      Width           =   5025
      _Version        =   65536
      _ExtentX        =   8864
      _ExtentY        =   6826
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
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
         Height          =   360
         Left            =   165
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   660
         Width           =   2205
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
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3405
         Width           =   2235
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
         Height          =   360
         Left            =   2565
         TabIndex        =   5
         Top             =   3390
         Width           =   2250
      End
      Begin VB.FileListBox File1 
         Height          =   1770
         Left            =   180
         MultiSelect     =   1  'Simple
         Pattern         =   "*.tab"
         TabIndex        =   4
         Top             =   1140
         Width           =   2115
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   2565
         TabIndex        =   3
         Top             =   1125
         Width           =   2250
      End
      Begin VB.Label Label3 
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
         Height          =   300
         Left            =   2580
         TabIndex        =   13
         Top             =   720
         Width           =   2235
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "驱动器："
         Height          =   240
         Index           =   3
         Left            =   2565
         TabIndex        =   12
         Top             =   3090
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "文件类型："
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   11
         Top             =   3105
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "目录："
         Height          =   240
         Index           =   1
         Left            =   2505
         TabIndex        =   10
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "文件名："
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   9
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   2550
         TabIndex        =   8
         Top             =   750
         Width           =   75
      End
   End
End
Attribute VB_Name = "Full_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim snum As Integer
Dim sfile(1 To 10) As String
Dim tfile(1 To 10) As String

Private Sub C_cancel_Click()
    Menu_Flag = 9999
    Unload Me
'    Set mapinfo = Nothing
End Sub

Private Sub C_ok_Click()
    msg = Label3.Caption
    Unload Me
End Sub


Private Sub Combo1_Click()
    On Error Resume Next
    File1.Pattern = Combo1.Text
End Sub

Sub Drive1_Change()
    Dim dd As Integer
    On Error GoTo fixit
    Dir1.path = Drive1.Drive
    Exit Sub
fixit:
    dd = MsgBox("无法读取磁盘驱动器 " + Drive1.Drive + Chr(10) + Chr(10) + "请检查驱动器的门是否已关好！   ", 64, "打开文件")
    Drive1.Drive = "c:"

End Sub

Sub Dir1_Change()
    On Error Resume Next
    File1.path = Dir1.path
    Label3.Caption = Dir1.path
End Sub

Private Sub File1_Click()
    Dim i As Integer
    Dim ddd As String, sd As String * 1
    On Error Resume Next
    Text1.Text = ""
    ali_num = 0
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) = True Then
           Text1.Text = Text1.Text + File1.List(i) + "  "
           ali_num = ali_num + 1

        msg = Label3.Caption
        If Right(msg, 1) = "\" Then
           msg = Left(msg, Len(msg) - 1)
        End If
           sfile(ali_num) = msg + "\" + File1.List(i)
           ddd = File1.List(i)
           ddd = Left(ddd, Len(ddd) - 4)
           tfile(ali_num) = msg + "\" + ddd + ".tab"
           sd = Mid(ddd, 1, 1)
           If Asc(sd) > 47 And Asc(sd) < 58 Then
              ali_xls(ali_num) = "_" + ddd
           Else
              ali_xls(ali_num) = ddd
           End If
        End If
    Next
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Text1.Text = "*.tab"
    Combo1.ListIndex = 0
    Gsm_FileName = Gsm_Path + "\normal"
    Dir1.path = Gsm_FileName
    Label3.Caption = Dir1.path
End Sub
