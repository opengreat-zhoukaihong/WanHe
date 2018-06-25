VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form File_cho 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tems测量数据文件选择"
   ClientHeight    =   5760
   ClientLeft      =   855
   ClientTop       =   3285
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5760
   ScaleWidth      =   8160
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
      Height          =   465
      Left            =   6480
      TabIndex        =   1
      Top             =   1155
      Width           =   1245
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
      Height          =   450
      Left            =   6480
      TabIndex        =   0
      Top             =   360
      Width           =   1245
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1140
      Left            =   240
      TabIndex        =   13
      Top             =   4395
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   2011
      _StockProps     =   14
      Caption         =   "转换选择"
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
      Begin Threed.SSOption Option1 
         Height          =   375
         Left            =   3000
         TabIndex        =   15
         Top             =   480
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "滤除相同经纬度"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   11.99
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption Option2 
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   480
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "地理点平滑处理"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   11.99
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4005
      Left            =   195
      TabIndex        =   2
      Top             =   210
      Width           =   5955
      _Version        =   65536
      _ExtentX        =   10504
      _ExtentY        =   7064
      _StockProps     =   14
      Caption         =   "打开文件"
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
         Left            =   3150
         TabIndex        =   7
         Top             =   1125
         Width           =   2250
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   495
         MultiSelect     =   1  'Simple
         TabIndex        =   6
         Top             =   1200
         Width           =   2220
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
         Left            =   3180
         TabIndex        =   5
         Top             =   3465
         Width           =   2250
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
         Height          =   300
         Left            =   495
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3480
         Width           =   2235
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
         Height          =   285
         Left            =   495
         TabIndex        =   3
         Top             =   735
         Width           =   2205
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
         Left            =   3165
         TabIndex        =   12
         Top             =   750
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "文件名："
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   495
         TabIndex        =   11
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "目录："
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3120
         TabIndex        =   10
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "文件类型："
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   9
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "驱动器："
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   3195
         TabIndex        =   8
         Top             =   3120
         Width           =   960
      End
   End
End
Attribute VB_Name = "File_cho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub C_cancel_Click()
    Unload Me
    Menu_Flag = 555
End Sub

Private Sub C_ok_Click()
    Dim tex As String
    Dim dd As Integer
    On Error Resume Next
    If Option1.Value = True Then
       tran_del = True
    Else
       tran_del = False
    End If
    
    tex = Text1.Text
    If Len(tex) > 0 Then
       tex = RTrim$(tex)
       If Len(tex) = 0 Then
          dd = MsgBox("无法打开文件   ", 48, "打开文件")
          Exit Sub
       End If
    Else
       dd = MsgBox("无法打开文件   ", 48, "打开文件")
       Exit Sub
    End If

    Err = 0
    Open tran_f(1) For Binary As #1
    If Err Then
       dd = MsgBox("无法打开文件   ", 48, "打开文件")
       Exit Sub
    End If
    Close #1
    
    Unload Me
End Sub

Private Sub Combo1_Click()
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
    File1.path = Dir1.path
    Label1.Caption = Dir1.path
End Sub

Private Sub File1_Click()
    Dim i As Integer
    Text1.Text = ""
    tran_fn = 0
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) = True Then
           Text1.Text = Text1.Text + File1.List(i) + "  "
           tran_fn = tran_fn + 1
 
           msg = Label1.Caption
          If Right(msg, 1) = "\" Then
              msg = Left(msg, Len(msg) - 1)
           End If
            tran_f(tran_fn) = msg + "\" + File1.List(i)
        End If
    Next
End Sub

Private Sub Form_Load()
    If Menu_Flag = 121 Then
       Text1.Text = "*.txt"
       Combo1.ListIndex = 0
       File_cho.Caption = "Tems测量数据文件选择"
    End If

    If Menu_Flag = 124 Then
       Text1.Text = "*.00*"
       Combo1.ListIndex = 1
       File_cho.Caption = "Grayson Surveyor 测量数据文件(USWEST ISG format)选择"
    End If
    Label1.Caption = Dir1.path
    tran_del = False
End Sub

