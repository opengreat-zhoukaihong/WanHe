VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form Replay 
   BackColor       =   &H00C0C0C0&
   Caption         =   "回放选择"
   ClientHeight    =   3690
   ClientLeft      =   2550
   ClientTop       =   2340
   ClientWidth     =   4155
   Icon            =   "Replay.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3690
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "信息选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   240
      TabIndex        =   9
      Top             =   1650
      Width           =   3660
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1530
         TabIndex        =   16
         Top             =   1080
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   503
         _Version        =   327680
         BuddyControl    =   "Replay_Times"
         BuddyDispid     =   196611
         OrigLeft        =   1515
         OrigTop         =   1080
         OrigRight       =   1755
         OrigBottom      =   1365
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Replay_Times 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1065
         TabIndex        =   15
         Text            =   "1"
         Top             =   1080
         Width           =   465
      End
      Begin VB.ComboBox Replay_Msg2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1065
         TabIndex        =   14
         Text            =   "Release"
         Top             =   705
         Width           =   2370
      End
      Begin VB.ComboBox Replay_Msg1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1065
         TabIndex        =   13
         Text            =   "Setup"
         Top             =   330
         Width           =   2370
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "触发次数"
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
         Index           =   2
         Left            =   210
         TabIndex        =   12
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "结束信息"
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
         Left            =   210
         TabIndex        =   11
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "启始信息"
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
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   375
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "背景选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   2475
      TabIndex        =   6
      Top             =   150
      Width           =   1425
      Begin VB.OptionButton Option5 
         Caption         =   "RxLevSub"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   855
         Width           =   1065
      End
      Begin VB.OptionButton Option4 
         Caption         =   "RxLevFull"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   7
         Top             =   495
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "回放方式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   240
      TabIndex        =   2
      Top             =   135
      Width           =   2145
      Begin VB.OptionButton Option3 
         Caption         =   "启始/结束信息触发"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   1050
         Width           =   1845
      End
      Begin VB.OptionButton Option2 
         Caption         =   "启始信息触发"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   705
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Caption         =   "当前点触发"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton PASSCANCEL 
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
      Left            =   2190
      TabIndex        =   1
      Top             =   3315
      Width           =   1080
   End
   Begin VB.CommandButton PASSOK 
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
      Left            =   885
      TabIndex        =   0
      Top             =   3315
      Width           =   1080
   End
End
Attribute VB_Name = "Replay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\tems_msg.tab"
    If Dir(Gsm_FileName) = "" Then
       GoTo no_tems_msg
    End If
  mapinfo.do "open table " + Chr(34) + Gsm_Path + "\tems_msg" + Chr(34)
  i = 0
  row = Val(mapinfo.eval("tableinfo(tems_msg,8)"))
  mapinfo.do "fetch First from tems_msg"
  While i < row
       Replay_Msg1.AddItem mapinfo.eval("tems_msg.message")
       Replay_Msg2.AddItem mapinfo.eval("tems_msg.message")
       mapinfo.do "fetch next from tems_msg"
       i = i + 1
  Wend
  mapinfo.do "close table tems_msg"
no_tems_msg:
  Replay_flag = 0
  Back_Sel = 0
End Sub


Private Sub PASSCANCEL_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub PASSOK_Click()
    On Error Resume Next
    If Option1.Value Then
    
    Else
    
    End If
    rmsg1 = Replay_Msg1.Text
    rmsg2 = Replay_Msg2.Text
    Replay_Time = Replay_Times.Text
    Unload Me
  i = Val(mapinfo.eval("selectionInfo(3)"))  ' SEL_INFO_NROWS
  If i <> 0 Then
'      MDIMain.SUB_532.Enabled = 1
      Load MapForm
      mapHWnd = Val(mapinfo.eval("WindowInfo(" & mapid & ",12)"))
      If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
         MapForm.WindowState = 0
      End If
      MapForm.Move 0, 10, 12000, 4050

      Load Graph
      'Graph.Move 0, 4050, 6950, 3150
      Graph.Move 0, 4050, 7000, 3495

      Load msgdis
      'msgdis.Move 6950, 4050, 5020, 3150
      msgdis.Move 6950, 4050, 5020, 3495
  End If
End Sub

Private Sub Option1_Click()
    On Error Resume Next
    Replay_flag = 0
End Sub

Private Sub Option2_Click()
    Replay_flag = 1
End Sub

Private Sub Option3_Click()
    Replay_flag = 2
End Sub

Private Sub Option4_Click()
    Back_Sel = 0
End Sub

Private Sub Option5_Click()
    Back_Sel = 1
End Sub
