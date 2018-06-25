VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Config_frm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "地图配置"
   ClientHeight    =   3435
   ClientLeft      =   2895
   ClientTop       =   3600
   ClientWidth     =   5085
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3435
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2705
      Left            =   330
      TabIndex        =   19
      Top             =   495
      Width           =   3105
      Begin VB.CheckBox Check3 
         Caption         =   "自动标注（小区CI）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   4
         Left            =   270
         TabIndex        =   26
         Top             =   2310
         Width           =   2025
      End
      Begin VB.CheckBox Check1 
         Caption         =   "自动标注（小区TCH）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   5
         Left            =   270
         TabIndex        =   25
         Top             =   1965
         Width           =   2115
      End
      Begin VB.CheckBox Check1 
         Caption         =   "自动标注（小区BCCH）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   6
         Left            =   270
         TabIndex        =   24
         Top             =   1620
         Width           =   2115
      End
      Begin VB.CheckBox Check3 
         Caption         =   "GSM宏基站"
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
         Index           =   0
         Left            =   270
         TabIndex        =   23
         Top             =   240
         Width           =   1155
      End
      Begin VB.CheckBox Check3 
         Caption         =   "GSM微蜂窝"
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
         Index           =   1
         Left            =   270
         TabIndex        =   22
         Top             =   590
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "GSM微微蜂窝"
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
         Index           =   2
         Left            =   270
         TabIndex        =   21
         Top             =   930
         Width           =   1440
      End
      Begin VB.CheckBox Check3 
         Caption         =   "DCS宏基站"
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
         Index           =   3
         Left            =   270
         TabIndex        =   20
         Top             =   1275
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2705
      Left            =   330
      TabIndex        =   11
      Top             =   495
      Width           =   3105
      Begin VB.CheckBox Check3 
         Caption         =   "自动标注（区域名）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   5
         Left            =   270
         TabIndex        =   18
         Top             =   1995
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "市镇"
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
         Index           =   4
         Left            =   270
         TabIndex        =   17
         Top             =   1650
         Width           =   945
      End
      Begin VB.CheckBox Check1 
         Caption         =   "区域地图"
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
         Index           =   0
         Left            =   270
         TabIndex        =   16
         Top             =   240
         Width           =   1155
      End
      Begin VB.CheckBox Check1 
         Caption         =   "山峰＋水域＋绿化带"
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
         Index           =   1
         Left            =   270
         TabIndex        =   15
         Top             =   590
         Width           =   2010
      End
      Begin VB.CheckBox Check1 
         Caption         =   "基站＋小区天线"
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
         Index           =   2
         Left            =   270
         TabIndex        =   14
         Top             =   945
         Width           =   1665
      End
      Begin VB.CheckBox Check1 
         Caption         =   "街区"
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
         Index           =   3
         Left            =   270
         TabIndex        =   13
         Top             =   1275
         Width           =   945
      End
      Begin VB.CheckBox Check3 
         Caption         =   "自动标注（市镇名）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   6
         Left            =   270
         TabIndex        =   12
         Top             =   2325
         Width           =   1920
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3105
      Left            =   225
      TabIndex        =   10
      Top             =   180
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   5477
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "区域地图"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "街道地图"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "网络资源"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2705
      Left            =   330
      TabIndex        =   2
      Top             =   495
      Visible         =   0   'False
      Width           =   3105
      Begin VB.CheckBox Check2 
         Caption         =   "用户定义层（USER_3）"
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
         Index           =   6
         Left            =   270
         TabIndex        =   9
         Top             =   2340
         Width           =   2115
      End
      Begin VB.CheckBox Check2 
         Caption         =   "用户定义层（USER_2）"
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
         Index           =   5
         Left            =   270
         TabIndex        =   8
         Top             =   1990
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Caption         =   "用户定义层（USER_1）"
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
         Index           =   4
         Left            =   270
         TabIndex        =   7
         Top             =   1640
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Caption         =   "邮电机构"
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
         Index           =   3
         Left            =   270
         TabIndex        =   6
         Top             =   1290
         Width           =   1110
      End
      Begin VB.CheckBox Check2 
         Caption         =   "收费站"
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
         Index           =   2
         Left            =   270
         TabIndex        =   5
         Top             =   940
         Width           =   1005
      End
      Begin VB.CheckBox Check2 
         Caption         =   "公众场所、大型建筑物"
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
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   590
         Width           =   2220
      End
      Begin VB.CheckBox Check2 
         Caption         =   "街道及主要公路、桥梁、隧道"
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
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
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
      Left            =   3780
      TabIndex        =   1
      Top             =   495
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
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
      Left            =   3780
      TabIndex        =   0
      Top             =   930
      Width           =   1080
   End
End
Attribute VB_Name = "Config_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim select1(1 To 16) As String * 1
Dim select2(1 To 16) As String * 1
Dim select3(1 To 10) As String * 1
Dim Current_Page As Integer

Private Sub Check3_Click(Index As Integer)
    On Error Resume Next
    If Index = 4 Then
        If Check3(4).Value = 1 Then
            Check1(5).Value = 0
            Check1(6).Value = 0
        End If
    End If
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    
    On Error Resume Next
    
    Gsm_FileName = Gsm_Path + "\ant.cfg"
    If Dir(Gsm_FileName, 0) <> "" Then
        Open Gsm_FileName For Binary As #1
        For i = 1 To 7
            Get #1, i, select1(i)
            Check1(i - 1).Value = select1(i)
        Next
        Seek #1, 9
        For i = 1 To 7
            Get #1, i + 8, select2(i)
            Check2(i - 1).Value = select2(i)
        Next
        For i = 1 To 7
            If EOF(1) Then
               Check3(0).Value = 1
               Check3(1).Value = 1
               Check3(2).Value = 1
               Check3(3).Value = 0
               Check3(4).Value = 0
               Check3(5).Value = 0
               Check3(6).Value = 0
               Exit For
            End If
            Get #1, i + 16, select3(i)
            If select3(i) <> "0" And select3(i) <> "1" Then
               Check3(0).Value = 1
               Check3(1).Value = 1
               Check3(2).Value = 1
               Check3(3).Value = 0
               Check3(4).Value = 0
               Check3(5).Value = 0
               Check3(6).Value = 0
               Exit For
            End If
            Check3(i - 1).Value = select3(i)
        Next
        Close
    End If
    Current_Page = 0
    TabStrip1_Click
End Sub

Private Sub Check1_Click(Index As Integer)

    On Error Resume Next
    If Index = 5 Then
       If Check1(5).Value = 1 Then
          Check1(6).Value = 0
          Check3(4).Value = 0
       End If
    ElseIf Index = 6 Then
       If Check1(6).Value = 1 Then
          Check1(5).Value = 0
          Check3(4).Value = 0
       End If
    End If
    'If Check1(Index).Value = 1 Then
    '   select1(Index + 1) = "1"
    'Else
    '   select1(Index + 1) = "0"
    'End If
End Sub

Private Sub Command1_Click()

    Dim endeof As String * 1
    On Error Resume Next
    endeof = Chr$(26)
    Gsm_FileName = Gsm_Path + "\ant.cfg"
    Open Gsm_FileName For Binary As #1
    For i = 1 To 8
        If i < 8 Then
           If Check1(i - 1).Value = 1 Then
              select1(i) = "1"
           Else
              select1(i) = "0"
           End If
        End If
        Put #1, i, select1(i)
    Next
    Seek #1, 9
    For j = 1 To 8
        If j < 8 Then
           If Check2(j - 1).Value = 1 Then
              select2(j) = "1"
           Else
              select2(j) = "0"
           End If
        End If
        Put #1, j + 8, select2(j)
    Next
    For j = 1 To 7
        If Check3(j - 1).Value = 1 Then
           select3(j) = "1"
        Else
           select3(j) = "0"
        End If
        Put #1, , select3(j)
    Next
    
    Put #1, , endeof
    Close
    Unload Me
    Call MDIMain.OPen_All_Map_Click
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub TabStrip1_Click()
    On Error Resume Next
    If TabStrip1.SelectedItem.Index = Current_Page Then
       Exit Sub
    End If
    If TabStrip1.SelectedItem.Index = 1 Then
       Frame2.Visible = False
       Frame3.Visible = False
       Frame1.Visible = True
       Frame1.ZOrder 0
       Current_Page = 1
    ElseIf TabStrip1.SelectedItem.Index = 2 Then
       Frame1.Visible = False
       Frame2.Visible = True
       Frame3.Visible = False
       Frame2.ZOrder 0
       Current_Page = 2
    Else
       Frame1.Visible = False
       Frame2.Visible = False
       Frame3.Visible = True
       Frame3.ZOrder 0
       Current_Page = 3
    
    End If
End Sub
