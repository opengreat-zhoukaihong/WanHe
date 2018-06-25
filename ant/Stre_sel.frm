VERSION 5.00
Begin VB.Form Stre_Sel 
   BackColor       =   &H00C0C0C0&
   Caption         =   "采集测量数据统计项选择"
   ClientHeight    =   4005
   ClientLeft      =   2130
   ClientTop       =   1545
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Stre_sel.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   6690
   Begin VB.Frame Frame1 
      Caption         =   "测试报告选择"
      Height          =   3360
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   240
         Left            =   1125
         TabIndex        =   21
         Text            =   "10"
         Top             =   2970
         Width           =   435
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TA 和 MS Power"
         Height          =   240
         Index           =   3
         Left            =   540
         TabIndex        =   20
         Top             =   1125
         Width           =   1560
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxLev 和 RxQual"
         Height          =   240
         Index           =   1
         Left            =   540
         TabIndex        =   19
         Top             =   735
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CommandButton Command1 
         Caption         =   "统计范围"
         Height          =   320
         Left            =   2250
         TabIndex        =   18
         Top             =   690
         Width           =   960
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   240
         Left            =   4950
         TabIndex        =   17
         Text            =   "9"
         Top             =   2085
         Width           =   435
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   240
         Left            =   4950
         TabIndex        =   14
         Text            =   "12"
         Top             =   1695
         Width           =   435
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ncell对BCCH、TCH 频率碰撞统计"
         Height          =   240
         Index           =   11
         Left            =   3390
         TabIndex        =   13
         Top             =   1275
         Width           =   2910
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sub"
         Enabled         =   0   'False
         Height          =   240
         Left            =   4635
         TabIndex        =   11
         Top             =   2535
         Width           =   750
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Full"
         Enabled         =   0   'False
         Height          =   240
         Left            =   3705
         TabIndex        =   10
         Top             =   2520
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   240
         Left            =   1380
         TabIndex        =   9
         Text            =   "3"
         Top             =   2640
         Width           =   435
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "无线环境质量参数统计报告"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   2460
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "信令事件统计报告"
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1530
         Width           =   1800
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "天线测量统计对比报告"
         Height          =   240
         Index           =   4
         Left            =   3390
         TabIndex        =   5
         Top             =   810
         Width           =   2160
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "小区质量排行榜"
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   4
         Top             =   1905
         Width           =   1605
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "系统响应时间统计报告"
         Height          =   240
         Index           =   9
         Left            =   3390
         TabIndex        =   3
         Top             =   345
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         Height          =   180
         Index           =   2
         Left            =   1665
         TabIndex        =   23
         Top             =   3000
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "FER > "
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   22
         Top             =   3000
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "邻频场强差值:"
         Height          =   180
         Left            =   3705
         TabIndex        =   16
         Top             =   2115
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "同频场强差值:"
         Height          =   180
         Left            =   3705
         TabIndex        =   15
         Top             =   1710
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxQual > "
         Height          =   180
         Index           =   0
         Left            =   570
         TabIndex        =   12
         Top             =   2670
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "信号质量门限:"
         Height          =   180
         Left            =   555
         TabIndex        =   8
         Top             =   2295
         Width           =   1170
      End
   End
   Begin VB.CommandButton C_ok 
      Caption         =   "确定"
      Height          =   320
      Left            =   2040
      TabIndex        =   1
      Top             =   3615
      Width           =   1080
   End
   Begin VB.CommandButton C_cancel 
      Caption         =   "取消"
      Height          =   320
      Left            =   3285
      TabIndex        =   0
      Top             =   3615
      Width           =   1080
   End
End
Attribute VB_Name = "Stre_Sel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub C_Cancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub C_OK_Click()
    Dim i As Integer, finds As Integer, dd As Integer
    
    On Error Resume Next
 If Menu_Flag = 7001 Then
    For i = 1 To 50
        If convert_filename(i) = "" Then
           stre_num = i - 1
           Exit For
        End If
        Err = 0
        mapinfo.do "open table " + Chr(34) + convert_filename(i) + Chr(34)
        If Err Then
           dd = MsgBox("无法打开文件   " + convert_filename(i), 48, "打开文件")
           Exit Sub
        End If
        stre_tab(i) = convert_filename(i)
        finds = InStr(stre_tab(i), ".")
        If finds > 0 Then
           stre_tab(i) = Left(stre_tab(i), finds - 1)
        End If
        finds = InStr(stre_tab(i), "\")
        Do While finds > 0
           stre_tab(i) = Right(stre_tab(i), Len(stre_tab(i)) - finds)
           finds = InStr(stre_tab(i), "\")
        Loop
        If Asc(Left(stre_tab(i), 1)) > 47 And Asc(Left(stre_tab(i), 1)) < 58 Then
           stre_tab(i) = "_" + stre_tab(i)
        End If
    Next
 End If
    If RangeNum = 0 Then
       RangeNum = 3
       RxLevRange(1, 1) = "27"
       RxLevRange(1, 2) = "17"
       RxLevRange(1, 3) = "0"
       RxLevRange(2, 1) = "63"
       RxLevRange(2, 2) = "27"
       RxLevRange(2, 3) = "17"
    End If
    For i = 0 To 11
        If i <> 6 And i <> 7 And i <> 8 And i <> 10 Then
           If Check1(i).Value = 1 Then
              stre_s(i) = True
           Else
              stre_s(i) = False
           End If
        Else
           stre_s(i) = False
        End If
    Next i
    If stre_s(0) = True And stre_s(1) = False And stre_s(3) = False Then
       stre_s(1) = True
    End If
    Report_Qual = Val(Text1.Text)
    Report_FER = Val(Text4.Text)
    Report_Rxlev1 = Val(Text2.Text)
    Report_Rxlev2 = Val(Text3.Text)
    If Option2.Value = True Then
       Report_Full = False
    Else
       Report_Full = True
    End If
'    If Check2.Value = 1 Then
'       Cell_Report = True
'       Unload Me
'       Load Cell_Report_Frm
'       Cell_Report_Frm.Move 4000, 2000, 3900, 3100
'    Else
       Cell_Report = False
    If Menu_Flag = 7001 Then
       For i = 1 To stre_num
           mapinfo.do "fetch first from " & stre_tab(i)
       Next
       mapinfo.do "open table " + Chr(34) + Gsm_Path + "\map\cell" + Chr(34)
       Unload Me
       My_Report
    Else
       Unload Me
    End If
'    End If
End Sub

Private Sub Check1_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
       Case 0
          If Check1(0).Value = 1 Then
             Command1.Enabled = True
             Check1(1).Enabled = True
             Check1(3).Enabled = True
          Else
             Command1.Enabled = False
             Check1(1).Enabled = False
             Check1(3).Enabled = False
          End If
       Case 1
          If Check1(1).Value = 1 Then
             Command1.Enabled = True
          Else
             Command1.Enabled = False
          End If
       Case 11
          If Check1(11).Value = 1 Then
             Text2.Enabled = True
             Text3.Enabled = True
             Option1.Enabled = True
             Option2.Enabled = True
          Else
             Text2.Enabled = False
             Text3.Enabled = False
             Option1.Enabled = False
             Option2.Enabled = False
          End If
       Case 5
          If Check1(5).Value = 1 Then
             Text1.Enabled = True
             Text4.Enabled = True
          Else
             Text1.Enabled = False
             Text4.Enabled = False
          End If
    End Select
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    FrmRange.Show 1
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    RangeNum = 0
    If IsQuickConvert Then
       Check1(9).Enabled = False
       Check1(2).Enabled = False
       'Check1(2).ToolTipText = "文件中存在快速转换文件，不能选择该项"
       'Check1(9).ToolTipText = "文件中存在快速转换文件，不能选择该项"
    End If
End Sub
