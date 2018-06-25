VERSION 5.00
Begin VB.Form frmNcellA 
   Caption         =   "邻小区合理性统计"
   ClientHeight    =   3645
   ClientLeft      =   5295
   ClientTop       =   1185
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNcellA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5130
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   330
      Left            =   2550
      TabIndex        =   2
      Top             =   3225
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   345
      Left            =   1380
      TabIndex        =   1
      Top             =   3225
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   2550
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   4830
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   4020
         TabIndex        =   24
         Top             =   405
         Width           =   645
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   4020
         TabIndex        =   23
         Top             =   780
         Width           =   645
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   4035
         TabIndex        =   22
         Top             =   1140
         Width           =   645
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   4035
         TabIndex        =   21
         Top             =   1500
         Width           =   645
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   2475
         TabIndex        =   16
         Top             =   405
         Width           =   645
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   2475
         TabIndex        =   15
         Top             =   780
         Width           =   645
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   2490
         TabIndex        =   14
         Top             =   1140
         Width           =   645
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   2490
         TabIndex        =   13
         Top             =   1500
         Width           =   645
      End
      Begin VB.CheckBox Check1 
         Caption         =   "统计同基站其他天线"
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   2055
         Width           =   1965
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   3
         Left            =   900
         TabIndex        =   9
         Top             =   1500
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   2
         Left            =   900
         TabIndex        =   7
         Top             =   1140
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   1
         Left            =   900
         TabIndex        =   5
         Top             =   780
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   0
         Left            =   900
         TabIndex        =   4
         Top             =   405
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CI_3:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   11
         Left            =   3510
         TabIndex        =   28
         Top             =   435
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ARFCN_3:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   10
         Left            =   3240
         TabIndex        =   27
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "BSIC_3:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   9
         Left            =   3345
         TabIndex        =   26
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LAC_3:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   3435
         TabIndex        =   25
         Top             =   1530
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CI_2:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   7
         Left            =   1965
         TabIndex        =   20
         Top             =   435
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ARFCN_2:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   6
         Left            =   1695
         TabIndex        =   19
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "BSIC_2:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   5
         Left            =   1800
         TabIndex        =   18
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LAC_2:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   1890
         TabIndex        =   17
         Top             =   1530
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LAC_1:"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   315
         TabIndex        =   10
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "BSIC_1:"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   225
         TabIndex        =   8
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ARFCN_1:"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   825
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CI_1:"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   3
         Top             =   450
         Width           =   450
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "·选择地图上的天线或基站可以自动获得所需参数"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   465
      TabIndex        =   12
      Top             =   2820
      Width           =   3960
   End
End
Attribute VB_Name = "frmNcellA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Dim i As Integer

    On Error Resume Next
    If Check1.Value = 1 Then
        For i = 0 To 3
            Text2(i).Enabled = True
            Text3(i).Enabled = True
        Next
        For i = 4 To 11
            Label2(i).Enabled = True
        Next
        If Text2(0).Text = "" And Text2(1).Text = "" Then
            'Text2(0).Text =format(val(Text2(0).Text)+1)
            Text2(2).Text = Text1(2).Text
            Text2(3).Text = Text1(3).Text
            'Text3(0).Text =format(val(Text2(0).Text)+2)
            Text3(2).Text = Text1(2).Text
            Text3(3).Text = Text1(3).Text
        End If
    Else
        For i = 0 To 3
            Text2(i).Enabled = False
            Text3(i).Enabled = False
        Next
        For i = 4 To 11
            Label2(i).Enabled = False
        Next
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    MyNRSelCellCI = Text1(0).Text
    MyNRSelCellBcch = Val(Text1(1).Text)
    MyNRSelCellBsic = Val(Text1(2).Text)
    MyNRSelCellLac = Text1(3).Text
    If Check1.Value = 1 Then
        MyNRSelCellCI_2 = Text2(0).Text
        MyNRSelCellBcch_2 = Val(Text2(1).Text)
        MyNRSelCellBsic_2 = Val(Text2(2).Text)
        MyNRSelCellCI_3 = Text3(0).Text
        MyNRSelCellBcch_3 = Val(Text3(1).Text)
        MyNRSelCellBsic_3 = Val(Text3(2).Text)
    Else
        MyNRSelCellCI_2 = ""
        MyNRSelCellBcch_2 = 0
        MyNRSelCellBsic_2 = 0
        MyNRSelCellCI_3 = ""
        MyNRSelCellBcch_3 = 0
        MyNRSelCellBsic_3 = 0
    End If

    If MyNRSelCellCI = "" And MyNRSelCellBcch = 0 And MyNRSelCellBsic = 0 And MyNRSelCellLac = "" Then
        MsgBox "没有输入有效参数，无法进行统计", 64, "提示"
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    MyNRSelCellCI = ""
    MyNRSelCellBcch = 0
    MyNRSelCellBsic = 0
    MyNRSelCellLac = ""
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
        
    On Error Resume Next
    If MyNRSelCellCI <> "" And MyNRSelCellBcch <> 0 Then
        Text1(0).Text = MyNRSelCellCI
        Text1(1).Text = Format(MyNRSelCellBcch)
        Text1(2).Text = Format(MyNRSelCellBsic)
        Text1(3).Text = MyNRSelCellLac
        If MyNRSelCellBcch_2 > 0 Then
            Check1.Value = 1
            For i = 0 To 3
                Text2(i).Enabled = True
                Text3(i).Enabled = True
            Next
            For i = 4 To 11
                Label2(i).Enabled = True
            Next
            Text2(0).Text = MyNRSelCellCI_2
            Text2(1).Text = MyNRSelCellBcch_2
            Text2(2).Text = MyNRSelCellBsic_2
            Text2(3).Text = MyNRSelCellLac
            If MyNRSelCellBcch_3 > 0 Then
                Text3(0).Text = MyNRSelCellCI_3
                Text3(1).Text = MyNRSelCellBcch_3
                Text3(2).Text = MyNRSelCellBsic_3
                Text3(3).Text = MyNRSelCellLac
            Else
                Text3(0).Text = ""
                Text3(1).Text = ""
                Text3(2).Text = ""
                Text3(3).Text = ""
            End If
        End If
    End If
    
End Sub
