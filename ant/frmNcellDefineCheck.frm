VERSION 5.00
Begin VB.Form frmNcellDefineCheck 
   Caption         =   "相邻小区定义检查"
   ClientHeight    =   4305
   ClientLeft      =   4200
   ClientTop       =   4245
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNcellDefineCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5070
   Begin VB.Frame Frame1 
      Caption         =   "小区选择"
      Height          =   2835
      Left            =   345
      TabIndex        =   2
      Top             =   150
      Width           =   3630
      Begin VB.ComboBox Combo2 
         DataField       =   " "
         DataSource      =   " "
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1920
         Width           =   1395
      End
      Begin VB.CheckBox Cell_1 
         Caption         =   "小区1"
         Height          =   240
         Left            =   675
         TabIndex        =   10
         Top             =   1185
         Value           =   1  'Checked
         Width           =   840
      End
      Begin VB.CheckBox Cell_2 
         Caption         =   "小区2"
         Height          =   240
         Left            =   1575
         TabIndex        =   9
         Top             =   1185
         Width           =   840
      End
      Begin VB.CheckBox Cell_3 
         Caption         =   "小区3"
         Height          =   240
         Left            =   2490
         TabIndex        =   8
         Top             =   1185
         Width           =   840
      End
      Begin VB.ComboBox Combo1 
         DataField       =   " "
         DataSource      =   " "
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Caption         =   "全网"
         Height          =   315
         Index           =   2
         Left            =   375
         TabIndex        =   5
         Top             =   2340
         Width           =   765
      End
      Begin VB.OptionButton Option1 
         Caption         =   "指定LAC"
         Height          =   315
         Index           =   1
         Left            =   375
         TabIndex        =   4
         Top             =   1560
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "指定基站"
         Height          =   315
         Index           =   0
         Left            =   375
         TabIndex        =   3
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "LAC选择："
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   1
         Left            =   675
         TabIndex        =   12
         Top             =   1995
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "基站选择："
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   0
         Left            =   675
         TabIndex        =   7
         Top             =   795
         Width           =   900
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      DragIcon        =   "frmNcellDefineCheck.frx":000C
      Height          =   320
      Left            =   2175
      TabIndex        =   1
      Top             =   3720
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      DragIcon        =   "frmNcellDefineCheck.frx":015E
      Height          =   320
      Left            =   975
      TabIndex        =   0
      Top             =   3720
      Width           =   1080
   End
End
Attribute VB_Name = "frmNcellDefineCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

