VERSION 5.00
Begin VB.Form frmCoverChoice 
   Caption         =   "条件选择"
   ClientHeight    =   2325
   ClientLeft      =   4035
   ClientTop       =   4860
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   4815
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
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
      Left            =   3375
      TabIndex        =   7
      Top             =   1425
      Width           =   1080
   End
   Begin VB.CommandButton SBSCancel 
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
      Left            =   3375
      TabIndex        =   6
      Top             =   1815
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "门限"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   180
      TabIndex        =   3
      Top             =   165
      Width           =   3030
      Begin VB.TextBox RxLevValue 
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
         Left            =   1260
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "17"
         Top             =   345
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cell RxLev:"
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
         Index           =   6
         Left            =   210
         TabIndex        =   5
         Top             =   345
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Full/Sub选择"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   180
      TabIndex        =   0
      Top             =   1335
      Width           =   1785
      Begin VB.OptionButton Option5 
         Caption         =   "Sub"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   330
         Width           =   570
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Full"
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
         Left            =   210
         TabIndex        =   1
         Top             =   345
         Value           =   -1  'True
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmCoverChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

