VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form NcellFrm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ncell 定义"
   ClientHeight    =   5535
   ClientLeft      =   2355
   ClientTop       =   1155
   ClientWidth     =   5820
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Nce1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   5820
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   64
      Top             =   5115
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   741
      ButtonWidth     =   1111
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "下一个"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "前一个"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "最后一个"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "第一个"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "增加"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "删除"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "自动下载"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "返回"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "NCELL"
      Height          =   3420
      Left            =   300
      TabIndex        =   9
      Top             =   1470
      Width           =   5235
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   660
         TabIndex        =   41
         Text            =   " "
         Top             =   675
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   660
         TabIndex        =   40
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   3
         Left            =   660
         TabIndex        =   39
         Top             =   1650
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   4
         Left            =   660
         TabIndex        =   38
         Top             =   1980
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   5
         Left            =   660
         TabIndex        =   37
         Top             =   2295
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   6
         Left            =   660
         TabIndex        =   36
         Top             =   2625
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   7
         Left            =   660
         TabIndex        =   35
         Top             =   2955
         Width           =   600
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   1470
         TabIndex        =   34
         Top             =   1005
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   1470
         TabIndex        =   33
         Top             =   1320
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   3
         Left            =   1470
         TabIndex        =   32
         Top             =   1650
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   4
         Left            =   1470
         TabIndex        =   31
         Top             =   1980
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   5
         Left            =   1470
         TabIndex        =   30
         Top             =   2310
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   6
         Left            =   1470
         TabIndex        =   29
         Top             =   2640
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   7
         Left            =   1470
         TabIndex        =   28
         Top             =   2970
         Width           =   840
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   8
         Left            =   3135
         TabIndex        =   27
         Top             =   690
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   9
         Left            =   3135
         TabIndex        =   26
         Top             =   1035
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   10
         Left            =   3135
         TabIndex        =   25
         Top             =   1365
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   12
         Left            =   3135
         TabIndex        =   24
         Top             =   2010
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   13
         Left            =   3135
         TabIndex        =   23
         Top             =   2340
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   15
         Left            =   3135
         TabIndex        =   22
         Top             =   2985
         Width           =   600
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   8
         Left            =   3945
         TabIndex        =   21
         Top             =   690
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   9
         Left            =   3945
         TabIndex        =   20
         Top             =   1020
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   10
         Left            =   3945
         TabIndex        =   19
         Top             =   1350
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   11
         Left            =   3945
         TabIndex        =   18
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   12
         Left            =   3945
         TabIndex        =   17
         Top             =   1995
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   13
         Left            =   3945
         TabIndex        =   16
         Top             =   2325
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   14
         Left            =   3945
         TabIndex        =   15
         Top             =   2655
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   15
         Left            =   3945
         TabIndex        =   14
         Top             =   2985
         Width           =   840
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   1470
         TabIndex        =   13
         Top             =   675
         Width           =   840
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   660
         TabIndex        =   12
         Top             =   1005
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   14
         Left            =   3135
         TabIndex        =   11
         Top             =   2670
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   11
         Left            =   3135
         TabIndex        =   10
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CI"
         Height          =   180
         Index           =   9
         Left            =   4140
         TabIndex        =   63
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BCCH "
         Height          =   270
         Index           =   8
         Left            =   3135
         TabIndex        =   62
         Top             =   375
         Width           =   705
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   270
         Index           =   7
         Left            =   2730
         TabIndex        =   61
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CI"
         Height          =   180
         Index           =   6
         Left            =   1665
         TabIndex        =   60
         Top             =   345
         Width           =   180
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BCCH "
         Height          =   270
         Index           =   5
         Left            =   660
         TabIndex        =   59
         Top             =   345
         Width           =   705
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   270
         Index           =   4
         Left            =   285
         TabIndex        =   58
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         Height          =   240
         Index           =   0
         Left            =   345
         TabIndex        =   57
         Top             =   690
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
         Height          =   270
         Index           =   1
         Left            =   345
         TabIndex        =   56
         Top             =   990
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   345
         TabIndex        =   55
         Top             =   1335
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4"
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   54
         Top             =   1650
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "5"
         Height          =   240
         Index           =   4
         Left            =   360
         TabIndex        =   53
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "6"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   52
         Top             =   2310
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "7"
         Height          =   240
         Index           =   6
         Left            =   360
         TabIndex        =   51
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "8"
         Height          =   240
         Index           =   7
         Left            =   360
         TabIndex        =   50
         Top             =   3000
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "9"
         Height          =   240
         Index           =   8
         Left            =   2775
         TabIndex        =   49
         Top             =   705
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "10"
         Height          =   225
         Index           =   9
         Left            =   2730
         TabIndex        =   48
         Top             =   1020
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "11"
         Height          =   210
         Index           =   10
         Left            =   2745
         TabIndex        =   47
         Top             =   1350
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "12"
         Height          =   255
         Index           =   11
         Left            =   2730
         TabIndex        =   46
         Top             =   1650
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "13"
         Height          =   255
         Index           =   12
         Left            =   2730
         TabIndex        =   45
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "14"
         Height          =   240
         Index           =   13
         Left            =   2730
         TabIndex        =   44
         Top             =   2325
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "15"
         Height          =   240
         Index           =   14
         Left            =   2730
         TabIndex        =   43
         Top             =   2670
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "16"
         Height          =   225
         Index           =   15
         Left            =   2730
         TabIndex        =   42
         Top             =   3015
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CELL"
      Height          =   1305
      Left            =   315
      TabIndex        =   0
      Top             =   105
      Width           =   5220
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1215
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   330
         Width           =   1590
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   1215
         TabIndex        =   3
         Top             =   825
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   3810
         TabIndex        =   2
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   3810
         TabIndex        =   1
         Top             =   810
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "基站名称:"
         Height          =   180
         Index           =   0
         Left            =   315
         TabIndex        =   8
         Top             =   390
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "基站序号:"
         Height          =   180
         Index           =   1
         Left            =   330
         TabIndex        =   7
         Top             =   855
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CI:"
         Height          =   180
         Index           =   3
         Left            =   3450
         TabIndex        =   6
         Top             =   840
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "BCCH:"
         Height          =   180
         Index           =   2
         Left            =   3270
         TabIndex        =   5
         Top             =   390
         Width           =   450
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5505
      Top             =   6045
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   26
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Nce1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Nce1.frx":08E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Nce1.frx":0EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Nce1.frx":14D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Nce1.frx":1AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Nce1.frx":20BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Nce1.frx":26B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Nce1.frx":2824
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Nce1.frx":2DFE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "NcellFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pp As Integer
Dim n_cha As Boolean  'Text Change
Dim add_frag As Boolean
Dim list_addname As Boolean
Dim moto_type As Boolean

Sub to_face()
    Dim i As Integer
    On Error Resume Next
    If moto_type = True Then
       For i = 0 To 15
           j = 5 + i * 4
           k = 6 + i * 4
           Text2(i).Text = mapinfo.eval("ncell.col" & j)
           Text3(i).Text = mapinfo.eval("ncell.col" & k)
       Next i
    Else
       For i = 0 To 15
           j = 5 + i * 4
           k = 8 + i * 4
           Text2(i).Text = mapinfo.eval("ncell.col" & j)
           Text3(i).Text = mapinfo.eval("ncell.col" & k)
       Next i
    End If
    Text1(0).Text = mapinfo.eval("ncell.col2")
    Text1(1).Text = mapinfo.eval("ncell.col4")
    Text1(2).Text = mapinfo.eval("ncell.ci")
End Sub

Sub to_table()
    Dim i As Integer
    Dim tt(1 To 16) As String, TT1(1 To 16) As String, tt2(1 To 16) As String
    Dim msg As String
    
    On Error Resume Next
    For i = 0 To 15
      tt(i + 1) = Trim(Text2(i).Text)
      TT1(i + 1) = Trim(Text3(i).Text)
    Next i
    msg = "UPDATE  ncell set  "
    msg = msg + "COL1=" + Chr(34) + Trim(Combo1.Text) + Chr(34) + ",col2=" + Chr(34) + Trim(Text1(0).Text) + Chr(34) + ",col3=" + Chr(34) + Trim(Text1(2).Text) + Chr(34) + ",col4=" + str(Val(Text1(1).Text)) + ","
    If moto_type = True Then
       For i = 0 To 15
           j = 5 + i * 4
           k = 6 + i * 4
           msg = msg + "COL" & j & " = " + Trim(str(Val(tt(i + 1)))) + ","
           If i = 15 Then
              msg = msg + "COL" & k & " = " + Trim(str(Val(TT1(i + 1))))
           Else
              msg = msg + "COL" & k & " = " + Trim(str(Val(TT1(i + 1)))) + ","
           End If
       Next i
    Else
       For i = 0 To 15
           j = 5 + i * 4
           k = 8 + i * 4
           msg = msg + "COL" & j & " = " + Trim(str(Val(tt(i + 1)))) + ","
           If i = 15 Then
              msg = msg + "COL" & k & " = " + Chr(34) + TT1(i + 1) + Chr(34)
           Else
              msg = msg + "COL" & k & " = " + Chr(34) + TT1(i + 1) + Chr(34) + ","
           End If
       Next i
    End If
    msg = msg + " WHERE ROWID =  " & pp
    mapinfo.do msg
    mapinfo.do "commit table ncell"

End Sub


Sub Combo1_Click()
    On Error Resume Next
    comb_click = True
    If list_addname = True Then
       Exit Sub
    End If
    If add_frag = True Then
       add_frag = False
       n_cha = False
       reco = 1
       mapinfo.do "fetch rec " & reco & "from ncell"
       pp = reco
       to_face
       Exit Sub
    End If
    If n_cha = True Then
       to_table
       n_cha = False
    End If
    reco = Val(Combo1.ListIndex) + 1
    mapinfo.do "fetch rec " & reco & "from ncell"
    pp = reco
    Call to_face

End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    n_cha = True
End Sub

Private Sub Form_Load()
    Dim all
    Dim MyRecord As Record
    Dim i As Integer
    
    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\map\ncell.tab"
    If UCase$(dir(Gsm_FileName, 0)) <> "NCELL.TAB" Then
       'MsgBox " NCELL.TAB 不存在！", 64, "提示"
       'Unload Me
       'Exit Sub
       Gsm_FileName = Gsm_Path + "\ncell.dbf"
       Gsm_File2 = Gsm_Path + "\map\ncell.dbf"
       FileCopy Gsm_FileName, Gsm_File2
       mapinfo.do "Register Table " + Chr(34) + Gsm_File2 + Chr(34) + " Type " + Chr(34) + "DBF" + Chr(34) + " Into " + Chr(34) + Gsm_Path + "\map\ncell.tab" + Chr(34)
    End If
    n_cha = False
    add_frag = False
    list_addname = False
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    Open Gsm_FileName For Binary As #1
    Get #1, 1, MyRecord  ' Read third record.
    Close #1
    If Val(MyRecord.exchange) = 1 Or Val(MyRecord.exchange) = 4 Then
       moto_type = True
    Else
       moto_type = False
    End If
    If moto_type = False Then
       Label1(6).Caption = "Base_No"
       Label1(9).Caption = "Base_No"
       For i = 0 To 15
           Text3(i).Width = 950
       Next
    End If
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\map\ncell" + Chr(34)
    mapinfo.do "fetch First from ncell"
    
    all = Val(mapinfo.eval("TABLEINFO(ncell, 8)"))
    For i = 1 To all
        add_name = mapinfo.eval("ncell.bs_name")
        Combo1.AddItem add_name
        mapinfo.do "fetch next from ncell"
    Next
    mapinfo.do "fetch first from ncell"
    If all > 0 Then
       Combo1.ListIndex = 0
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mapinfo.do "close table ncell"

End Sub

Private Sub SS_add_Click()
    Dim i As Integer
    Dim row
    On Error Resume Next
    If add_frag = True Then
       add_frag = False
       n_cha = False
       If Trim(Combo1.Text) <> "" Then
          Combo1.AddItem Trim(Combo1.Text)
          mapinfo.do "insert into ncell (col1) values(0)"
          to_table
       End If
    End If
    If n_cha = True Then
       to_table
       n_cha = False
    End If
    row = mapinfo.eval("TABLEINFO(ncell, 8)")
    pp = row + 1
    
    add_frag = True
    For i = 0 To 15
        Text2(i).Text = ""
        Text3(i).Text = ""
    Next
    For i = 0 To 2
        Text1(i).Text = ""
    Next
    list_addname = True
    Combo1.Text = ""
    Combo1.SetFocus
    list_addname = False
End Sub

Private Sub SS_del_Click()
    Dim all
    Dim i As Integer
    On Error Resume Next
    If add_frag = True Then
       add_frag = False
       n_cha = False
       Combo1.ListIndex = Combo1.ListCount - 1
       Exit Sub
    End If
    If (MsgBox("当前数据将被删除，确定吗？", 33, "提示")) = 1 Then
       mapinfo.do "delete from ncell where rowid=" & pp
       mapinfo.do "commit table ncell"
       mapinfo.do "pack table ncell Graphic Data  Data  Interactive  "
       mapinfo.do "fetch first from ncell"
       all = mapinfo.eval("TABLEINFO(ncell, 8)")
       Combo1.Clear
       For i = 1 To all
           add_name = mapinfo.eval("ncell.bs_name")
           Combo1.AddItem add_name
           mapinfo.do "fetch next from ncell"
       Next
       mapinfo.do "fetch first from ncell"
       Combo1.Text = Combo1.List(0)
       Combo1.ListIndex = 0
    End If
End Sub

Private Sub SS_first_Click()
    On Error Resume Next
    If add_frag = True Then
       add_frag = False
       n_cha = False
       If Trim(Combo1.Text) <> "" Then
          Combo1.AddItem Trim(Combo1.Text)
          mapinfo.do "insert into ncell (col1) values(0)"
          to_table
       End If
    End If
    If n_cha = True Then
       Call to_table
       n_cha = False
    End If
    Combo1.ListIndex = 0
    
End Sub

Private Sub SS_last_Click()
    On Error Resume Next
    If add_frag = True Then
       add_frag = False
       n_cha = False
       If Trim(Combo1.Text) <> "" Then
          Combo1.AddItem Trim(Combo1.Text)
          mapinfo.do "insert into ncell (col1) values(0)"
          to_table
       End If
    End If
    If n_cha = True Then
       Call to_table
       n_cha = False
    End If
    Combo1.ListIndex = Combo1.ListCount - 1
    
End Sub


Private Sub SS_next_Click()
    On Error Resume Next
    If add_frag = True Then
       add_frag = False
       n_cha = False
       If Trim(Combo1.Text) <> "" Then
          Combo1.AddItem Trim(Combo1.Text)
          mapinfo.do "insert into ncell (col1) values(0)"
          to_table
       End If
    End If
    If n_cha = True Then
       Call to_table
       n_cha = False
    End If
    If (Combo1.ListIndex + 1) < Combo1.ListCount Then
       Combo1.ListIndex = Combo1.ListIndex + 1
    End If
    
End Sub

Private Sub SS_prev_Click()
    On Error Resume Next
    If add_frag = True Then
       add_frag = False
       n_cha = False
       If Trim(Combo1.Text) <> "" Then
          Combo1.AddItem Trim(Combo1.Text)
          mapinfo.do "insert into ncell (col1) values(0)"
          to_table
       End If
    End If
    If n_cha = True Then
       Call to_table
       n_cha = False
    End If
    If Combo1.ListIndex > 0 Then
        Combo1.ListIndex = Combo1.ListIndex - 1
    End If
        
End Sub


Private Sub SS_return_Click()
    ncell_tip = 0
    If add_frag = True Then
       add_frag = False
       n_cha = False
       If Trim(Combo1.Text) <> "" Then
          Combo1.AddItem Trim(Combo1.Text)
          mapinfo.do "insert into ncell (col1) values(0)"
          to_table
       End If
    End If
    If n_cha = True Then
       to_table
       n_cha = False
    End If
    Unload Me
End Sub


Private Sub SS_update_Click()
    Dim all As Integer
    Dim MyRecord As Record
    Dim Convert_Flag As Boolean
    
    On Error Resume Next
    ncell_tip = 0
    If add_frag = True Then
       add_frag = False
       n_cha = False
       If Trim(Combo1.Text) <> "" Then
          Combo1.AddItem Trim(Combo1.Text)
          mapinfo.do "insert into ncell (col1) values(0)"
          to_table
       End If
    End If
    Menu_Flag = 2303
    Convert_Flag = False
    mapinfo.do "close table  ncell"
   
'8************************************************
    MDIMain.StatusBar.Panels(1).Text = "   数据转换"
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    Open Gsm_FileName For Binary As #1
    Get #1, 1, MyRecord  ' Read third record.
    Close #1
    If Val(MyRecord.exchange) = 0 Or Val(MyRecord.exchange) = 1 Or Val(MyRecord.exchange) = 4 Then
       For i = 1 To 50
           convert_filename(i) = ""
       Next
open_again:
       MDIMain.FileDialog.DialogTitle = "数据转换文件选择"
       If Val(MyRecord.exchange) = 0 Then
          MDIMain.FileDialog.Filter = "All Files|*.*"
          MDIMain.FileDialog.DefaultExt = ""
          MDIMain.FileDialog.Flags = &H200 Or &H80000
       Else
          MDIMain.FileDialog.Filter = "*.txt Files|*.TXT|All Files|*.*"
          MDIMain.FileDialog.DefaultExt = "*.TXT"
          If Val(MyRecord.exchange) = 4 Then
             MDIMain.FileDialog.Flags = &H200 Or &H80000
          Else
             MDIMain.FileDialog.Flags = &H80000
          End If
       End If
       MDIMain.FileDialog.InitDir = Gsm_Path
       MDIMain.FileDialog.ShowOpen
       buff = Trim(MDIMain.FileDialog.filename)
       If buff = "" Then
          GoTo complete
       End If
       finds = InStr(buff, Chr(0))
       If finds > 0 Then
          mypath = Left(buff, finds - 1) + "\"
          buff = Trim(Right(buff, Len(buff) - finds))
          finds = InStr(buff, Chr(0))
          i = 1
          Do While finds > 0
             convert_filename(i) = mypath + Left(buff, finds - 1)
             buff = Trim(Right(buff, Len(buff) - finds))
             finds = InStr(buff, Chr(0))
             i = i + 1
          Loop
          convert_filename(i) = mypath + buff
       Else
          convert_filename(1) = buff
       End If
       MDIMain.FileDialog.filename = ""
       If dir(convert_filename(1)) = "" Then
          GoTo err_exit
       End If
       If Val(MyRecord.exchange) = 0 Then
'          If Right(convert_filename(1), 5) <> "rlnrp" Then
           If InStr(UCase(convert_filename(1)), "RLNRP") = 0 Then
             i = MsgBox("文件名应为RLNRP！", 48, "打开文件")
             GoTo open_again
          End If
       End If
       Screen.MousePointer = 11
       Data_Convert.Show 1
       Screen.MousePointer = 0
       Convert_Flag = True
       GoTo complete
err_exit:
       i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
       GoTo open_again
    Else
       MsgBox "该交换机类型的数据转换暂未挂接!", 64, "提示"
    End If
complete:
    MDIMain.StatusBar.Panels(1).Text = " "
'8*************************************************
    mapinfo.do "Register Table  " + " " + Chr(34) + Gsm_Path + "\map\ncell.dbf" + Chr(34) + "Type " + " " + Chr(34) + "DBF" + Chr(34) + "Into  " + Chr(34) + Gsm_Path + "\map\ncell.tab" + Chr(34)
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\map\ncell" + Chr(34)
    mapinfo.do "fetch First from ncell"
    
    If Convert_Flag = False Then
       GoTo no_convert
    End If
    
    Combo1.Clear
    all = Val(mapinfo.eval("TABLEINFO(ncell, 8)"))
    For i = 1 To all
        add_name = mapinfo.eval("ncell.bs_name")
        Combo1.AddItem add_name
        mapinfo.do "fetch next from ncell"
    Next
    mapinfo.do "fetch first from ncell"
    
no_convert:
    Combo1.ListIndex = 0
    Screen.MousePointer = 0
End Sub

Private Sub Text2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    n_cha = True
End Sub

Private Sub Text3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    n_cha = True
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    n_cha = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error Resume Next
    Select Case Button.Index
       Case 2
            SS_next_Click
       Case 3
            SS_prev_Click
       Case 4
            SS_last_Click
       Case 5
            SS_first_Click
       Case 7
            SS_add_Click
       Case 8
            SS_del_Click
       Case 10
            SS_update_Click
       Case 12
            SS_return_Click
    End Select
End Sub
