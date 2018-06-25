VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form New_Base 
   Caption         =   "网络资源管理"
   ClientHeight    =   6990
   ClientLeft      =   720
   ClientTop       =   1335
   ClientWidth     =   10335
   Icon            =   "Newbase.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6990
   ScaleWidth      =   10335
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   18
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "增站"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "增区"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "删除"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "排序"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "统计"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "换网"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "下载"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "生成"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "升级"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "导出"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "导入"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "退出"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   690
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6630
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "NCELL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4710
      TabIndex        =   131
      Top             =   6420
      Visible         =   0   'False
      Width           =   4080
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         MaxLength       =   64
         TabIndex        =   50
         Top             =   0
         Width           =   4065
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "NCELL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2325
      TabIndex        =   130
      Top             =   6420
      Visible         =   0   'False
      Width           =   8025
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         MaxLength       =   10
         TabIndex        =   52
         Top             =   255
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         MaxLength       =   10
         TabIndex        =   51
         Top             =   0
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   6930
         MaxLength       =   10
         TabIndex        =   66
         Top             =   255
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   6930
         MaxLength       =   10
         TabIndex        =   65
         Top             =   0
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   5940
         MaxLength       =   10
         TabIndex        =   64
         Top             =   255
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   5940
         MaxLength       =   10
         TabIndex        =   63
         Top             =   0
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   4950
         MaxLength       =   10
         TabIndex        =   62
         Top             =   255
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   4950
         MaxLength       =   10
         TabIndex        =   61
         Top             =   0
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   60
         Top             =   255
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   59
         Top             =   0
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2970
         MaxLength       =   10
         TabIndex        =   58
         Top             =   255
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2970
         MaxLength       =   10
         TabIndex        =   57
         Top             =   0
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   56
         Top             =   255
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   55
         Top             =   0
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   990
         MaxLength       =   10
         TabIndex        =   54
         Top             =   255
         Width           =   990
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   990
         MaxLength       =   10
         TabIndex        =   53
         Top             =   0
         Width           =   990
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4050
      Left            =   120
      TabIndex        =   3
      Top             =   2550
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   7144
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "基站名"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "信道数"
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "邻小区数"
         Object.Width           =   1076
      EndProperty
   End
   Begin VB.Frame Frame6 
      Caption         =   "运营网络"
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
      Height          =   1320
      Left            =   150
      TabIndex        =   76
      Top             =   720
      Width           =   3150
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   1605
         TabIndex        =   2
         Top             =   780
         Width           =   390
      End
      Begin VB.TextBox Text5 
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
         Index           =   0
         Left            =   585
         TabIndex        =   0
         Top             =   420
         Width           =   1035
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   585
         TabIndex        =   1
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "次设置"
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
         Index           =   5
         Left            =   2070
         TabIndex        =   108
         Top             =   825
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "期工程"
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
         Left            =   1020
         TabIndex        =   107
         Top             =   825
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "GSM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   1680
         TabIndex        =   77
         Top             =   465
         Width           =   345
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "小区"
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
      Height          =   4035
      Left            =   3405
      TabIndex        =   71
      Top             =   2550
      Width           =   6795
      Begin VB.Frame Frame3 
         Caption         =   "第三小区"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Index           =   2
         Left            =   4500
         TabIndex        =   97
         Top             =   240
         Width           =   2160
         Begin VB.TextBox Text3 
            DataField       =   "CI"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1020
            MaxLength       =   5
            TabIndex        =   40
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox Text3 
            DataField       =   "CELL_NAME"
            DataSource      =   "Data3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   39
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1020
            MaxLength       =   5
            TabIndex        =   41
            Top             =   870
            Width           =   555
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   42
            Top             =   1155
            Width           =   555
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   43
            Top             =   1440
            Width           =   555
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   44
            Top             =   1725
            Width           =   555
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   45
            Top             =   2010
            Width           =   555
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   46
            Top             =   2295
            Width           =   555
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1020
            MaxLength       =   2
            TabIndex        =   47
            Top             =   2580
            Width           =   555
         End
         Begin VB.CommandButton Command2 
            Caption         =   "NCELL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1020
            TabIndex        =   49
            Top             =   3210
            Width           =   780
         End
         Begin VB.CommandButton Command1 
            Caption         =   "TCH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1020
            TabIndex        =   48
            Top             =   2880
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "dBm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   53
            Left            =   1620
            TabIndex        =   127
            Top             =   2595
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "米"
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
            Index           =   52
            Left            =   1635
            TabIndex        =   126
            Top             =   2325
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "度"
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
            Index           =   47
            Left            =   1635
            TabIndex        =   121
            Top             =   2025
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "度"
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
            Index           =   46
            Left            =   1635
            TabIndex        =   120
            Top             =   1770
            Width           =   180
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "CI："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   34
            Left            =   705
            TabIndex        =   106
            Top             =   615
            Width           =   300
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "BS_NO："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   33
            Left            =   315
            TabIndex        =   105
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "LAC："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   32
            Left            =   525
            TabIndex        =   104
            Top             =   915
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "BCCH："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   31
            Left            =   420
            TabIndex        =   103
            Top             =   1200
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "BSIC："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   30
            Left            =   495
            TabIndex        =   102
            Top             =   1485
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "方向角："
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
            Index           =   29
            Left            =   285
            TabIndex        =   101
            Top             =   1770
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "下倾角："
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
            Index           =   28
            Left            =   285
            TabIndex        =   100
            Top             =   2055
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "天线高度："
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
            Index           =   27
            Left            =   105
            TabIndex        =   99
            Top             =   2340
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "发射功率："
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
            Index           =   26
            Left            =   105
            TabIndex        =   98
            Top             =   2625
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "第二小区"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Index           =   1
         Left            =   2310
         TabIndex        =   87
         Top             =   240
         Width           =   2160
         Begin VB.TextBox Text2 
            DataField       =   "CI"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1020
            MaxLength       =   5
            TabIndex        =   29
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox Text2 
            DataField       =   "CELL_NAME"
            DataSource      =   "Data2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   1005
            MaxLength       =   10
            TabIndex        =   28
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1020
            MaxLength       =   5
            TabIndex        =   30
            Top             =   870
            Width           =   555
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   31
            Top             =   1155
            Width           =   555
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   32
            Top             =   1440
            Width           =   555
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   33
            Top             =   1725
            Width           =   555
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   34
            Top             =   2010
            Width           =   555
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   35
            Top             =   2295
            Width           =   555
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1020
            MaxLength       =   2
            TabIndex        =   36
            Top             =   2580
            Width           =   555
         End
         Begin VB.CommandButton Command2 
            Caption         =   "NCELL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1020
            TabIndex        =   38
            Top             =   3210
            Width           =   780
         End
         Begin VB.CommandButton Command1 
            Caption         =   "TCH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1020
            TabIndex        =   37
            Top             =   2880
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "dBm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   51
            Left            =   1635
            TabIndex        =   125
            Top             =   2595
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "米"
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
            Index           =   50
            Left            =   1650
            TabIndex        =   124
            Top             =   2325
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "度"
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
            Index           =   45
            Left            =   1650
            TabIndex        =   119
            Top             =   2040
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "度"
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
            Index           =   44
            Left            =   1650
            TabIndex        =   118
            Top             =   1770
            Width           =   180
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "CI："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   25
            Left            =   705
            TabIndex        =   96
            Top             =   615
            Width           =   300
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "BS_NO："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   24
            Left            =   330
            TabIndex        =   95
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "LAC："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   23
            Left            =   540
            TabIndex        =   94
            Top             =   915
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "BCCH："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   21
            Left            =   435
            TabIndex        =   93
            Top             =   1200
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "BSIC："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   20
            Left            =   510
            TabIndex        =   92
            Top             =   1485
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "方向角："
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
            Index           =   11
            Left            =   285
            TabIndex        =   91
            Top             =   1770
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "下倾角："
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
            Index           =   10
            Left            =   285
            TabIndex        =   90
            Top             =   2055
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "天线高度："
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
            Index           =   9
            Left            =   105
            TabIndex        =   89
            Top             =   2340
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "发射功率："
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
            Index           =   8
            Left            =   105
            TabIndex        =   88
            Top             =   2625
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "第一小区"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Index           =   0
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   2160
         Begin VB.CommandButton Command1 
            Caption         =   "TCH"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1020
            TabIndex        =   26
            Top             =   2880
            Width           =   780
         End
         Begin VB.CommandButton Command2 
            Caption         =   "NCELL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1020
            TabIndex        =   27
            Top             =   3210
            Width           =   780
         End
         Begin VB.TextBox Text1 
            DataField       =   "MAX_TX_BTS"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1020
            MaxLength       =   2
            TabIndex        =   25
            Top             =   2580
            Width           =   555
         End
         Begin VB.TextBox Text1 
            DataField       =   "ANT_HEIGH"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   24
            Top             =   2295
            Width           =   555
         End
         Begin VB.TextBox Text1 
            DataField       =   "DOWNTILT"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   23
            Top             =   2010
            Width           =   555
         End
         Begin VB.TextBox Text1 
            DataField       =   "BEARING"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   22
            Top             =   1725
            Width           =   555
         End
         Begin VB.TextBox Text1 
            DataField       =   "BSIC"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   21
            Top             =   1440
            Width           =   555
         End
         Begin VB.TextBox Text1 
            DataField       =   "ARFCN"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   20
            Top             =   1155
            Width           =   555
         End
         Begin VB.TextBox Text1 
            DataField       =   "LAC"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1020
            MaxLength       =   5
            TabIndex        =   19
            Top             =   870
            Width           =   555
         End
         Begin VB.TextBox Text1 
            DataField       =   "CELL_NAME"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text1 
            DataField       =   "CI"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1020
            MaxLength       =   5
            TabIndex        =   18
            Top             =   570
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "dBm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   49
            Left            =   1620
            TabIndex        =   123
            Top             =   2595
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "米"
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
            Index           =   48
            Left            =   1635
            TabIndex        =   122
            Top             =   2325
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "度"
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
            Index           =   43
            Left            =   1635
            TabIndex        =   117
            Top             =   2055
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "度"
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
            Index           =   42
            Left            =   1635
            TabIndex        =   116
            Top             =   1755
            Width           =   180
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "发射功率："
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
            Index           =   22
            Left            =   105
            TabIndex        =   86
            Top             =   2625
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "天线高度："
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
            Index           =   19
            Left            =   105
            TabIndex        =   85
            Top             =   2340
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "下倾角："
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
            Index           =   18
            Left            =   285
            TabIndex        =   84
            Top             =   2055
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "方向角："
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
            Index           =   17
            Left            =   285
            TabIndex        =   83
            Top             =   1770
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "BSIC："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   16
            Left            =   510
            TabIndex        =   82
            Top             =   1485
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "BCCH："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   15
            Left            =   435
            TabIndex        =   81
            Top             =   1200
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "LAC："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   14
            Left            =   540
            TabIndex        =   80
            Top             =   915
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "BS_NO："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   13
            Left            =   330
            TabIndex        =   79
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "CI："
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   12
            Left            =   720
            TabIndex        =   78
            Top             =   615
            Width           =   300
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "基站"
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
      Height          =   1830
      Left            =   3405
      TabIndex        =   68
      Top             =   720
      Width           =   6795
      Begin VB.OptionButton Option1 
         Caption         =   "宏蜂窝（DCS）"
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
         Index           =   3
         Left            =   4125
         TabIndex        =   16
         Top             =   1455
         Width           =   1485
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   1710
         TabIndex        =   11
         Text            =   "200"
         Top             =   1410
         Width           =   465
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1500
         TabIndex        =   4
         Text            =   "李绪华李绪华李"
         Top             =   270
         Width           =   1350
      End
      Begin VB.OptionButton Option1 
         Caption         =   "微微蜂窝"
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
         Index           =   2
         Left            =   4125
         TabIndex        =   15
         Top             =   1170
         Width           =   1080
      End
      Begin VB.OptionButton Option1 
         Caption         =   "微蜂窝"
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
         Index           =   1
         Left            =   4125
         TabIndex        =   14
         Top             =   900
         Width           =   900
      End
      Begin VB.OptionButton Option1 
         Caption         =   "宏蜂窝（GSM）"
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
         Index           =   0
         Left            =   4125
         TabIndex        =   13
         Top             =   615
         Width           =   1485
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3720
         TabIndex        =   12
         Top             =   270
         Width           =   2325
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   2220
         TabIndex        =   10
         Top             =   990
         Width           =   360
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   1590
         TabIndex        =   9
         Top             =   990
         Width           =   360
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   855
         TabIndex        =   8
         Top             =   990
         Width           =   465
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   2220
         TabIndex        =   7
         Top             =   660
         Width           =   360
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   1590
         TabIndex        =   6
         Top             =   660
         Width           =   360
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   855
         TabIndex        =   5
         Top             =   660
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "米"
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
         Index           =   55
         Left            =   2250
         TabIndex        =   129
         Top             =   1455
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "天线显示高度："
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
         Index           =   54
         Left            =   420
         TabIndex        =   128
         Top             =   1455
         Width           =   1260
      End
      Begin VB.Image Image2 
         Height          =   195
         Left            =   5115
         Picture         =   "Newbase.frx":030A
         Top             =   930
         Width           =   195
      End
      Begin VB.Image Image1 
         Height          =   225
         Left            =   5715
         Picture         =   "Newbase.frx":081C
         Top             =   645
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "秒"
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
         Index           =   41
         Left            =   2655
         TabIndex        =   115
         Top             =   1035
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "秒"
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
         Index           =   40
         Left            =   2655
         TabIndex        =   114
         Top             =   690
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "分"
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
         Index           =   39
         Left            =   2010
         TabIndex        =   113
         Top             =   1020
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "分"
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
         Index           =   38
         Left            =   2010
         TabIndex        =   112
         Top             =   690
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "度"
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
         Index           =   37
         Left            =   1365
         TabIndex        =   111
         Top             =   1035
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "度"
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
         Index           =   36
         Left            =   1365
         TabIndex        =   110
         Top             =   690
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "基站中文名："
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
         Index           =   3
         Left            =   405
         TabIndex        =   75
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "基站类型："
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
         Index           =   4
         Left            =   3180
         TabIndex        =   73
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "站址："
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
         Left            =   3180
         TabIndex        =   72
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Lat："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   450
         TabIndex        =   70
         Top             =   1050
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Lon："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   420
         TabIndex        =   69
         Top             =   720
         Width           =   435
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2340
      Top             =   1365
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":0D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":1224
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":16FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":1BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":20A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":257C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":2A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":2F28
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":33FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":38D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":3DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":4280
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":4756
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":4C2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":5102
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":55D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":5AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":5C8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":5E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Newbase.frx":6180
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "小区列表/信道分配数/邻小区数："
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
      Index           =   35
      Left            =   180
      TabIndex        =   109
      Top             =   2235
      Width           =   2700
   End
End
Attribute VB_Name = "New_Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyDatabase As Database
Dim MyRecordset As Recordset
Dim CurrentList As Integer
Dim NeedRebuild As Boolean
Dim AddressRecordset As Recordset
Dim AddressBookmark As Variant
'Dim Base_Address() As String, Cell_Length() As String
Dim strTchTemp As String
Dim strNcellTemp(15) As String
Dim CheckFlag(1) As Boolean
Dim DcsFilecopy As Boolean
Dim CurrentNewBaseIndex As Integer
Dim TextEnable(2) As Boolean

Dim szTch(2) As String
Dim szNcell(2, 15) As String
Dim EditTchFlag As Integer
Dim EditNcellFlag As Integer
Dim AddressRSFlag As Boolean

Dim IsShow As Boolean

Private Sub Command1_Click(Index As Integer)
        
    On Error Resume Next
    If EditNcellFlag > 0 Then
       SaveNcellChange
       Frame4.Visible = False
       EditNcellFlag = 0
    End If
    If EditTchFlag = 0 Then
       Frame5.Visible = True
       'Frame4.ZOrder 0
    Else
       Command1(EditTchFlag - 1).Font.Bold = False
       szTch(EditTchFlag - 1) = Trim(Text8.Text)
       If Not (IsNull(varBookmark(EditTchFlag - 1))) Then
          If szTch(EditTchFlag - 1) <> strTchTemp Then
             EditFlag(EditTchFlag - 1) = True
          End If
       End If
    End If
    Command1(Index).Font.Bold = True
    EditTchFlag = Index + 1
    Text8.Text = szTch(EditTchFlag - 1)
    strTchTemp = Text8.Text
    Text8.SetFocus

End Sub

Private Sub Command2_Click(Index As Integer)
    Dim i As Integer
    On Error Resume Next
    
    If EditTchFlag > 0 Then
       szTch(EditTchFlag - 1) = Trim(Text8.Text)
       If Not (IsNull(varBookmark(EditTchFlag - 1))) Then
          If szTch(EditTchFlag - 1) <> strTchTemp Then
             EditFlag(EditTchFlag - 1) = True
          End If
       End If
       Frame5.Visible = False
       EditTchFlag = 0
    End If
    If EditNcellFlag = 0 Then
       Frame4.Visible = True
       Frame4.ZOrder 0
    Else
       Command2(EditNcellFlag - 1).Font.Bold = False
       SaveNcellChange
    End If
    Command2(Index).Font.Bold = True
    EditNcellFlag = Index + 1
    For i = 0 To 15
        Text7(i).Text = szNcell(EditNcellFlag - 1, i)
        strNcellTemp(i) = szNcell(EditNcellFlag - 1, i)
    Next
    Text7(0).SetFocus
    
End Sub

Private Sub Form_Load()
    Dim i As Integer, j As Integer
    Dim CellName As String, NameTmp As String
    Dim itmX As ListItem
    Dim MystrTmp As String
    Dim MySubItemstmp As Integer
    Dim CellHeadData As ScanHead
    Dim IsSetitmX As Boolean
    
    'GoTo next1:
    IsShow = False
    If Dir(Gsm_Path & "\map\cell.dbf", 0) = "" Then
       hDbfFile = FreeFile
       Open Gsm_Path & "\map\cell.dbf" For Binary As #hDbfFile
       MakeCell1800File
       Close #hDbfFile
    Else
       hDbfFile = FreeFile
       Open Gsm_Path & "\map\cell.dbf" For Binary As #hDbfFile
       Get #hDbfFile, , CellHeadData
       Close #hDbfFile
       If CellHeadData.RecordLen <> (35 + 1) * 32 + 1 And CellHeadData.RecordLen <> 336 + 5 Then
           UpdateFileName = Gsm_Path & "\map\cell.dbf"
           Menu_Flag = 9999
           Data_Convert.Show 1
       End If
    End If
    For i = 0 To 2
        TextEnable(i) = True
    Next
    FileCopy Gsm_Path & "\map\cell.dbf", Gsm_Path & "\map\cell_a.dbf"
    CellFileName = "gsmcell"
    Set MyDatabase = OpenDatabase(Gsm_Path & "\map", False, False, "FoxPro 2.5;")
    Set MyRecordset = MyDatabase.OpenRecordset("SELECT * " & "FROM cell where basetype <> ""3"" or basetype=null ORDER BY cell_name ", dbOpenDynaset)    '有点奇怪？ :)
    If Dir(Gsm_Path & "\map\base_add.dbf", 0) <> "" Then
       Set AddressRecordset = MyDatabase.OpenRecordset("SELECT * " & "FROM base_add ORDER BY bs_name", dbOpenDynaset)
       AddressRSFlag = True
    Else
       AddressRSFlag = False
    End If
    'Set Data1.Recordset = MyRecordset
    'DBGrid1.Refresh
    If MyRecordset.RecordCount = 0 Then
       AddRecordBase
       Call ListView1_ItemClick(ListView1.ListItems(1))
       Exit Sub
    End If
    MyRecordset.MoveFirst
    CellName = ""
    NameTmp = ""
    CurrentNewBaseIndex = 1
    For i = 1 To MyRecordset.RecordCount
        If IsNull(MyRecordset.Fields("cell_name").Value) Then
           CellName = ""
        Else
            CellName = Trim(MyRecordset.Fields("cell_name").Value)
            If Asc(Right(CellName, 1)) >= 48 And Asc(Right(CellName, 1)) <= 57 Then
               CellName = Left(CellName, Len(CellName) - 1)
            End If
        End If
        If CellName <> NameTmp Then
           Set itmX = ListView1.ListItems.Add(, , CStr(CellName))
           IsSetitmX = True
           If IsNull(MyRecordset.Fields("non_bcch").Value) Then
              itmX.SubItems(1) = GetTchNcellNum("")
           Else
              itmX.SubItems(1) = GetTchNcellNum(Trim(MyRecordset.Fields("non_bcch").Value))
           End If
           itmX.SubItems(2) = "0"
           For j = 1 To 16
               If IsNull(MyRecordset.Fields("ncell" & Format(j)).Value) Then
                  Exit For
               Else
                  itmX.SubItems(2) = Format(j)
               End If
           Next
           NameTmp = CellName
        Else
           If IsNull(MyRecordset.Fields("non_bcch").Value) Then
              If Not IsSetitmX Then
                 If IsNull(MyRecordset.Fields("arfcn").Value) And IsNull(MyRecordset.Fields("bsic").Value) Or MyRecordset.Fields("arfcn").Value = 0 And MyRecordset.Fields("bsic").Value = 0 Then
                    MyRecordset.Delete
                    GoTo RecordMoveNext
                 End If
              Else
                 itmX.SubItems(1) = itmX.SubItems(1) & "/" & GetTchNcellNum("")
              End If
           Else
              itmX.SubItems(1) = itmX.SubItems(1) & "/" & GetTchNcellNum(Trim(MyRecordset.Fields("non_bcch").Value))
           End If
           
           For j = 1 To 16
               If IsNull(MyRecordset.Fields("ncell" & Format(j)).Value) Then
                  Exit For
               End If
           Next
           If IsSetitmX Then
              itmX.SubItems(2) = itmX.SubItems(2) & "/" & Format(j - 1)
           End If
        End If
RecordMoveNext:
        MyRecordset.MoveNext
    Next
    If ListView1.ListItems.Count > 0 Then
       Call ListView1_ItemClick(ListView1.ListItems(1))
    Else
       AddRecordBase
    End If
    ReloadIni
    CMIsCDD = True
    
End Sub

Sub ShowValue(textname, ClearFlag As Boolean, BMIsNull As Boolean)
    Dim MyIndex As Integer
    Dim j As Integer
    Dim Ncellnum As Integer, Tchnum As Integer
    
    On Error Resume Next
    Select Case textname(0).Name
       Case "Text1":
           MyIndex = 0
       Case "Text2":
           MyIndex = 1

       Case "Text3":
           MyIndex = 2
    
    End Select
    If ClearFlag Then
       textname(0) = ""
       textname(1) = ""
       textname(2) = ""
       textname(3) = ""
       textname(4) = ""
       textname(5) = ""
       textname(6) = ""
       textname(7) = ""
       textname(8) = ""
       szTch(MyIndex) = ""
       For j = 0 To 15
           szNcell(MyIndex, j) = ""
       Next
       If BMIsNull Then
          Call TextSetting(textname, False)
       End If
    Else
       Call TextSetting(textname, True)
       If MyIndex = 0 Then
          If AddressRSFlag Then
              If Not IsNull(AddressBookmark) Then
                 If IsNull(AddressRecordset.Fields("address").Value) Then
                    Text4(8).Text = ""
                 Else
                    Text4(8).Text = Trim(AddressRecordset.Fields("address").Value)
                 End If
                   'If IsNull(AddressRecordset.Fields("length").Value) Then
                   '   Text4(7).Text = ""
                   'Else
                   '   Text4(7).Text = Trim(AddressRecordset.Fields("length").Value)
                   'End If
              Else
                 Text4(8).Text = ""
                 'Text4(7).Text = ""
              End If
          Else
              Text4(8).Text = ""
              'Text4(7).Text = ""
          End If
       
       End If
       If Not IsShow Then
           If IsNull(MyRecordset.Fields("length").Value) Then
               Text4(7).Text = "200"
           Else
               Text4(7).Text = Trim(MyRecordset.Fields("length").Value)
           End If
           IsShow = True
       End If
       If IsNull(MyRecordset.Fields("bs_no").Value) Then
          textname(0) = ""
       Else
          textname(0) = Trim(MyRecordset.Fields("bs_no").Value)
       End If
       If IsNull(MyRecordset.Fields("ci").Value) Then
          textname(1) = ""
       Else
          textname(1) = Trim(MyRecordset.Fields("ci").Value)
       End If
       If IsNull(MyRecordset.Fields("lac").Value) Then
          textname(2) = ""
       Else
          textname(2) = Trim(MyRecordset.Fields("lac").Value)
       End If
       If IsNull(MyRecordset.Fields("arfcn").Value) Then
          textname(3) = ""
       Else
          textname(3) = Trim(MyRecordset.Fields("arfcn").Value)
       End If
       If IsNull(MyRecordset.Fields("bsic").Value) Then
          textname(4) = ""
       Else
          textname(4) = Trim(MyRecordset.Fields("bsic").Value)
       End If
       If IsNull(MyRecordset.Fields("bearing").Value) Then
          textname(5) = ""
       Else
          textname(5) = Trim(MyRecordset.Fields("bearing").Value)
       End If
       If IsNull(MyRecordset.Fields("downtilt").Value) Then
          textname(6) = ""
       Else
          textname(6) = Trim(MyRecordset.Fields("downtilt").Value)
       End If
       If IsNull(MyRecordset.Fields("ant_heigh").Value) Then
          textname(7) = ""
       Else
          textname(7) = Trim(MyRecordset.Fields("ant_heigh").Value)
       End If
       If IsNull(MyRecordset.Fields("max_tx_bts").Value) Then
          textname(8) = ""
       Else
          textname(8) = Trim(MyRecordset.Fields("max_tx_bts").Value)
       End If
       If IsNull(MyRecordset.Fields("non_bcch").Value) Then
          szTch(MyIndex) = ""
       Else
          szTch(MyIndex) = Trim(MyRecordset.Fields("non_bcch").Value)
       End If
       For j = 1 To 16
           If IsNull(MyRecordset.Fields("ncell" & Format(j)).Value) Then
              szNcell(MyIndex, j - 1) = ""
           Else
              szNcell(MyIndex, j - 1) = MyRecordset.Fields("ncell" & Format(j)).Value
           End If
       Next
       Ncellnum = j - 1
       Tchnum = GetTchNcellNum(szTch(MyIndex))
       ShowValueFlag = True
       If IsNull(MyRecordset.Fields("basetype").Value) Then
          Option1(0).Value = True
       Else
          Option1(Val(MyRecordset.Fields("basetype").Value)).Value = True
       End If
        Select Case textname(0).Name
           Case "Text1":
                'If InStr(ListView1.ListItems(CurrentList).SubItems(1), "/") Then
'Must modify
                
                'End If
'               ListView1.ListItems(CurrentList).SubItems(1) = GetTchNcellNum(Trim(MyRecordset.Fields("non_bcch").Value))
'               ListView1.ListItems(CurrentList).SubItems(2) = GetTchNcellNum(Trim(MyRecordset.Fields("ncellid").Value))
           Case "Text2":
'               ListView1.ListItems(CurrentList).SubItems(1) = ListView1.ListItems(CurrentList).SubItems(1) & "/" & GetTchNcellNum(Trim(MyRecordset.Fields("non_bcch").Value))
'               ListView1.ListItems(CurrentList).SubItems(2) = ListView1.ListItems(CurrentList).SubItems(2) & "/" & GetTchNcellNum(Trim(MyRecordset.Fields("ncellid").Value))
    
           Case "Text3":
'               ListView1.ListItems(CurrentList).SubItems(1) = ListView1.ListItems(CurrentList).SubItems(1) & "/" & GetTchNcellNum(Trim(MyRecordset.Fields("non_bcch").Value))
'               ListView1.ListItems(CurrentList).SubItems(2) = ListView1.ListItems(CurrentList).SubItems(2) & "/" & GetTchNcellNum(Trim(MyRecordset.Fields("ncellid").Value))
        
        End Select
       ShowValueFlag = False
    
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveIni
    CloseMyDatabase
    CMIsCDD = False
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
    Dim strFind As String
    Dim PosTmp As Single
    
    On Error Resume Next
    UpdateRecord
    
    CurrentList = ListView1.SelectedItem.Index
    If AddressRSFlag Then
        strFind = "bs_name= """ & ListView1.SelectedItem & """"
        AddressRecordset.FindFirst strFind
        If AddressRecordset.NoMatch Then
            AddressBookmark = Null
        Else
            AddressBookmark = AddressRecordset.Bookmark
        End If
    End If
    CheckFlag(0) = False
    CheckFlag(1) = False
    strFind = "instr(Cell_Name, """ & ListView1.SelectedItem & """ )>0 and (( Cell_Name= """ & ListView1.SelectedItem & """) or ( Left(Cell_Name, Len(Cell_Name) - 1) = """ & ListView1.SelectedItem & """))"
    MyRecordset.FindFirst strFind
    varBookmark(0) = MyRecordset.Bookmark
    CheckValue(0) = True
    Call ShowValue(Text1, False, False)
     
    Text4(0).Text = ListView1.SelectedItem
    If IsNull(MyRecordset.Fields("lon").Value) Then
       PosTmp = 0
    Else
       PosTmp = MyRecordset.Fields("lon").Value
    End If
    Text4(1).Text = Format(Int(PosTmp))
    Text4(2).Text = Format(Int((PosTmp - Int(PosTmp)) * 60))
    Text4(3).Text = Format(Int((((PosTmp - Int(PosTmp)) * 60) - Int((PosTmp - Int(PosTmp)) * 60)) * 60))
    If IsNull(MyRecordset.Fields("lat").Value) Then
       PosTmp = 0
    Else
       PosTmp = MyRecordset.Fields("lat").Value
    End If
    Text4(4).Text = Format(Int(PosTmp))
    Text4(5).Text = Format(Int((PosTmp - Int(PosTmp)) * 60))
    Text4(6).Text = Format(Int((((PosTmp - Int(PosTmp)) * 60) - Int((PosTmp - Int(PosTmp)) * 60)) * 60))
    
    MyRecordset.FindNext strFind
    If MyRecordset.NoMatch Then
       varBookmark(1) = Null
       Call ShowValue(Text2, True, True)
       CheckValue(1) = False
    Else
       varBookmark(1) = MyRecordset.Bookmark
       CheckValue(1) = True
       Call ShowValue(Text2, False, False)
    End If
    MyRecordset.FindNext strFind
    If MyRecordset.NoMatch Then
       varBookmark(2) = Null
       Call ShowValue(Text3, True, True)
       CheckValue(2) = False
    Else
       varBookmark(2) = MyRecordset.Bookmark
       Call ShowValue(Text3, False, False)
       CheckValue(2) = True
    End If
    Text4(0).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
    
    On Error Resume Next
    If Not ShowValueFlag Then
        EditFlag(0) = True
        EditFlag(1) = True
        EditFlag(2) = True
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 46 Then
       EditFlag(0) = True
    ElseIf KeyCode = 40 Or KeyCode = 38 Then
       Call MoveCursor(Text1, Index, KeyCode)
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii <> 13 Then
       EditFlag(0) = True
    Else
       Call MoveCursor(Text1, Index, KeyAscii)
    End If

End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Dim TextTemp As String
    
    On Error Resume Next
    TextTemp = Trim(Text1(Index).Text)
    Select Case Index
        Case 0
            If TextTemp <> "" Then
                If Asc(Right(TextTemp, 1)) >= 48 And Asc(Right(TextTemp, 1)) <= 57 Then
                    If Trim(Text2(0).Text) = "" And Text2(0).Enabled Then
                        Text2(0).Text = Left(TextTemp, Len(TextTemp) - 1) & Format(Val(Right(TextTemp, 1)) + 1)
                    End If
                    If Trim(Text3(0).Text) = "" And Text3(0).Enabled Then
                        Text3(0).Text = Left(TextTemp, Len(TextTemp) - 1) & Format(Val(Right(TextTemp, 1)) + 2)
                    End If
                Else
                    If Trim(Text2(0).Text) = "" And Text2(0).Enabled Then
                        Text2(0).Text = TextTemp & "1"
                    End If
                    If Trim(Text3(0).Text) = "" And Text3(0).Enabled Then
                        Text3(0).Text = TextTemp & "2"
                    End If
                End If
            End If
        Case 1
            If Trim(Text2(Index).Text) = "" And Text2(Index).Enabled And TextTemp <> "" Then
                Text2(Index).Text = Left(TextTemp, Len(TextTemp) - 1) & Format(Val(Right(TextTemp, 1)) + 1)
            End If
            If Trim(Text3(Index).Text) = "" And Text3(Index).Enabled And TextTemp <> "" Then
                Text3(Index).Text = Left(TextTemp, Len(TextTemp) - 1) & Format(Val(Right(TextTemp, 1)) + 2)
            End If
        Case 2, 4
            If Trim(Text2(Index).Text) = "" And Text2(Index).Enabled And TextTemp <> "" Then
                Text2(Index).Text = TextTemp
            End If
            If Trim(Text3(Index).Text) = "" And Text3(Index).Enabled And TextTemp <> "" Then
                Text3(Index).Text = TextTemp
            End If
    End Select
    
End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 46 Then
       EditFlag(1) = True
    ElseIf KeyCode = 40 Or KeyCode = 38 Then
       Call MoveCursor(Text2, Index, KeyCode)
    End If

End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii <> 13 Then
       EditFlag(1) = True
    Else
       Call MoveCursor(Text2, Index, KeyAscii)
    End If

End Sub


Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 46 Then
       EditFlag(2) = True
    ElseIf KeyCode = 40 Or KeyCode = 38 Then
       Call MoveCursor(Text3, Index, KeyCode)
    End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii <> 13 Then
       EditFlag(2) = True
    Else
       Call MoveCursor(Text3, Index, KeyAscii)
    End If

End Sub

Private Sub Text4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 46 Then
       EditFlag(0) = True
       EditFlag(1) = True
       EditFlag(2) = True
    End If
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii <> 13 Then
       EditFlag(0) = True
       EditFlag(1) = True
       EditFlag(2) = True
    End If

End Sub

Sub UpdateRecord()
    Dim i As Integer, j As Integer
    Dim MyTch(2) As String, MyNcell(2) As String
    Dim MystrTmp As String
    
    On Error Resume Next
    If EditTchFlag > 0 Then
       szTch(EditTchFlag - 1) = Trim(Text8.Text)
       If Not (IsNull(varBookmark(EditTchFlag - 1))) Then
          If szTch(EditTchFlag - 1) <> strTchTemp Then
             EditFlag(EditTchFlag - 1) = True
          End If
       End If
       Command1(EditTchFlag - 1).Font.Bold = False
       Frame5.Visible = False
       EditTchFlag = 0
    End If
    If EditNcellFlag > 0 Then
       SaveNcellChange
       Command2(EditNcellFlag - 1).Font.Bold = False
       Frame4.Visible = False
       EditNcellFlag = 0
    End If
    
    For i = 0 To 2
        If CheckValue(i) And IsNull(varBookmark(i)) Then
            MyRecordset.AddNew
            MyRecordset.Fields("cell_name").Value = "新小区1"
            MyRecordset.Update
            MyRecordset.MoveLast
            varBookmark(i) = MyRecordset.Bookmark
            EditFlag(i) = True
        ElseIf Not CheckValue(i) And (Not (IsNull(varBookmark(i)) Or IsEmpty(varBookmark(i)))) Then
            MyRecordset.Bookmark = varBookmark(i)
            MyRecordset.Delete
            varBookmark(i) = Null
            EditFlag(i) = True
            szTch(i) = ""
            For j = 0 To 15
                szNcell(i, j) = ""
            Next
        End If
    Next
    If EditFlag(0) Or EditFlag(1) Or EditFlag(2) Then
        For j = 0 To 2
            MyNcell(j) = ""
        Next
        For j = 1 To 16
            If szNcell(0, j - 1) = "" And MyNcell(0) = "" Then
                MyNcell(0) = Format(j - 1)
            End If
            If szNcell(1, j - 1) = "" And MyNcell(1) = "" Then
                MyNcell(1) = Format(j - 1)
            End If
            If szNcell(2, j - 1) = "" And MyNcell(2) = "" Then
                MyNcell(2) = Format(j - 1)
            End If
            If szNcell(0, j - 1) = "" And szNcell(1, j - 1) = "" And szNcell(2, j - 1) = "" Then
                Exit For
            End If
        Next
        If Text3(0).Enabled Then
            ListView1.ListItems(CurrentList).SubItems(1) = GetTchNcellNum(szTch(0)) & "/" & GetTchNcellNum(szTch(1)) & "/" & GetTchNcellNum(szTch(2))
            ListView1.ListItems(CurrentList).SubItems(2) = MyNcell(0) & "/" & MyNcell(1) & "/" & MyNcell(2)
        ElseIf Text2(0).Enabled Then
            ListView1.ListItems(CurrentList).SubItems(1) = GetTchNcellNum(szTch(0)) & "/" & GetTchNcellNum(szTch(1))
            ListView1.ListItems(CurrentList).SubItems(2) = MyNcell(0) & "/" & MyNcell(1)
        Else
            ListView1.ListItems(CurrentList).SubItems(1) = GetTchNcellNum(szTch(0))
            ListView1.ListItems(CurrentList).SubItems(2) = MyNcell(0)
        End If
        GoTo MyIgnore
        
        MystrTmp = ListView1.ListItems(CurrentList).SubItems(1)
        For i = 0 To 2
            If InStr(MystrTmp, "/") > 0 Then
               MyTch(i) = "/" & Left(MystrTmp, InStr(MystrTmp, "/") - 1)
               MystrTmp = Right(MystrTmp, Len(MystrTmp) - InStr(MystrTmp, "/"))
            Else
               If MystrTmp = "" Then
                  MyTch(i) = ""
               Else
                  'If i = 2 Then
                     MyTch(i) = "/" & MystrTmp
                  'Else
                  '   MyTch(i) = MyStrtmp & "/"
                     MystrTmp = ""
                  'End If
               End If
            End If
        Next

        MystrTmp = ListView1.ListItems(CurrentList).SubItems(2)
        For i = 0 To 2
            If InStr(MystrTmp, "/") > 0 Then
               MyNcell(i) = "/" & Left(MystrTmp, InStr(MystrTmp, "/") - 1)
               MystrTmp = Right(MystrTmp, Len(MystrTmp) - InStr(MystrTmp, "/"))
            Else
               If MystrTmp = "" Then
                  MyNcell(i) = ""
               Else
                  'If i = 2 Then
                     MyNcell(i) = "/" & MystrTmp
                  'Else
                     
                  '   MyNcell(i) = MyStrtmp '& "/"
                     MystrTmp = ""
                  'End If
               End If
            End If
        Next
    End If
MyIgnore:
    For i = 0 To 2
        If EditFlag(i) Then
           If Not (IsNull(varBookmark(i))) Then
                MyRecordset.Bookmark = varBookmark(i)
                MyRecordset.Edit
                MyRecordset.Fields("lon").Value = Val(Text4(1).Text) + Val(Text4(2).Text) / 60 + Val(Text4(3).Text) / 3600
                MyRecordset.Fields("lat").Value = Val(Text4(4).Text) + Val(Text4(5).Text) / 60 + Val(Text4(6).Text) / 3600
                For j = 0 To 3
                    If Option1(j).Value = True Then
                       MyRecordset.Fields("basetype").Value = Format(j)
                    End If
                Next
                
                For j = 0 To 15
                    If Trim(szNcell(i, j)) = "" Then
                       Exit For
                    End If
                Next
                
                Select Case i
                    Case 0
                         'ListView1.ListItems(CurrentList).SubItems(1) = GetTchNcellNum(szTch(i)) & MyTch(1) & MyTch(2)
                         'ListView1.ListItems(CurrentList).SubItems(2) = Format(j) & MyNcell(1) & MyNcell(2)
                        
                         ListView1.ListItems(CurrentList).Text = Trim(Text4(0).Text)
                         MyRecordset.Fields("cell_name").Value = Trim(Text4(0).Text) & "1"
                         
                         MyRecordset.Fields("bs_no").Value = Text1(0).Text
                         MyRecordset.Fields("ci").Value = Text1(1).Text
                         MyRecordset.Fields("lac").Value = Val(Text1(2).Text)
                         MyRecordset.Fields("arfcn").Value = Val(Text1(3).Text)
                         MyRecordset.Fields("bsic").Value = Val(Text1(4).Text)
                         MyRecordset.Fields("bearing").Value = Val(Text1(5).Text)
                         MyRecordset.Fields("downtilt").Value = Val(Text1(6).Text)
                         MyRecordset.Fields("ant_heigh").Value = Text1(7).Text
                         MyRecordset.Fields("max_tx_bts").Value = Text1(8).Text
                         MyRecordset.Fields("length").Value = Trim(Text4(7).Text)
                         
                         If AddressRSFlag Then
                            If IsNull(AddressBookmark) Then
                               AddressRecordset.AddNew
                            Else
                               AddressRecordset.Bookmark = AddressBookmark
                               AddressRecordset.Edit
                            End If
                             AddressRecordset.Fields("bs_name").Value = Trim(Text4(0).Text)
                             AddressRecordset.Fields("lon").Value = MyRecordset.Fields("lon").Value
                             AddressRecordset.Fields("lat").Value = MyRecordset.Fields("lat").Value
    '                         AddressRecordset.Fields("length").Value = Trim(Text4(7).Text)
                             AddressRecordset.Fields("address").Value = Trim(Text4(8).Text)
                             AddressRecordset.Update
                             If IsNull(AddressBookmark) Then
                                AddressRecordset.MoveLast
                                AddressBookmark = AddressRecordset.Bookmark
                             End If
                         End If
                    
                    Case 1
                         'If InStr(ListView1.ListItems(CurrentList).SubItems(2), "/") = 0 Then
                         '   ListView1.ListItems(CurrentList).SubItems(2) = ListView1.ListItems(CurrentList).SubItems(2) & "/" & GetTchNcellNum(szTch(i))
                         '   ListView1.ListItems(CurrentList).SubItems(1) = ListView1.ListItems(CurrentList).SubItems(1) & "/" & Format(j)
                         'Else
                         '   ListView1.ListItems(CurrentList).SubItems(2) = Right(MyNcell(0), Len(MyNcell(0)) - 1) & "/" & GetTchNcellNum(szTch(i)) & MyNcell(2)
                         '   ListView1.ListItems(CurrentList).SubItems(1) = Right(MyTch(0), Len(MyTch(0)) - 1) & "/" & Format(j) & MyTch(2)
                         'End If
                        
                         MyRecordset.Fields("cell_name").Value = Trim(Text4(0).Text) & "2"
                         
                         MyRecordset.Fields("bs_no").Value = Text2(0).Text
                         MyRecordset.Fields("ci").Value = Text2(1).Text
                         MyRecordset.Fields("lac").Value = Val(Text2(2).Text)
                         MyRecordset.Fields("arfcn").Value = Val(Text2(3).Text)
                         MyRecordset.Fields("bsic").Value = Val(Text2(4).Text)
                         MyRecordset.Fields("bearing").Value = Val(Text2(5).Text)
                         MyRecordset.Fields("downtilt").Value = Val(Text2(6).Text)
                         MyRecordset.Fields("ant_heigh").Value = Text2(7).Text
                         MyRecordset.Fields("max_tx_bts").Value = Text2(8).Text
                         MyRecordset.Fields("length").Value = Trim(Text4(7).Text)
                    Case 2
                         'MyStrtmp = ListView1.ListItems(CurrentList).SubItems(2)
                         'MyStrtmp = Right(MyStrtmp, Len(MyStrtmp) - InStr(MyStrtmp, "/"))
                         'If InStr(MyStrtmp, "/") = 0 Then
                         '   ListView1.ListItems(CurrentList).SubItems(2) = ListView1.ListItems(CurrentList).SubItems(2) & "/" & GetTchNcellNum(szTch(i))
                         '   ListView1.ListItems(CurrentList).SubItems(1) = ListView1.ListItems(CurrentList).SubItems(1) & "/" & Format(j)
                         'Else
                         '   ListView1.ListItems(CurrentList).SubItems(2) = Right(MyNcell(0), Len(MyNcell(0)) - 1) & MyNcell(1) & "/" & GetTchNcellNum(szTch(i))
                         '   ListView1.ListItems(CurrentList).SubItems(1) = Right(MyTch(0), Len(MyTch(0)) - 1) & MyTch(1) & "/" & Format(j)
                         'End If
                         MyRecordset.Fields("cell_name").Value = Trim(Text4(0).Text) & "3"
                                                  
                         MyRecordset.Fields("bs_no").Value = Text3(0).Text
                         MyRecordset.Fields("ci").Value = Text3(1).Text
                         MyRecordset.Fields("lac").Value = Val(Text3(2).Text)
                         MyRecordset.Fields("arfcn").Value = Val(Text3(3).Text)
                         MyRecordset.Fields("bsic").Value = Val(Text3(4).Text)
                         MyRecordset.Fields("bearing").Value = Val(Text3(5).Text)
                         MyRecordset.Fields("downtilt").Value = Val(Text3(6).Text)
                         MyRecordset.Fields("ant_heigh").Value = Text3(7).Text
                         MyRecordset.Fields("max_tx_bts").Value = Text3(8).Text
                         MyRecordset.Fields("length").Value = Trim(Text4(7).Text)
                End Select
                
                
                MyRecordset.Fields("non_bcch").Value = szTch(i)
                For j = 1 To 16
                    MyRecordset.Fields("ncell" & Format(j)).Value = szNcell(i, j - 1)
                Next
                MyRecordset.Update
                If Not NeedRebuild Then
                   NeedRebuild = True
                End If
           End If
           EditFlag(i) = False
        End If
    Next
    
End Sub

Sub AddRecordBase()
    Dim itmX As ListItem
    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next
    Call TextSetting(Text1, True)
    Call TextSetting(Text2, True)
    Call TextSetting(Text3, True)
    For i = 0 To 8
        Text4(i).Text = ""
    Next
    'Text4(7).Text = "200"
    For i = 0 To 2
        CheckValue(i) = True
        
        MyRecordset.AddNew
        MyRecordset.Fields("bs_no").Value = ""
        MyRecordset.Fields("ci").Value = ""
        MyRecordset.Fields("lac").Value = 0
        MyRecordset.Fields("arfcn").Value = 0
        MyRecordset.Fields("bsic").Value = 0
        MyRecordset.Fields("bearing").Value = 0
        MyRecordset.Fields("downtilt").Value = 0
        MyRecordset.Fields("ant_heigh").Value = ""
        MyRecordset.Fields("max_tx_bts").Value = ""
        MyRecordset.Fields("non_bcch").Value = ""
        MyRecordset.Fields("length").Value = "200"
        
                For j = 1 To 16
                    MyRecordset.Fields("ncell" & Format(j)).Value = ""
                Next
        If CellFileName = "gsmcell" Then
            MyRecordset.Fields("basetype").Value = "0"
        Else
            MyRecordset.Fields("basetype").Value = "3"
        End If
        
        Select Case i
            Case 0
                MyRecordset.Fields("cell_name").Value = "新小区" & Format(CurrentNewBaseIndex) & "1"
                MyRecordset.Update
                MyRecordset.MoveLast
                varBookmark(0) = MyRecordset.Bookmark
                Call ShowValue(Text1, True, False)
            Case 1
                MyRecordset.Fields("cell_name").Value = "新小区" & Format(CurrentNewBaseIndex) & "2"
                MyRecordset.Update
                MyRecordset.MoveLast
                varBookmark(1) = MyRecordset.Bookmark
                Call ShowValue(Text2, True, False)
            Case 2
                MyRecordset.Fields("cell_name").Value = "新小区" & Format(CurrentNewBaseIndex) & "3"
                MyRecordset.Update
                MyRecordset.MoveLast
                varBookmark(2) = MyRecordset.Bookmark
                Call ShowValue(Text3, True, False)
        End Select
    Next
    Set itmX = ListView1.ListItems.Add(, , CStr("新小区" & Format(CurrentNewBaseIndex)))
    CurrentNewBaseIndex = CurrentNewBaseIndex + 1
    Set ListView1.SelectedItem = ListView1.ListItems(ListView1.ListItems.Count)
    CurrentList = ListView1.SelectedItem.Index
End Sub

Sub DeleteRecord()
    Dim i As Integer
    On Error Resume Next
    frmDelete.Show 1
    If Not CheckValue(0) And Not CheckValue(1) And Not CheckValue(2) Then
        For i = 0 To 2
            If Not IsNull(varBookmark(i)) Then
               MyRecordset.Bookmark = varBookmark(i)
               MyRecordset.Delete
               varBookmark(i) = Null
            End If
        Next
        ListView1.ListItems.Remove CurrentList
        If ListView1.ListItems.Count > 0 Then
           Call ListView1_ItemClick(ListView1.ListItems(ListView1.SelectedItem.Index))
        End If
    End If
End Sub

Private Sub Text7_LostFocus(Index As Integer)
    On Error Resume Next
    If EditNcellFlag > 0 Then
        If Not (UCase(Screen.ActiveControl.Name) = "TEXT7" Or UCase(Screen.ActiveControl.Name) = "COMMAND2") Then
           Command2(EditNcellFlag - 1).Font.Bold = False
           SaveNcellChange
           Frame4.Visible = False
           EditNcellFlag = 0
        End If
    End If
End Sub

Private Sub Text8_LostFocus()
    On Error Resume Next
    If EditTchFlag > 0 Then
        If Not (UCase(Screen.ActiveControl.Name) = "TEXT8" Or UCase(Screen.ActiveControl.Name) = "COMMAND1") Then
            Command1(EditTchFlag - 1).Font.Bold = False
            szTch(EditTchFlag - 1) = Trim(Text8.Text)
            If Not (IsNull(varBookmark(EditTchFlag - 1))) Then
               If szTch(EditTchFlag - 1) <> strTchTemp Then
                  EditFlag(EditTchFlag - 1) = True
               End If
            End If
            Frame5.Visible = False
            EditTchFlag = 0
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Dim LastSortType As Byte
    
    On Error Resume Next
    UpdateRecord
    Select Case Button.Index
        Case 2
            AddRecordBase
        Case 3
            AddRecordBaseCell
        Case 3 + 1
            DeleteRecord
        Case 5 + 1
            LastSortType = SortType
            frmSort.Show 1
            'If SortType <> LastSortType Then
               ReloadRecord (SortType)
            'End If
        Case 6 + 1
            StatRecord
        Case 8 + 1
            ChangeGsmDcs
        Case 10 + 1
            DownloadCell
        Case 11 + 1
            RebuildCell
        Case 12 + 1
            UpgradeCell
        Case 15
            ExcelExport
        Case 16
            ExcelImport
        Case 18
            ExitModify
    End Select
End Sub

Sub ReloadRecord(MySortType As Byte)
    Dim strGsmDcs As String
    Dim i As Integer, j As Integer
    Dim CellName As String, NameTmp As String
    Dim itmX As ListItem
    
    On Error Resume Next
    If CellFileName = "gsmcell" Then
       strGsmDcs = "basetype <> ""3"" Or BASETYPE = Null"
    Else
       strGsmDcs = "basetype = ""3"""
    End If
    Select Case MySortType
        Case 0
            'Set MyRecordset = MyDatabase.OpenRecordset("SELECT * " & "FROM " & CellFileName & " ORDER BY cell_name", dbOpenDynaset)
            Set MyRecordset = MyDatabase.OpenRecordset("SELECT * " & "FROM cell where " & strGsmDcs & " ORDER BY cell_name ", dbOpenDynaset)
        Case 1
            Set MyRecordset = MyDatabase.OpenRecordset("SELECT * " & "FROM cell where " & strGsmDcs & " ORDER BY lac,cell_name", dbOpenDynaset)
        Case 2
            Set MyRecordset = MyDatabase.OpenRecordset("SELECT * " & "FROM cell where " & strGsmDcs & " ORDER BY ci,cell_name", dbOpenDynaset)
        Case 3
            Set MyRecordset = MyDatabase.OpenRecordset("SELECT * " & "FROM cell where " & strGsmDcs & " ORDER BY microcell,cell_name", dbOpenDynaset)
        Case 4
            Set MyRecordset = MyDatabase.OpenRecordset("SELECT * " & "FROM cell where " & strGsmDcs & " ORDER BY time,cell_name", dbOpenDynaset)
    End Select
    
    'Set Data1.Recordset = MyRecordset
    'DBGrid1.Refresh
    If MyRecordset.RecordCount = 0 Then
       ListView1.ListItems.Clear
       AddRecordBase
       Call ListView1_ItemClick(ListView1.ListItems(1))
       Exit Sub
    End If
    MyRecordset.MoveFirst
    CellName = ""
    NameTmp = ""
    ListView1.ListItems.Clear
    For i = 1 To MyRecordset.RecordCount
        If IsNull(MyRecordset.Fields("cell_name").Value) Then
           CellName = ""
        Else
           CellName = Trim(MyRecordset.Fields("cell_name").Value)
           If Asc(Right(CellName, 1)) >= 48 And Asc(Right(CellName, 1)) <= 57 Then
              CellName = Left(CellName, Len(CellName) - 1)
           End If
        End If
        If CellName <> NameTmp Then
           Set itmX = ListView1.ListItems.Add(, , CStr(CellName))
           If Not IsNull(MyRecordset.Fields("non_bcch").Value) Then
              itmX.SubItems(1) = GetTchNcellNum(Trim(MyRecordset.Fields("non_bcch").Value))
           Else
              itmX.SubItems(1) = 0
           End If
           For j = 1 To 16
               If IsNull(MyRecordset.Fields("ncell" & Format(j)).Value) Then
                  Exit For
               End If
           Next
           itmX.SubItems(2) = Format(j - 1)
           NameTmp = CellName
        Else
           If IsNull(MyRecordset.Fields("non_bcch").Value) Then
              itmX.SubItems(1) = itmX.SubItems(1) & "/" & "0"
           Else
              itmX.SubItems(1) = itmX.SubItems(1) & "/" & GetTchNcellNum(Trim(MyRecordset.Fields("non_bcch").Value))
           End If
           
           For j = 1 To 16
               If IsNull(MyRecordset.Fields("ncell" & Format(j)).Value) Then
                  Exit For
               End If
           Next
           itmX.SubItems(2) = itmX.SubItems(2) & "/" & Format(j - 1)
           'If Trim(MyRecordset.Fields("non_bcch").Value) <> "" Then
           '   itmX.SubItems(1) = itmX.SubItems(1) & "/" & InStr(MyRecordset.Fields("non_bcch").Value, ",") + 1
           'End If
           'If Trim(MyRecordset.Fields("ncellid").Value) <> "" Then
           '   itmX.SubItems(2) = itmX.SubItems(2) & "/" & InStr(MyRecordset.Fields("ncellid").Value, ",") + 1
           'End If
        End If
        MyRecordset.MoveNext
    Next
    Call ListView1_ItemClick(ListView1.ListItems(1))
       
End Sub

Sub ExitModify()
    On Error Resume Next
    CloseMyDatabase
    If NeedRebuild Then
       If (MsgBox("保存对小区库存所做的修改吗？", 33, "提示")) = 1 Then
       
       Else
          FileCopy Gsm_Path & "\map\cell_a.dbf", Gsm_Path & "\map\cell.dbf"
          Kill Gsm_Path & "\map\cell_a.dbf"
       End If
    End If
    Unload Me
End Sub

Sub RebuildCell()
    Dim i As Integer, j As Integer, k As Integer, HH As Integer
    Dim CellRows As Long
    Dim MyCellColor As Long
    Dim AntennaL As Single
    Dim old_name As String, o_name As String
    Dim BASETYPE As String
    Dim BaseColor As Long
    Dim bcch(3), BSIC(3), ci(3), MyDir(3), other(4) As String
    Dim Lac As String, bs_name As String, bs_no As String
    Dim LeeFinds As Integer
    Dim row As Long, LacRow As Long
    Dim MyLon As Variant, MyLat As Variant
    
    On Error Resume Next
    If CellFileName = "cell" Then
       NeedRebuild = False
    End If
    SaveIni
    
    UpdateRecord
    CloseMyDatabase
    
        mapinfo.do "Register Table " + Chr(34) + Gsm_Path & "\map\cell.dbf" + Chr(34) + "Type ""DBF"" Into " + Chr(34) + Gsm_Path & "\map\cell.tab" + Chr(34)
        mapinfo.do "open table " + Chr(34) + Gsm_Path & "\map\cell.tab" + Chr(34)
        mapinfo.do "pack table cell Graphic Data  Data  Interactive  "
        CellRows = mapinfo.eval("tableinfo(cell,8)")
        mapinfo.do "Create Map For cell CoordSys Earth Projection 1, 0"
        mapinfo.do "fetch first from cell"
        For i = 1 To CellRows
            'If Val(mapinfo.eval("cell.length")) = 0 Then
            If Val(Text4(7).Text) = 0 Then
               AntennaL = 0.002
               Text4(7).Text = "200"
            Else
               AntennaL = Val(Text4(7).Text) / 100000
            End If
'            mapinfo.do " x1 = cell.Lon + " & Format(AntennaL) & " * Sin(cell.bearing * 0.01745329252)" '  DEG_2_RAD)"
'            mapinfo.do " y1 = cell.Lat + " & Format(AntennaL) & " * Cos(cell.bearing * 0.01745329252)"  ' DEG_2_RAD)"
'            MyCellColor = MyCellRndColor(Val(mapinfo.eval("cell.arfcn")))
            If Val(mapinfo.eval("cell.basetype")) = 0 Then
               mapinfo.do " x1 = cell.Lon + " & Format(AntennaL) & " * Sin(cell.bearing * 0.01745329252)" '  DEG_2_RAD)"
               mapinfo.do " y1 = cell.Lat + " & Format(AntennaL) & " * Cos(cell.bearing * 0.01745329252)"  ' DEG_2_RAD)"
               If mapinfo.eval("cell.arfcn") > 124 Then
                  MyCellColor = MyCellRndColor(Val(mapinfo.eval("cell.arfcn")) Mod 124)
               Else
                  MyCellColor = MyCellRndColor(Val(mapinfo.eval("cell.arfcn")))
               End If
               mapinfo.do "Set Style Pen MakePen(1,60," & Format(MyCellColor) & ")"
               mapinfo.do "update cell  set Obj= CreateLine(x1,y1,cell.lon, cell.Lat)  where rowid=" & i
               'If AntennaL = 0.002 Then
               If mapinfo.eval("cell.length") <> Text4(7).Text Then
                  mapinfo.do "update cell  set LENGTH= " & Text4(7).Text & "  where rowid=" & i
               End If
            ElseIf Val(mapinfo.eval("cell.basetype")) = 3 Then
               mapinfo.do " x1 = cell.Lon + " & Format(AntennaL / 1.5) & " * Sin(cell.bearing * 0.01745329252)" '  DEG_2_RAD)"
               mapinfo.do " y1 = cell.Lat + " & Format(AntennaL / 1.5) & " * Cos(cell.bearing * 0.01745329252)" ' DEG_2_RAD)"
               If mapinfo.eval("cell.arfcn") > 124 Then
                  MyCellColor = MyCellRndColor(Val(mapinfo.eval("cell.arfcn")) Mod 124)
               Else
                  MyCellColor = MyCellRndColor(Val(mapinfo.eval("cell.arfcn")))
               End If
               mapinfo.do "Set Style Pen MakePen(1,60,0)"
               mapinfo.do "update cell  set Obj= CreateLine(x1,y1,cell.lon, cell.Lat)  where rowid=" & i
               If mapinfo.eval("cell.length") <> Text4(7).Text Then
                  mapinfo.do "update cell  set LENGTH= " & Text4(7).Text & "  where rowid=" & i
               End If
            Else
               mapinfo.do "set style symbol MakeFontSymbol(59,16711680,12,""MapInfo Weather"",256,-Cell.bearing)"
               mapinfo.do "update cell set Obj= CreatePoint(cell.Lon,cell.Lat ) where rowid=" & i
            End If
            mapinfo.do "fetch next from cell"
        Next
        mapinfo.do "commit table cell"

    If Dir(Gsm_Path & "\map\base.tab", 0) = "" Then
        hDbfFile = FreeFile
        Open Gsm_Path & "\map\base.dbf" For Binary As #hDbfFile
        MakeBase1800File
        Close #hDbfFile
        mapinfo.do "Register Table " + Chr(34) + Gsm_Path & "\map\base.dbf" + Chr(34) + "Type ""DBF"" Into " + Chr(34) + Gsm_Path & "\map\base.tab" + Chr(34)
    End If
    mapinfo.do "open table " & Chr(34) + Gsm_Path & "\map\base.tab" + Chr(34)
    
    k = 1
    j = 1
    old_name = " "
    mapinfo.do "fetch first from cell"
    old_name = Trim(mapinfo.eval("cell.cell_name"))
    Do While mapinfo.eval("EOT(cell)") <> "T"
       If old_name <> "" Then Exit Do
       mapinfo.do "fetch next from cell"
       old_name = Trim(mapinfo.eval("cell.cell_name"))
    Loop
    If old_name = "" Then
       mapinfo.do "fetch first from base"
       mapinfo.do "delete from base"
       mapinfo.do "commit table base"
       mapinfo.do "pack table base Graphic Data Data Interactive  "
       GoTo no_cell
    End If
    Call getname(old_name)
    mapinfo.do "fetch First from base"
    
    mapinfo.do "delete from base"
    mapinfo.do "commit table Base"
    mapinfo.do "pack table base Graphic Data Data Interactive  "
    For i = 1 To 3
        ci(i) = " "
        MyDir(i) = "0"
        BSIC(i) = "0"
        bcch(i) = "0"
    Next i
   
   o_name = old_name
   While mapinfo.eval("EOT(cell)") <> "T"
         If old_name = o_name Then
cc:         bcch(k) = mapinfo.eval("cell.arfcn")
            ci(k) = mapinfo.eval("cell.ci")
            MyDir(k) = mapinfo.eval("cell.bearing")
            BSIC(k) = mapinfo.eval("cell.bsic")
            Lac = mapinfo.eval("cell.lac")
            bs_name = mapinfo.eval("cell.cell_name")
            bs_no = mapinfo.eval("cell.bs_no")
            BASETYPE = mapinfo.eval("cell.basetype")
            If BASETYPE = "" Then
               BASETYPE = "0"
            End If
            HH = 0
            For i = 1 To 4
                HH = i + 15
                other(i) = mapinfo.eval("cell.col" & HH)
            Next i
            mapinfo.do "x1 = cell.lon"
            mapinfo.do "y1 = cell.lat"
            k = k + 1
         Else
            LeeFinds = InStr(bs_no, Chr(0))
            If LeeFinds > 0 Then
               bs_no = Trim(Left(bs_no, LeeFinds - 1))
            End If
            Msg = "insert into  base  (col1,col2,col3,col4,col5,col6,col7,col8,col9,col10,col11,col12,col13,col14,col15,col16,col17,col18,col19) values ( "
            Msg = Msg + Chr(34) + old_name + Chr(34) + ","
            Msg = Msg + Chr(34) + bs_no + Chr(34) + ","
            Msg = Msg + bcch(1) + ","
            Msg = Msg + bcch(2) + ","
            Msg = Msg + bcch(3) + ","
            Msg = Msg + Chr(34) + ci(1) + Chr(34) + ","
            Msg = Msg + Chr(34) + ci(2) + Chr(34) + ","
            Msg = Msg + Chr(34) + ci(3) + Chr(34) + ","
            Msg = Msg + BSIC(1) + ","
            Msg = Msg + BSIC(2) + ","
            Msg = Msg + BSIC(3) + ","
            Msg = Msg + MyDir(1) + ","
            Msg = Msg + MyDir(2) + ","
            Msg = Msg + MyDir(3) + ","
            Msg = Msg + Chr(34) + Lac + Chr(34) + ","
            'msg = msg + Chr(34) + other(1) + Chr(34) + ","
            Msg = Msg + Chr(34) + " " + Chr(34) + ","
            Msg = Msg + Chr(34) + BASETYPE + Chr(34) + ","
            Msg = Msg + Chr(34) + " " + Chr(34) + ","
            Msg = Msg + Chr(34) + " " + Chr(34) + ")"
            mapinfo.do Msg
            Msg = "UPDATE  base  set lon = x1,lat = y1 where  rowid = " & j
            mapinfo.do Msg
            For i = 1 To 3
                ci(i) = " "
                MyDir(i) = "0"
                BSIC(i) = "0"
                bcch(i) = "0"
            Next i
            k = 1
            j = j + 1
            GoTo cc
         End If
         old_name = o_name
         Do While mapinfo.eval("EOT(cell)") <> "T"
            mapinfo.do "fetch next from cell"
            o_name = Trim(mapinfo.eval("cell.cell_name"))
            If o_name <> "" Then Exit Do
         Loop
         If o_name = "" Then GoTo exit_do
'         o_name = Mid(o_name, 1, 4)
         Call getname(o_name)

    Wend
exit_do:
           Msg = "insert into  base  (col1,col2,col3,col4,col5,col6,col7,col8,col9,col10,col11,col12,col13,col14,col15,col16,col17,col18,col19) values ( "
           Msg = Msg + Chr(34) + old_name + Chr(34) + ","
           Msg = Msg + Chr(34) + bs_no + Chr(34) + ","
           Msg = Msg + bcch(1) + ","
           Msg = Msg + bcch(2) + ","
           Msg = Msg + bcch(3) + ","
           Msg = Msg + Chr(34) + ci(1) + Chr(34) + ","
           Msg = Msg + Chr(34) + ci(2) + Chr(34) + ","
           Msg = Msg + Chr(34) + ci(3) + Chr(34) + ","
           Msg = Msg + BSIC(1) + ","
           Msg = Msg + BSIC(2) + ","
           Msg = Msg + BSIC(3) + ","
           Msg = Msg + MyDir(1) + ","
           Msg = Msg + MyDir(2) + ","
           Msg = Msg + MyDir(3) + ","
           Msg = Msg + Chr(34) + Lac + Chr(34) + ","
           Msg = Msg + Chr(34) + "" + Chr(34) + ","
           Msg = Msg + Chr(34) + BASETYPE + Chr(34) + ","
           Msg = Msg + Chr(34) + "" + Chr(34) + ","
           Msg = Msg + Chr(34) + "" + Chr(34) + ")"
           mapinfo.do Msg
           Msg = "UPDATE base set lon= x1,lat= y1 WHERE ROWID=" & j
           mapinfo.do Msg
           
    mapinfo.do "commit table Base"
    mapinfo.do "DROP MAP Base"
    mapinfo.do "Create Map For Base CoordSys Earth Projection 1, 0"

    i = 0
    'If Is_New = False Then
    '   Gsm_FileName = Gsm_Path + "\base_add.dbf"
    '   Gsm_File2 = Gsm_Path + "\map\base_add.dbf"
    '   Kill Gsm_File2
    '   FileCopy Gsm_FileName, Gsm_File2
    '   Gsm_FileName = Gsm_Path + "\map\base_add.tab"
    '   Kill Gsm_FileName
    '   mapinfo.Do "Register Table " + Chr(34) + Gsm_File2 + Chr(34) + "Type ""DBF"" Into " + Chr(34) + Gsm_FileName + Chr(34)
    'End If
    'mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\map\base_add.tab" + Chr(34)
    row = Val(mapinfo.eval("TABLEINFO(Base, 8)"))
    mapinfo.do "SELECT lac FROM base where lac>0 group by lac order by lac desc into mytemp"
    LacRow = Val(mapinfo.eval("TABLEINFO(mytemp, 8)"))
    mapinfo.do "fetch first from Base"
    i = 1
    While i <= row
          mapinfo.do "fetch first from mytemp"
          If mapinfo.eval("base.lac") = 0 Then
             BaseColor = 0
          Else
             For k = 1 To LacRow
                 If mapinfo.eval("base.lac") = mapinfo.eval("mytemp.lac") Then
                    Exit For
                 End If
                 mapinfo.do "fetch next from mytemp"
             Next
             BaseColor = MyLacColor(k - 1)
          End If
'          msg = "base.bsic_1"
'          j = Val(mapinfo.eval(msg))
'          j = j * 12345678 + j * 876543
          'msg = "Set Style Symbol MakeFontSymbol(168," & j & ",12,""Symbol"",0,0)"
          'mapinfo.do "Set Style Symbol MakeFontSymbol(39," & j & ",12,""Wingdings 2"",256,0)"
         
          mapinfo.do "set style symbol MakeFontSymbol(39," & Format(BaseColor) & ",8,""MapInfo Cartographic"",0,0)"
          mapinfo.do "update Base  set Obj= CreatePoint(Lon,Lat ) where rowid=" & i
          old_name = mapinfo.eval("base.bs_name")
          Call getname(old_name)
          mapinfo.do "x1 = base.lon"
          mapinfo.do "y1 = base.lat"
             
          'msg = "insert into base_add (bs_name,address,lon,lat) values (" + Chr(34) + old_name + Chr(34) + "," + Chr(34) + Base_Address(i) + Chr(34) + ",x1,y1)"
          'mapinfo.Do msg
          'mapinfo.Do "fetch next from Base_add"
          mapinfo.do "fetch next from Base"
          i = i + 1
    Wend
        
'*********************************************************************************
no_cell:
    mapinfo.do "commit table base"
    mapinfo.do "close table base"
    
    mapinfo.do "close table cell"
    
    CellFileName = "gsmcell"
    Set MyDatabase = OpenDatabase(Gsm_Path & "\map", False, False, "FoxPro 2.5;")
    ReloadRecord (0)
    
End Sub

Sub StatRecord()
    Dim MyStatRecordset As Recordset
    Dim i As Integer
    Dim MyrecordTmp As Integer
    
    On Error Resume Next
    StatString = Trim(Text5(0).Text) & Label1(7).Caption & Chr(13) & Chr(10)
    StatString = StatString & "第 " & Trim(Text5(1).Text) & " 期工程" & Chr(13) & Chr(10)
    StatString = StatString & "第 " & Trim(Text5(2).Text) & " 次设置" & Chr(13) & Chr(10)
    StatString = StatString & Chr(13) & Chr(10)
    StatString = StatString & "基站数量：" & Format(ListView1.ListItems.Count) & Chr(13) & Chr(10)
    StatString = StatString & "小区数量" & Chr(13) & Chr(10)
    Set MyStatRecordset = MyDatabase.OpenRecordset("SELECT * FROM cell where basetype =null ", dbOpenDynaset)
    MyrecordTmp = MyStatRecordset.RecordCount
    Set MyStatRecordset = MyDatabase.OpenRecordset("SELECT basetype, Count([basetype]) AS counter " & "FROM cell GROUP BY basetype order by basetype ", dbOpenDynaset)
    For i = 1 To MyStatRecordset.RecordCount
        Select Case MyStatRecordset.Fields("basetype").Value
            Case "0"
                MyrecordTmp = MyrecordTmp + MyStatRecordset.Fields("counter").Value
                StatString = StatString & "    宏蜂窝：" & Format(MyrecordTmp) & Chr(13) & Chr(10)
            Case "1"
                'If i = 2 Then
                    StatString = StatString & "    微蜂窝：" & Format(MyStatRecordset.Fields("counter").Value) & Chr(13) & Chr(10)
                'End If
            Case "2"
                'If i = 3 Then
                    StatString = StatString & "    微微蜂窝：" & Format(MyStatRecordset.Fields("counter").Value) & Chr(13) & Chr(10)
                'End If
        End Select
        MyStatRecordset.MoveNext
    Next
    StatString = StatString & Chr(13) & Chr(10)
    StatString = StatString & "BCCH载频数：" & Chr(13) & Chr(10)
    Set MyStatRecordset = MyDatabase.OpenRecordset("SELECT  arfcn, Count([arfcn]) AS counter  " & "FROM cell GROUP BY arfcn ORDER by Count([arfcn]) DESC,arfcn DESC", dbOpenDynaset)
    For i = 1 To MyStatRecordset.RecordCount
        StatString = StatString & "    " & Format(MyStatRecordset.Fields("arfcn").Value) & ": " & Format(MyStatRecordset.Fields("counter").Value) & Chr(13) & Chr(10)
        MyStatRecordset.MoveNext
    Next
    StatString = StatString & Chr(13) & Chr(10)
    StatString = StatString & Label1(7).Caption & "占用带宽： "
    Set MyStatRecordset = MyDatabase.OpenRecordset("SELECT " & "Min(arfcn) AS ArfcnMin, " & "Max(arfcn) AS ArfcnMax " & "FROM cell", dbOpenDynaset)
    StatString = StatString & Format((MyStatRecordset.Fields("ArfcnMax").Value - MyStatRecordset.Fields("ArfcnMin").Value) * 0.2, "0.00")
    StatString = StatString & " MHz" & Chr(13) & Chr(10)
    MyStatRecordset.Close
    
    frmStat.Show 1
    
End Sub

Sub ChangeGsmDcs()
    Dim myFilename As String, strLine As String
    
    On Error Resume Next
    UpdateRecord
    SaveIni
    If CellFileName = "gsmcell" Then
       CellFileName = "dcscell"
       Label1(7).Caption = "DCS"
    Else
       CellFileName = "gsmcell"
       Label1(7).Caption = "GSM"
    End If
    SortType = 0
    ReloadRecord (0)
        
    ReloadIni

End Sub

Sub DownloadCell()
    Dim MyRecord As Record
    Dim mypath As String, buff As String
    Dim finds As Integer, i As Integer
    Dim hFreefile As Integer
    Dim myFilename As String
    Dim MyPutStr As String
    
    On Error Resume Next
    UpdateRecord
    CloseMyDatabase
    
    MDIMain.StatusBar.Panels(2).Text = " 自动下载"
    Menu_Flag = 2302
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    hFreefile = FreeFile
    Open Gsm_FileName For Binary As #hFreefile
    Get #hFreefile, 1, MyRecord  ' Read third record.
    Close #hFreefile
    If Val(MyRecord.exchange) = 0 Or Val(MyRecord.exchange) = 1 Or Val(MyRecord.exchange) = 4 Or Val(MyRecord.exchange) = 5 Then
       For i = 1 To 50
           convert_filename(i) = ""
       Next
open_again:
       MDIMain.FileDialog.DialogTitle = "数据转换文件选择"
       If Val(MyRecord.exchange) = 0 Then                    '爱立信
          MDIMain.FileDialog.Filter = "All Files|*.*"
          MDIMain.FileDialog.DefaultExt = ""
          MDIMain.FileDialog.Flags = &H80000
       Else
          MDIMain.FileDialog.Filter = "*.txt Files|*.TXT|*.xls Files|*.XLS|All Files|*.*"
          MDIMain.FileDialog.DefaultExt = "*.TXT"
          If Val(MyRecord.exchange) = 4 Then                'ITALTEL
             'MDIMain.FileDialog.Flags = &H200 Or &H80000
             MDIMain.FileDialog.Flags = &H80000
          Else
             MDIMain.FileDialog.Flags = &H80000             'MOTOROLA
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
       If Dir(convert_filename(1)) = "" Then
          GoTo err_exit
       End If
       Screen.MousePointer = 11
       Data_Convert.Show 1
       Screen.MousePointer = 0
       GoTo complete
err_exit:
       i = MsgBox("无法打开文件 " + convert_filename(1), 48, "打开文件")
       GoTo open_again
    Else
       MsgBox "该交换机类型的数据转换暂未挂接!", 64, "提示"
    End If
complete:
    MDIMain.StatusBar.Panels(2).Text = " "
'****************************************************
    
    hFreefile = FreeFile
    myFilename = Gsm_Path & "\User\DownLoad.log"
    If Dir(myFilename, 0) <> "" Then
        Kill myFilename
    End If
    Open myFilename For Binary As #hFreefile
    MyPutStr = "交换机文件中未找到匹配小区名：" & Chr(13) & Chr(10)
    Put #hFreefile, , MyPutStr
    Put #hFreefile, , FileNoMatch
    MyPutStr = Chr(13) & Chr(10)
    Put #hFreefile, , MyPutStr
    MyPutStr = "小区库中未更新小区名：" & Chr(13) & Chr(10)
    Put #hFreefile, , MyPutStr
    Set MyDatabase = OpenDatabase(Gsm_Path & "\map", False, False, "FoxPro 2.5;")
    Set MyRecordset = MyDatabase.OpenRecordset("SELECT cell_name,bs_no FROM cell where time <> """ & year(Now) & Format(month(Now), "00") & Format(day(Now), "00") & """ or time=null ORDER BY cell_name", dbOpenDynaset)
    If MyRecordset.RecordCount = 0 Then
        MyPutStr = "        小区库已全部更新"
        Put #hFreefile, , MyPutStr
    Else
        For i = 1 To MyRecordset.RecordCount
            MyPutStr = "        " & Format(MyRecordset.Fields("bs_no").Value) & Chr(13) & Chr(10)
            Put #hFreefile, , MyPutStr
            MyRecordset.MoveNext
        Next
    End If
    Close #hFreefile

'***************************************************
    SortType = 0
    ReloadRecord (0)
        
End Sub

Sub SaveIni()
    Dim myFilename As String, strLine As String
    Dim hFreefile As Integer

    On Error Resume Next
    myFilename = Gsm_Path & "\" & CellFileName & ".ini"
    If Dir(myFilename, 0) <> "" Then
       Kill myFilename
    End If
    hFreefile = FreeFile
    Open myFilename For Output As #hFreefile
    strLine = Trim(Text5(0).Text)
    Print #hFreefile, strLine
    strLine = Trim(Text5(1).Text)
    Print #hFreefile, strLine
    strLine = Trim(Text5(2).Text)
    Print #hFreefile, strLine
    Close #hFreefile

End Sub

Sub ReloadIni()
    Dim myFilename As String, strLine As String
    Dim hFreefile As Integer
    
    On Error Resume Next
    myFilename = Gsm_Path & "\gsmcell.ini"
    hFreefile = FreeFile
    If Dir(myFilename, 0) <> "" Then
       Open myFilename For Input As #hFreefile
       Line Input #hFreefile, strLine
       Text5(0).Text = Trim(strLine)
       Line Input #hFreefile, strLine
       Text5(1).Text = Trim(strLine)
       Line Input #hFreefile, strLine
       Text5(2).Text = Trim(strLine)
       Close #hFreefile
    Else
       Text5(0).Text = ""
       Text5(1).Text = ""
       Text5(2).Text = ""
    End If
    
End Sub

Sub MoveCursor(textname, Index, KeyAscii)
    
    On Error Resume Next
    If KeyAscii = 13 Or KeyAscii = 40 Then
       Select Case textname(0).Name
          Case "Text5":
              If Index = 2 Then
                 Text4(0).SetFocus
              Else
                 textname(Index + 1).SetFocus
              End If
          Case "Text4":
              If Index = 8 Then
                 Text1(0).SetFocus
              Else
                 textname(Index + 1).SetFocus
              End If
          Case "Text1":
              If Index = 8 Then
                 If Text2(0).Enabled Then
                    Text2(0).SetFocus
                 Else
                    Text5(0).SetFocus
                 End If
              Else
                 textname(Index + 1).SetFocus
              End If
          Case "Text2":
              If Index = 8 Then
                 If Text3(0).Enabled Then
                    Text3(0).SetFocus
                 Else
                    Text5(0).SetFocus
                 End If
              Else
                 textname(Index + 1).SetFocus
              End If
          Case "Text3":
              If Index = 8 Then
                 Text5(0).SetFocus
              Else
                 textname(Index + 1).SetFocus
              End If
       End Select
    ElseIf KeyAscii = 38 Then
       Select Case textname(0).Name
          Case "Text5":
              If Index = 0 Then
                 If Text3(8).Enabled Then
                    Text3(8).SetFocus
                 ElseIf Text2(8).Enabled Then
                    Text2(8).SetFocus
                 Else
                    Text1(8).SetFocus
                 End If
              Else
                 textname(Index - 1).SetFocus
              End If
          Case "Text4":
              If Index = 0 Then
                 Text5(2).SetFocus
              Else
                 textname(Index - 1).SetFocus
              End If
          Case "Text1":
              If Index = 0 Then
                 Text4(8).SetFocus
              Else
                 textname(Index - 1).SetFocus
              End If
          Case "Text2":
              If Index = 0 Then
                 Text1(8).SetFocus
              Else
                 textname(Index - 1).SetFocus
              End If
          Case "Text3":
              If Index = 0 Then
                 Text2(8).SetFocus
              Else
                 textname(Index - 1).SetFocus
              End If
       End Select
    End If

End Sub

Function GetTchNcellNum(MyStr As String) As String
    Dim MySubItemstmp As Integer
    Dim MystrTemp As String
    
    On Error Resume Next
    MystrTemp = MyStr
    If MystrTemp <> "" Then
       MySubItemstmp = 1
       Do While InStr(MystrTemp, ",") > 0
          MySubItemstmp = MySubItemstmp + 1
          MystrTemp = Right(MystrTemp, Len(MystrTemp) - InStr(MystrTemp, ","))
       Loop
       GetTchNcellNum = Format(MySubItemstmp)
    Else
       GetTchNcellNum = "0"
    End If

End Function

Sub UpgradeCell()
    
    On Error Resume Next
    frmUpgradeCell.Show 1
    
End Sub

Sub AddRecordBaseCell()

    On Error Resume Next
       If Not CheckValue(1) And Not CheckFlag(0) Then
            Call TextSetting(Text2, True)
            Text2(0).Text = Left(Trim(Text1(0).Text), Len(Trim(Text1(0).Text)) - 1) & Format(Val(Right(Trim(Text1(0).Text), 1)) + 1)
            Text2(1).Text = Left(Trim(Text1(1).Text), Len(Trim(Text1(1).Text)) - 1) & Format(Val(Right(Trim(Text1(1).Text), 1)) + 1)
            Text2(2).Text = Trim(Text1(2).Text)
            Text2(4).Text = Trim(Text1(4).Text)
            CheckValue(1) = True
            CheckFlag(0) = True
       
       ElseIf Not CheckValue(2) And Not CheckFlag(1) Then
            Call TextSetting(Text3, True)
            Text3(0).Text = Left(Trim(Text1(0).Text), Len(Trim(Text1(0).Text)) - 1) & Format(Val(Right(Trim(Text1(0).Text), 1)) + 2)
            Text3(1).Text = Left(Trim(Text1(1).Text), Len(Trim(Text1(1).Text)) - 1) & Format(Val(Right(Trim(Text1(1).Text), 1)) + 2)
            Text3(2).Text = Trim(Text1(2).Text)
            Text3(4).Text = Trim(Text1(4).Text)
            CheckValue(2) = True
            CheckFlag(1) = True
       End If

End Sub

Sub ExcelExport()
    Dim My_Excel As Object
    Dim ExportFile As String
    
    On Error Resume Next
    'ExportFile = MyOpenFile(".xls", CellFileName & ".xls", True)
    ExportFile = MyOpenFile(".xls", "cell.xls", True)
    If ExportFile = "" Then
       Exit Sub
    End If
    MyRecordset.Close
    Screen.MousePointer = 11
    Set My_Excel = CreateObject("Excel.Application")
    My_Excel.Visible = False
    Gsm_FileName = Gsm_Path + "\map\cell.dbf"
    Gsm_File2 = "c:\celltmp.dbf"
    FileCopy Gsm_FileName, Gsm_File2
    My_Excel.Workbooks.Open filename:=Gsm_File2
    'My_Excel.ActiveWorkbook.SaveAs filename:=ExportFile, FileFormat:=-4143
    My_Excel.ActiveWorkbook.Saveas filename:=ExportFile, FileFormat:=39
    My_Excel.ActiveWindow.Close
    My_Excel.quit
    Set My_Excel = Nothing
    Kill Gsm_File2

    ReloadRecord (SortType)
    Screen.MousePointer = 0
    
End Sub

Sub ExcelImport()
    Dim XLSRows As Long
    Dim hFreefile As Integer
    Dim CellData As NewCell1800
    Dim MyCellColor As String
    Dim i As Long
    Dim j As Integer
    Dim ImportFile As String
    
    On Error Resume Next
    MyRecordset.Close
    ImportFile = MyOpenFile(".xls", "", False)
    If ImportFile = "" Then
       Exit Sub
    End If
    Gsm_FileName = ImportFile
    frmDownLoad.Show 1
    ReloadRecord (SortType)

    Exit Sub
    mapinfo.do "Register Table " + Chr(34) + ImportFile + Chr(34) + " TYPE XLS Into " + Chr(34) + Gsm_Path + "\CellTemp.tab" + Chr(34)
    'mapinfo.do "Register Table " + Chr(34) + ImportFile + Chr(34) + " TYPE XLS Into " + Chr(34) + Left(ImportFile, Len(ImportFile) - 4) & ".tab" + Chr(34)
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\CellTemp.tab" + Chr(34)
    If Err Then
       MsgBox "无法打开文件 " & ImportFile & "或文件格式错误，" & Chr(10) & "请确定该文件是Excel 5.0/95格式并且只有一个工作表再做导入。", 64, "提示"
       Screen.MousePointer = 0
       Exit Sub
    End If
    XLSRows = mapinfo.eval("tableinfo(CellTemp,8)")
    If Dir(Gsm_Path & "\map\cell.dbf", 0) <> "" Then
       Kill Gsm_Path & "\map\cell.dbf"
    End If
    hFreefile = funcCreateCell(Gsm_Path & "\map\cell.dbf")
    Seek #hFreefile, 21 * 32 + 2
    
    CellData.b = " "
    CellData.Name = " "
    CellData.bs_no = " "
    CellData.ci = " "
    CellData.ARFCN = " "
    CellData.BSIC = " "
    CellData.bearing = " "
    CellData.Lac = " "
    CellData.NONBCCH = " "
    CellData.downtilt = " "
    CellData.MAX_BTS = " "
    CellData.ANT_HEIGH = " "
    CellData.MAX_MS = " "
    CellData.ANT_GAIN = " "
    CellData.ANT_TYPE = " "
    CellData.BASETYPE = " "
    For i = 1 To 16
        CellData.NCELL(i) = " "
    Next
    CellData.lon = " "
    CellData.lat = " "
    CellData.time = " "
    mapinfo.do "fetch first from CellTemp"
    mapinfo.do "fetch next from CellTemp"
    For i = 1 To XLSRows - 1
        CellData.Name = Trim(mapinfo.eval("celltemp.col1"))
        If InStr(CellData.Name, Chr(0)) > 0 Then
           CellData.Name = Trim(Left(CellData.Name, InStr(CellData.Name, Chr(0)) - 1))
        End If
        CellData.bs_no = Trim(mapinfo.eval("celltemp.col2"))
        CellData.ci = Trim(mapinfo.eval("celltemp.col3"))
        CellData.ARFCN = Trim(mapinfo.eval("celltemp.col4"))
        CellData.BSIC = Trim(mapinfo.eval("celltemp.col5"))
        CellData.bearing = Trim(mapinfo.eval("celltemp.col6"))
        CellData.Lac = Trim(mapinfo.eval("celltemp.col7"))
        CellData.NONBCCH = Trim(mapinfo.eval("celltemp.col8"))
        CellData.downtilt = Trim(mapinfo.eval("celltemp.col9"))
        CellData.MAX_BTS = Trim(mapinfo.eval("celltemp.col10"))
        CellData.ANT_HEIGH = Trim(mapinfo.eval("celltemp.col11"))
        CellData.MAX_MS = Trim(mapinfo.eval("celltemp.col12"))
        CellData.ANT_GAIN = Trim(mapinfo.eval("celltemp.col13"))
        CellData.ANT_TYPE = Trim(mapinfo.eval("celltemp.col14"))
        CellData.time = Trim(mapinfo.eval("celltemp.col15"))
        CellData.lon = Trim(mapinfo.eval("celltemp.col16"))
        CellData.lat = Trim(mapinfo.eval("celltemp.col17"))
        CellData.BASETYPE = Trim(mapinfo.eval("celltemp.col18"))
        For j = 1 To 16
            CellData.NCELL(j) = Trim(mapinfo.eval("celltemp.col" & Format(j + 18)))
        Next
        Put #hFreefile, , CellData
        mapinfo.do "fetch next from CellTemp"
    Next
    Seek #hFreefile, 5
    XLSRows = XLSRows - 1
    Put #hFreefile, , XLSRows
    Close #hFreefile
    
        mapinfo.do "Register Table " + Chr(34) + Gsm_Path & "\map\cell.dbf" + Chr(34) + "Type ""DBF"" Into " + Chr(34) + Gsm_Path & "\map\cell.tab" + Chr(34)
        mapinfo.do "open table " + Chr(34) + Gsm_Path & "\map\cell.tab" + Chr(34)
        mapinfo.do "Create Map For cell CoordSys Earth Projection 1, 0 "
        mapinfo.do "fetch first from cell"
        For i = 1 To XLSRows
            mapinfo.do " x1 = cell.Lon + 0.002 * Sin(cell.bearing * 0.01745329252)" '  DEG_2_RAD)"
            mapinfo.do " y1 = cell.Lat + 0.002 * Cos(cell.bearing * 0.01745329252)"  ' DEG_2_RAD)"
            MyCellColor = Format(MyCellRndColor(Val(mapinfo.eval("cell.arfcn"))))
            If Val(mapinfo.eval("cell.basetype")) = 0 Then
               mapinfo.do "Set Style Pen MakePen(1,60," & MyCellColor & ")"
               mapinfo.do "update cell set Obj= CreateLine(x1,y1,cell.lon, cell.Lat)  where rowid=" & i
            Else
               mapinfo.do "set style symbol MakeFontSymbol(59,16711680,12,""MapInfo Weather"",256,-cell.bearing)"
               mapinfo.do "update cell set Obj= CreatePoint( cell.Lon, cell.Lat ) where rowid=" & i
            End If
            mapinfo.do "fetch next from cell"
        Next
        mapinfo.do "commit table cell"
        mapinfo.do "close table cell"
    
    Kill Gsm_Path + "\CellTemp.*"


End Sub

Sub CloseMyDatabase()
    On Error Resume Next
    If Not IsEmpty(MyRecordset) Then
       MyRecordset.Close
    End If
    If AddressRSFlag Then
       AddressRecordset.Close
       AddressRSFlag = False
    End If
    If Not IsEmpty(MyDatabase) Then
       MyDatabase.Close
    End If
    
End Sub

Function MyOpenFile(MyDefaultExt As String, MyDefaultFile As String, IsSaveas As Boolean)
    Dim hFreefile As Integer
    Dim myFilename As String
    Dim i As Integer
    
    On Error Resume Next
    myFilename = MyDefaultFile
    MDIMain.FileDialog.CancelError = True
    MDIMain.FileDialog.filename = myFilename
    MDIMain.FileDialog.Filter = MyDefaultExt
    MDIMain.FileDialog.DefaultExt = MyDefaultExt
    MDIMain.FileDialog.Flags = &H80000
    MDIMain.FileDialog.InitDir = Gsm_Path & "\User"
    Err = 0
open_again:
    If IsSaveas Then
       MDIMain.FileDialog.ShowSave
    Else
       MDIMain.FileDialog.ShowOpen
    End If
    If Err Then
       'GoTo ExitWindow
       MDIMain.FileDialog.filename = ""
       MDIMain.FileDialog.CancelError = False
       MyOpenFile = ""
       Exit Function
    End If
    If MDIMain.FileDialog.filename <> "" Then
       myFilename = MDIMain.FileDialog.filename
       If IsSaveas Then
            If Dir(myFilename, 0) <> "" Then
               i = MsgBox(myFilename & " 已存在，是否将它覆盖？", 49, "保存文件")
               If i = 2 Then
                  MDIMain.FileDialog.filename = ""
                  GoTo open_again
               Else
                  Kill myFilename
               End If
            End If
       End If
    Else
       GoTo ExitWindow
    End If
    
ExitWindow:
    MDIMain.FileDialog.filename = ""
    MDIMain.FileDialog.CancelError = False
    MyOpenFile = myFilename

End Function

Sub getname(MyName)
    Dim mychar As String
    Dim mycode As Integer, finds As Integer
    
    On Error Resume Next
    finds = InStr(MyName, Chr(0))
    If finds > 0 Then
       MyName = Left(MyName, finds - 1)
    End If
    MyName = Trim(MyName)
    If Len(MyName) > 0 Then
       mychar = Right(MyName, 1)
       mycode = Asc(mychar)
       'If mycode >= 65 And mycode <= 90 Or mycode >= 97 And mycode <= 122 Or mycode >= 48 And mycode <= 57 Then
       If mycode >= 48 And mycode <= 57 Then
          MyName = Left(MyName, Len(MyName) - 1)
          MyName = Trim(MyName)
       End If
    End If
End Sub

Private Sub SaveNcellChange()
    Dim i As Integer
    
    On Error Resume Next
    If IsNull(varBookmark(EditNcellFlag - 1)) Then
       Exit Sub
    End If
    For i = 0 To 15
        szNcell(EditNcellFlag - 1, i) = Trim(Text7(i).Text)
        If szNcell(EditNcellFlag - 1, i) <> strNcellTemp(i) Then
           EditFlag(EditNcellFlag - 1) = True
        End If
    Next
   
End Sub

Private Sub TextSetting(MyTextName, MySetting As Boolean)
    Dim i As Integer, j As Integer
    
    On Error Resume Next
    Select Case MyTextName(0).Name
        Case "Text1"
            j = 0
        Case "Text2"
            j = 1
        Case "Text3"
            j = 2
    End Select
    If TextEnable(j) = MySetting Then
        Exit Sub
    End If
    For i = 0 To 8
        MyTextName(i).Enabled = MySetting
    Next
    If j = 1 Then
        For i = 23 To 25
            Label1(i).Enabled = MySetting
        Next
        For i = 20 To 21
            Label1(i).Enabled = MySetting
        Next
        For i = 8 To 11
            Label1(i).Enabled = MySetting
        Next
        For i = 44 To 45
            Label1(i).Enabled = MySetting
        Next
        For i = 50 To 51
            Label1(i).Enabled = MySetting
        Next
        Command1(1).Enabled = MySetting
        Command2(1).Enabled = MySetting
    ElseIf j = 2 Then
        For i = 26 To 34
            Label1(i).Enabled = MySetting
        Next
        For i = 46 To 47
            Label1(i).Enabled = MySetting
        Next
        For i = 52 To 53
            Label1(i).Enabled = MySetting
        Next
        Command1(2).Enabled = MySetting
        Command2(2).Enabled = MySetting
    End If
    TextEnable(j) = MySetting
    
End Sub

