VERSION 5.00
Begin VB.Form Face 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  "
   ClientHeight    =   4800
   ClientLeft      =   1995
   ClientTop       =   2250
   ClientWidth     =   7455
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF8080&
   Icon            =   "Face.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Face.frx":030A
   ScaleHeight     =   4800
   ScaleWidth      =   7455
   Begin VB.CommandButton SCANOK 
      Caption         =   "扫频分析"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   6180
      TabIndex        =   1
      Top             =   4275
      Width           =   1095
   End
   Begin VB.CommandButton cmdStartStop 
      BackColor       =   &H00000000&
      Caption         =   "通话分析"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   4995
      TabIndex        =   0
      Top             =   4275
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 4.0.3（2000版）"
      DataSource      =   "&H00000000&"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   3870
      TabIndex        =   9
      Top             =   1695
      Width           =   2475
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   6
      Left            =   3885
      TabIndex        =   8
      Top             =   4485
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WanHe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   5
      Left            =   3255
      TabIndex        =   7
      Top             =   4485
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright(C) 2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   4
      Left            =   1620
      TabIndex        =   6
      Top             =   4485
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G-2000-20318"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   3
      Left            =   2385
      TabIndex        =   5
      Top             =   3975
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "珠海"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Index           =   2
      Left            =   2385
      TabIndex        =   4
      Top             =   3690
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "序列号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Index           =   1
      Left            =   1620
      TabIndex        =   3
      Top             =   4005
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "使用者："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   180
      Index           =   0
      Left            =   1620
      TabIndex        =   2
      Top             =   3690
      Width           =   720
   End
End
Attribute VB_Name = "Face"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    On Error Resume Next
    Menu_Flag = 0
    Label2(2).Caption = Trim(USERNAME)

End Sub

Private Sub SCANOK_Click()
    On Error Resume Next
    sys = 1
    If Face_show > 0 Then
        MDIMain.SUB_121.Enabled = 0
        MDIMain.Tems98_Convert.Enabled = False
        MDIMain.ANTSurveyor.Enabled = 0
        MDIMain.SUB_Obtel.Enabled = 0
        MDIMain.Sub_31.Enabled = 0
        MDIMain.ViewNcell.Enabled = 0
        MDIMain.MnuC_A.Enabled = 0
      '  MDIMain.MnuSql.Enabled = 0
        MDIMain.MnuLabel.Enabled = False
        MDIMain.MnuLabelMark.Enabled = 0
        MDIMain.RadioLink.Enabled = 0
        MDIMain.Hopping.Enabled = 0
        MDIMain.SUB_431.Enabled = 0
      '  MDIMain.NetworkBlind.Enabled = 0
    '    MDIMain.NetworkDisturb.Enabled = 0
   '     MDIMain.View_Cope.Enabled = 0
        
        'SUB_32.Enabled = 0
        MDIMain.SUB_33.Enabled = 0
        MDIMain.SUB_41.Enabled = 0
    '    SUB_42.Enabled = 0
    '    SUB_43.Enabled = 0
        MDIMain.SUB_441.Enabled = 0
        MDIMain.SUB_442.Enabled = 0
  '      MDIMain.SUB_443.Enabled = 0
        MDIMain.Mnu_Replay.Enabled = 0
 '       MDIMain.SUB_45.Enabled = 0
        MDIMain.SUB_123.Enabled = 1
        MDIMain.ScanPilot.Enabled = 1
    '    SCAN_2.Enabled = 1
    '    SCAN_3.Enabled = 1
        MDIMain.My_ScanPlay.Enabled = True
        MDIMain.Arfcn_Changing.Enabled = True
        MDIMain.SCAN_4.Enabled = 1
        MDIMain.SCAN_5.Enabled = 1
        MDIMain.SCAN_6.Enabled = 1
        MDIMain.SCAN_7.Enabled = 1
        MDIMain.SCAN_8.Enabled = 1
        MDIMain.TRAN_C_I.Enabled = 1
        MDIMain.My_Over.Enabled = 1
        MDIMain.Mnudistributing.Enabled = 1
    '    OPen_Str_Data.Enabled = 0
    '    Static_Pad.Enabled = 0
    '    report.Enabled = 0
    '    STREET_AN.Enabled = 0
        MDIMain.StatusBar.Panels(4).Text = "扫频分析"
    End If
    Face.Hide
    Unload Face
End Sub
Private Sub cmdStartStop_Click()
    On Error Resume Next
    sys = 0
    If Face_show > 0 Then
        MDIMain.SUB_123.Enabled = 0
        MDIMain.ScanPilot.Enabled = 0
    '    SCAN_2.Enabled = 0
    '    SCAN_3.Enabled = 0
        MDIMain.My_ScanPlay.Enabled = False
        MDIMain.Arfcn_Changing.Enabled = False
        MDIMain.SCAN_4.Enabled = 0
        MDIMain.SCAN_5.Enabled = 0
        MDIMain.SCAN_6.Enabled = 0
        MDIMain.SCAN_7.Enabled = 0
        MDIMain.SCAN_8.Enabled = 0
        MDIMain.My_Over.Enabled = 0
        MDIMain.Mnudistributing.Enabled = 0
        MDIMain.TRAN_C_I.Enabled = 0
    
        MDIMain.SUB_121.Enabled = 1
        MDIMain.Tems98_Convert.Enabled = True
        MDIMain.ANTSurveyor.Enabled = 1
        MDIMain.SUB_Obtel.Enabled = 1
        MDIMain.Sub_31.Enabled = 1
        MDIMain.ViewNcell.Enabled = 1
        MDIMain.MnuC_A.Enabled = 1
        'MDIMain.MnuSql.Enabled = 1
        MDIMain.MnuLabel.Enabled = True
        MDIMain.MnuLabelMark.Enabled = 1
        MDIMain.RadioLink.Enabled = 1
        MDIMain.Hopping.Enabled = 1
        MDIMain.SUB_431.Enabled = 1
        'MDIMain.NetworkBlind.Enabled = 1
     '   MDIMain.NetworkDisturb.Enabled = 1
        'MDIMain.View_Cope.Enabled = 1
        'SUB_32.Enabled = 1
        MDIMain.SUB_33.Enabled = 1
        MDIMain.SUB_41.Enabled = 1
    '    SUB_42.Enabled = 1
        MDIMain.SUB_441.Enabled = 1
        MDIMain.SUB_442.Enabled = 1
        'MDIMain.SUB_443.Enabled = 1
        'SUB_43.Enabled = 1
        'MDIMain.SUB_45.Enabled = 1
        MDIMain.Mnu_Replay.Enabled = 1
        MDIMain.StatusBar.Panels(4).Text = "通话分析"
    End If
    Face.Hide
    Unload Face
End Sub

