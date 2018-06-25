VERSION 5.00
Begin VB.Form Iland_Base 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "基站选择"
   ClientHeight    =   2790
   ClientLeft      =   3345
   ClientTop       =   2115
   ClientWidth     =   3420
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Base_sel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   3420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "小区选择"
      Height          =   1545
      Left            =   270
      TabIndex        =   4
      Top             =   600
      Width           =   2865
      Begin VB.OptionButton Cell_3 
         Caption         =   "小区-3"
         Height          =   240
         Left            =   300
         TabIndex        =   13
         Top             =   1140
         Width           =   915
      End
      Begin VB.OptionButton Cell_2 
         Caption         =   "小区-2"
         Height          =   240
         Left            =   300
         TabIndex        =   12
         Top             =   765
         Width           =   915
      End
      Begin VB.OptionButton Cell_1 
         Caption         =   "小区-1"
         Height          =   240
         Left            =   300
         TabIndex        =   11
         Top             =   390
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox Arfcn_1 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Enabled         =   0   'False
         Height          =   270
         Left            =   2070
         TabIndex        =   7
         Text            =   "  "
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Arfcn_2 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Enabled         =   0   'False
         Height          =   270
         Left            =   2070
         TabIndex        =   6
         Text            =   " "
         Top             =   735
         Width           =   495
      End
      Begin VB.TextBox Arfcn_3 
         BackColor       =   &H00E0E0E0&
         DataField       =   " "
         DataSource      =   " "
         Enabled         =   0   'False
         Height          =   270
         Left            =   2070
         TabIndex        =   5
         Top             =   1110
         Width           =   495
      End
      Begin VB.Label Cell1Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1485
         TabIndex        =   10
         Top             =   390
         Width           =   525
         WordWrap        =   -1  'True
      End
      Begin VB.Label Cell2Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1485
         TabIndex        =   9
         Top             =   765
         Width           =   540
      End
      Begin VB.Label Cell3Label 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ARFCN:"
         Height          =   180
         Left            =   1485
         TabIndex        =   8
         Top             =   1140
         Width           =   540
      End
   End
   Begin VB.ComboBox Combo1 
      DataField       =   " "
      DataSource      =   " "
      Height          =   300
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   165
      Width           =   1695
   End
   Begin VB.CommandButton SBSCancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   1785
      TabIndex        =   3
      Top             =   2355
      Width           =   1080
   End
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   555
      TabIndex        =   2
      Top             =   2355
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "基站选择："
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "Iland_Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cell_1_Click()
    On Error Resume Next
    If Cell_1.Value = True Then Menu_Flag = 1
End Sub

Private Sub cell_2_Click()
    On Error Resume Next
    If Cell_2.Value = True Then Menu_Flag = 2
End Sub

Private Sub cell_3_Click()
    On Error Resume Next
    If Cell_3.Value = True Then Menu_Flag = 3
End Sub

Private Sub Combo1_Click()
   On Error Resume Next
        If Combo1.Text <> "" Then
           i = 0
           row = Val(mapinfo.eval("tableinfo(base,8)"))
           mapinfo.do "fetch First from base"
           msg = Mid$(mapinfo.eval("base.bs_NAME"), 1, 5)
           While i < row And msg <> Combo1.Text
              mapinfo.do "fetch next from base"
              msg = Mid$(mapinfo.eval("base.bs_NAME"), 1, 5)
              i = i + 1
           Wend
           Arfcn_1.Enabled = 1
           Arfcn_2.Enabled = 1
           Arfcn_3.Enabled = 1
           
           SBSOK.Enabled = 1
           SBSCancel.Enabled = 1
           
           Arfcn_1.Text = mapinfo.eval("base.BCCH_1")
           Arfcn_2.Text = mapinfo.eval("base.BCCH_2")
           Arfcn_3.Text = mapinfo.eval("base.BCCH_3")
     End If
End Sub

Private Sub Form_Load()
  On Error Resume Next
  i = 0
  row = Val(mapinfo.eval("tableinfo(base,8)"))
  mapinfo.do "fetch First from base"
  Combo1.Text = mapinfo.eval("base.bs_NAME")
  While i < row
       Combo1.AddItem mapinfo.eval("base.bs_NAME")
       mapinfo.do "fetch next from base"
       i = i + 1
  Wend

  Arfcn_1.Text = mapinfo.eval("base.BCCH_1")
  Arfcn_2.Text = mapinfo.eval("base.BCCH_2")
  Arfcn_3.Text = mapinfo.eval("base.BCCH_3")
  mapinfo.do "fetch First from base"
  Combo1.Text = mapinfo.eval("base.bs_NAME")
  Full_Flag = 1
End Sub

Private Sub SBSCancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub SBSOK_Click()
 Dim X, Y, x1, y1 As Single
  
 Dim ci(4), bs_name, bs_overlay  As String
 Dim i, j, row, BSIC(3), ARFCN(3), Lac As Integer
 Dim Ta, Rxqual As Integer

  
 On Error Resume Next
 Screen.MousePointer = 11
 If Combo1.Text <> "" Then
    i = 0
    row = Val(mapinfo.eval("tableinfo(base,8)"))
    mapinfo.do "fetch First from base"
    msg = Mid$(mapinfo.eval("base.bs_NAME"), 1, 5)
    While i < row And msg <> Combo1.Text
             mapinfo.do "fetch next from base"
             msg = Mid$(mapinfo.eval("base.bs_NAME"), 1, 5)
             i = i + 1
    Wend
  
        bs_name = Combo1.Text
        ARFCN(1) = Val(mapinfo.eval("base.BCCH_1"))
        ARFCN(2) = Val(mapinfo.eval("base.BCCH_2"))
        ARFCN(3) = Val(mapinfo.eval("base.BCCH_3"))

        BSIC(1) = Val(mapinfo.eval("base.BSIC_1"))
        BSIC(2) = Val(mapinfo.eval("base.BSIC_2"))
        BSIC(3) = Val(mapinfo.eval("base.BSIC_3"))

        ci(1) = CStr(Val(mapinfo.eval("base.ci_1")))
        ci(2) = CStr(Val(mapinfo.eval("base.ci_2")))
        ci(3) = CStr(Val(mapinfo.eval("base.ci_3")))

  Unload Me


          rmsg1 = ci(Menu_Flag)
          rmsg2 = bs_name

          MapForm.Show
          mapHWnd = Val(mapinfo.eval("WindowInfo(" & mapid & ",12)"))
          If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
             MapForm.WindowState = 0
          End If
          MapForm.Move 0, 10, 12000, 4450

          Iland_Dis.Show
          Iland_Dis.Move 10, 4450, Iland_Dis.Width, Iland_Dis.Height

 End If
  Screen.MousePointer = 0
  Unload Me
End Sub


