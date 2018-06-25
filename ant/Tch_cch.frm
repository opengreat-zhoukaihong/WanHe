VERSION 5.00
Begin VB.Form TCH_CCH_SEL 
   BackColor       =   &H00C0C0C0&
   Caption         =   "TCH/CCH 基站选择"
   ClientHeight    =   2925
   ClientLeft      =   2370
   ClientTop       =   1725
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tch_cch.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2925
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   240
      TabIndex        =   8
      Top             =   2055
      Width           =   1965
      Begin VB.CheckBox CCH 
         Caption         =   "CCH"
         Height          =   240
         Left            =   1230
         TabIndex        =   10
         Top             =   300
         Width           =   600
      End
      Begin VB.CheckBox TCH 
         Caption         =   "TCH"
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   300
         Width           =   585
      End
   End
   Begin VB.CommandButton DEL 
      Caption         =   "<<"
      Height          =   280
      Left            =   1995
      TabIndex        =   7
      Top             =   1260
      Width           =   465
   End
   Begin VB.CommandButton ADD 
      Caption         =   ">>"
      Height          =   280
      Left            =   1995
      TabIndex        =   6
      Top             =   870
      Width           =   465
   End
   Begin VB.ListBox List2 
      Height          =   1500
      Left            =   2640
      TabIndex        =   3
      Top             =   465
      Width           =   1560
   End
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   240
      TabIndex        =   2
      Top             =   465
      Width           =   1560
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   3135
      TabIndex        =   1
      Top             =   2550
      Width           =   1080
   End
   Begin VB.CommandButton OK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   3135
      TabIndex        =   0
      Top             =   2160
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "已选基站："
      Height          =   180
      Left            =   2655
      TabIndex        =   5
      Top             =   165
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "可选基站："
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   165
      Width           =   900
   End
End
Attribute VB_Name = "TCH_CCH_SEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ADD_Click()
    On Error Resume Next
    If List1.ListIndex >= 0 And List2.ListCount < 6 Then
      List2.AddItem List1.List(List1.ListIndex)
      List1.RemoveItem List1.ListIndex
    End If
End Sub

Private Sub Cancel_Click()
    On Error Resume Next
    mapinfo.do "close table cch_sts"
    mapinfo.do "close table tch_sts"
   
    Unload Me
End Sub

Private Sub DEL_Click()
    On Error Resume Next
    If List2.ListIndex >= 0 And List2.ListCount >= 0 Then
       List1.AddItem List2.List(List2.ListIndex)
       List2.RemoveItem List2.ListIndex
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim row As Integer
    mapinfo.do "close table tch_sts"
    On Error Resume Next
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\tch_sts.tab" + Chr(34)
    mapinfo.do "fetch first from tch_sts"
    row = Val(mapinfo.eval("tableinfo(""tch_sts"", 8)"))
    While row >= 1
       List1.AddItem mapinfo.eval("tch_sts.col1")
       mapinfo.do "fetch next from tch_sts"
       row = row - 1
    Wend
    mapinfo.do "close table ""tch_sts"""
    TCH.Value = 1
End Sub

Private Sub OK_Click()
    If MapForm.WindowState = 1 Or MapForm.WindowState = 2 Then
       MapForm.WindowState = 0
    End If
    MapForm.Move 0, 10, 12000, 4000
    Dim brow_id As Integer
    Gsm_FileName = Gsm_Path + "\sts"
    ChDir Gsm_FileName
    On Error Resume Next

  mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
  mapinfo.do "set paper units ""pt"""
  If List2.ListCount = 0 Then
    If TCH.Value = 1 And CCH.Value = 0 Then
       mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\tch_sts" + Chr(34)
       mapinfo.do "browse *  from tch_sts  "
       mapinfo.do "set window Frontwindow() Position(0,250) Width 600 Height 160 "
    Else
       If CCH.Value = 1 And TCH.Value = 0 Then
          mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\cch_sts" + Chr(34)
          mapinfo.do "browse *  from cch_sts  "
          mapinfo.do "set window Frontwindow() Position(0,250) Width 600 Height 160 "
        End If
    End If

    If TCH.Value = 1 And CCH.Value = 1 Then
       mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\tch_sts" + Chr(34)
       mapinfo.do "browse *  from tch_sts  "
       mapinfo.do "set window Frontwindow() Position(0,250) Width 600 Height 80 "

       mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"

       mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\cch_sts" + Chr(34)
       mapinfo.do "browse *  from cch_sts  "
       mapinfo.do "set window Frontwindow() Position(0,335) Width 600 Height 80 "
    End If
Else

    If TCH.Value = 1 And CCH.Value = 0 Then
       mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\tch_sts" + Chr(34)
       msg = "select * from tch_sts  where col1=" + Chr(34) + List2.List(0) + Chr(34) + " OR col1=" + Chr(34) + List2.List(1) + Chr(34) + " OR col1=" + Chr(34) + List2.List(2) + Chr(34) + " OR col1=" + Chr(34) + List2.List(3) + Chr(34) + " OR col1=" + Chr(34) + List2.List(4) + Chr(34) + " OR col1=" + Chr(34) + List2.List(5) + Chr(34) + " OR col1=" + Chr(34) + List2.List(6) + Chr(34) + " Into TCH"
       mapinfo.do msg
       mapinfo.do "browse *  from tch  "
       mapinfo.do "set window Frontwindow() Position(0,250) Width 600 Height 160 "
    Else
       If CCH.Value = 1 And TCH.Value = 0 Then
          mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\cch_sts" + Chr(34)
          msg = "select * from Cch_sts  where col1=" + Chr(34) + List2.List(0) + Chr(34) + " OR col1=" + Chr(34) + List2.List(1) + Chr(34) + " OR col1=" + Chr(34) + List2.List(2) + Chr(34) + " OR col1=" + Chr(34) + List2.List(3) + Chr(34) + " OR col1=" + Chr(34) + List2.List(4) + Chr(34) + " OR col1=" + Chr(34) + List2.List(5) + Chr(34) + " OR col1=" + Chr(34) + List2.List(6) + Chr(34) + " Into CCH"
          mapinfo.do msg
          mapinfo.do "browse *  from cch  "
          mapinfo.do "set window Frontwindow() Position(0,250) Width 600 Height 160 "
        End If
    End If

    If TCH.Value = 1 And CCH.Value = 1 Then
       mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\tch_sts" + Chr(34)
       msg = "select * from tch_sts  where col1=" + Chr(34) + List2.List(0) + Chr(34) + " OR col1=" + Chr(34) + List2.List(1) + Chr(34) + " OR col1=" + Chr(34) + List2.List(2) + Chr(34) + " OR col1=" + Chr(34) + List2.List(3) + Chr(34) + " OR col1=" + Chr(34) + List2.List(4) + Chr(34) + " OR col1=" + Chr(34) + List2.List(5) + Chr(34) + " OR col1=" + Chr(34) + List2.List(6) + Chr(34) + " Into TCH"
       mapinfo.do msg

       mapinfo.do "browse *  from tch  "
       mapinfo.do "set window Frontwindow() Position(0,250) Width 600 Height 80 "

       mapinfo.do "Set Next Document Parent " & MDIMain.hwnd & " Style 2"
          mapinfo.do "open table " + Chr(34) + Gsm_Path + "\sts\cch_sts" + Chr(34)
          msg = "select * from Cch_sts  where col1=" + Chr(34) + List2.List(0) + Chr(34) + " OR col1=" + Chr(34) + List2.List(1) + Chr(34) + " OR col1=" + Chr(34) + List2.List(2) + Chr(34) + " OR col1=" + Chr(34) + List2.List(3) + Chr(34) + " OR col1=" + Chr(34) + List2.List(4) + Chr(34) + " OR col1=" + Chr(34) + List2.List(5) + Chr(34) + " OR col1=" + Chr(34) + List2.List(6) + Chr(34) + " Into CCH"
          mapinfo.do msg
          mapinfo.do "browse *  from cch  "
          mapinfo.do "set window Frontwindow() Position(0,335) Width 600 Height 80 "

    End If
End If
   Unload Me
End Sub

