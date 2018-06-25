VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form Idle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IDLE数据转移"
   ClientHeight    =   1500
   ClientLeft      =   795
   ClientTop       =   6660
   ClientWidth     =   4620
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Idle.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1500
   ScaleWidth      =   4620
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3675
      Top             =   45
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   320
      Left            =   1680
      TabIndex        =   1
      Top             =   1095
      Width           =   1080
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   300
      TabIndex        =   0
      Top             =   615
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   476
      _Version        =   327680
      Appearance      =   0
      MouseIcon       =   "Idle.frx":030A
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "正在处理"
      Height          =   180
      Left            =   300
      TabIndex        =   2
      Top             =   210
      Width           =   720
   End
End
Attribute VB_Name = "Idle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    On Error Resume Next
    If (MsgBox("确实要中止转换吗？", 33, "提示")) = 1 Then
       Convert_Stop = True
    End If
End Sub

Private Sub Timer1_Timer()
    Idle_Function
    Unload Me
End Sub

Sub Idle_Function()
    Dim Tab_Rows As Variant
    Dim Convert_UseName As String, SaveAsName As String
    Dim percent_step As Integer, nline As Integer
    Dim i As Integer, j As Integer, finds As Integer

    On Error Resume Next
    i = 1
    Convert_Stop = False
    Is_Done = False
    Do While convert_filename(i) <> ""
       Idle.Label1.Caption = "正在处理 " + convert_filename(i)
       Idle.ProgressBar1.Value = 0
       mapinfo.do "open table " + Chr(34) + convert_filename(i) + Chr(34)
       Convert_UseName = convert_filename(i)
       finds = InStr(Convert_UseName, ".")
       If finds > 0 Then
          Convert_UseName = Left(Convert_UseName, finds - 1)
       End If
       finds = InStr(Convert_UseName, "\")
       Do While finds > 0
          Convert_UseName = Right(Convert_UseName, Len(Convert_UseName) - finds)
          finds = InStr(Convert_UseName, "\")
       Loop
       If Asc(Left(Convert_UseName, 1)) > 47 And Asc(Left(Convert_UseName, 1)) < 58 Then
          Convert_UseName = "_" + Convert_UseName
       End If
       If Val(mapinfo.eval("tableinfo(" & Convert_UseName & ",4)")) <> 59 Then
          MsgBox "文件格式错误 " + convert_filename(i), 48, "提示"
          GoTo next_file
       End If
       SaveAsName = convert_filename(i)
       finds = InStr(SaveAsName, ".")
       If finds > 0 Then
          SaveAsName = Left(SaveAsName, finds - 1)
       End If
       SaveAsName = SaveAsName + "M.tab"
       mapinfo.do "commit table " & Convert_UseName & " as " + Chr(34) + SaveAsName + Chr(34)
       mapinfo.do "close  table " & Convert_UseName
       Convert_UseName = Convert_UseName + "M"
       mapinfo.do "open table " + Chr(34) + SaveAsName + Chr(34)
       Tab_Rows = mapinfo.eval("tableinfo(" & Convert_UseName & ",8)")
       percent_step = Int(Tab_Rows / 100)
       If percent_step = 0 Then
          percent_step = 1
       End If
       mapinfo.do "fetch first from " & Convert_UseName
       Is_Done = True
       nline = 0
       For j = 1 To Tab_Rows
           nline = nline + 1
           If nline = percent_step And Idle.ProgressBar1.Value + 1 < 100 Then
              DoEvents
              If Convert_Stop = True Then
                 mapinfo.do "commit table " & Convert_UseName
                 mapinfo.do "close table " & Convert_UseName
                 Exit Do
              End If
              Idle.ProgressBar1.Value = Idle.ProgressBar1.Value + 1
              nline = 0
           End If
           If Val(mapinfo.eval(Convert_UseName & ".Rxlev_s")) = 0 Then
              mapinfo.do "update " & Convert_UseName & " set Rxlev_S = Rxlev_f Where rowid=" & j
           End If
           mapinfo.do "fetch next from " & Convert_UseName
       Next
       mapinfo.do "commit table " & Convert_UseName
       mapinfo.do "close table " & Convert_UseName
       
next_file:
       i = i + 1
    Loop
End Sub
