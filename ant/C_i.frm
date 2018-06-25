VERSION 5.00
Begin VB.Form C_I 
   BackColor       =   &H00C0C0C0&
   Caption         =   "C/I 选择"
   ClientHeight    =   2625
   ClientLeft      =   3330
   ClientTop       =   1665
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "C_i.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "C/I 选择"
      Height          =   1770
      Left            =   315
      TabIndex        =   5
      Top             =   150
      Width           =   2970
      Begin VB.ComboBox ARF_VALUE 
         Height          =   300
         Index           =   2
         Left            =   1260
         TabIndex        =   2
         Text            =   " "
         Top             =   1260
         Width           =   1365
      End
      Begin VB.ComboBox ARF_VALUE 
         Height          =   300
         Index           =   1
         Left            =   1245
         TabIndex        =   1
         Text            =   " "
         Top             =   825
         Width           =   1365
      End
      Begin VB.ComboBox ARF_VALUE 
         Height          =   300
         Index           =   0
         Left            =   1245
         TabIndex        =   0
         Text            =   " "
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "I2(干扰):"
         Height          =   180
         Index           =   2
         Left            =   345
         TabIndex        =   8
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "I1(干扰):"
         Height          =   180
         Index           =   1
         Left            =   345
         TabIndex        =   7
         Top             =   870
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   " C(载频):"
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   6
         Top             =   450
         Width           =   810
      End
   End
   Begin VB.CommandButton SBSCancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   1905
      TabIndex        =   4
      Top             =   2175
      Width           =   1080
   End
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   675
      TabIndex        =   3
      Top             =   2175
      Width           =   1080
   End
End
Attribute VB_Name = "C_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim col_num, i As Integer
    Dim str As String

    On Error Resume Next
    
    msg = "TableInfo(""" & tblname & """, 4)"
    col_num = Val(mapinfo.eval(msg))
    For i = 4 To col_num Step 2
        msg = "Columninfo(""" & tblname & """,""COL" & i & """, 1)"
        On Error Resume Next
        str = mapinfo.eval(msg)
        
        ARF_VALUE(0).AddItem str
        ARF_VALUE(1).AddItem str
        ARF_VALUE(2).AddItem str
    Next i
    ARF_VALUE(0).Text = str
    ARF_VALUE(1).Text = str
    ARF_VALUE(2).Text = str
   
End Sub

Private Sub SBSCancel_Click()
    Unload Me
End Sub


Private Sub SBSOK_Click()
    Dim bbb, c, I1, I2, c_i1, c_i2 As String
    Dim X, Y, z As Integer
    On Error Resume Next
    
    Screen.MousePointer = 11
    c = ARF_VALUE(0).Text
    I1 = ARF_VALUE(1).Text
    I2 = ARF_VALUE(2).Text
    Unload Me
    bbb = tblname + "C"
    If Left(bbb, 1) = "_" Then
       aaa = Mid(bbb, 2, Len(bbb))
    Else
       aaa = bbb
    End If
    ccc = Gsm_Path + "\scan\" + aaa + ".dbf"
    Gsm_FileName = Gsm_Path + "\c_i.dbf"
    FileCopy Gsm_FileName, ccc
    ccc = Gsm_Path + "\scan\" + aaa + ".tab"
    Gsm_FileName = Gsm_Path + "\c_i.tab"
    FileCopy Gsm_FileName, ccc

    mapinfo.do "open  table " + Chr(34) + ccc + Chr(34)
    mapinfo.do "insert into  " & bbb & "  (col1,col2,col3) Select col1,col2,col3 from   " & tblname
    mapinfo.do "Create Map For  " & bbb & "  CoordSys Earth Projection 1, 0"
    mapinfo.do "update " + bbb + " set Obj= CreatePoint(Lon, Lat)"
    i = 1
    row = Val(mapinfo.eval("tableinfo(" & tblname & ",8)"))
    mapinfo.do "fetch First from " & tblname
    mapinfo.do "fetch First from " & bbb
    While i <= row
        msg = tblname + "." + c
        X = Val(mapinfo.eval(msg))
        msg = tblname + "." + I1
        Y = Val(mapinfo.eval(msg))
        msg = tblname + "." + I2
        z = Val(mapinfo.eval(msg))
        c_i1 = CStr(X - Y)
        c_i2 = CStr(X - z)
        msg = "update  " & bbb & " set col4 =" & c_i1 & ",col5 =" & c_i2 & " Where rowid=" & i
        mapinfo.do msg
        mapinfo.do "fetch next from " & tblname
        mapinfo.do "fetch next from " & bbb
        i = i + 1
    Wend
    mapinfo.do "Create Map For  " & bbb & "  CoordSys Earth Projection 1, 0"
    mapinfo.do "Set Style Symbol MakeSymbol(33,0,2)"
    mapinfo.do "update " + bbb + " set Obj= CreatePoint(Lon, Lat)"
    Screen.MousePointer = 0
    mapinfo.do "commit  table " & bbb
    mapinfo.do "close  table " & bbb
End Sub

