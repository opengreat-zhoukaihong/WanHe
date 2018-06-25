VERSION 5.00
Begin VB.Form STREET_DATA 
   BackColor       =   &H00C0C0C0&
   Caption         =   "选择数据表"
   ClientHeight    =   3180
   ClientLeft      =   2715
   ClientTop       =   3210
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3180
   ScaleWidth      =   4995
   Begin VB.ListBox TblList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&C 取消"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3375
      TabIndex        =   1
      Top             =   1290
      Width           =   1230
   End
   Begin VB.CommandButton OK 
      Caption         =   "&O 确认"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "可选数据表："
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "STREET_DATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Dim item   As String
 Dim i, row As Integer
 
 On Error Resume Next
'             TblList.Clear
'             mapinfo.do " fetch  first from  " & Sel_Street
'             row = Val(mapinfo.eval("tableinfo(" & Sel_Street & ",8)"))
'             i = 0
'             Msg = ""
 '            While i < row
 ''               item = mapinfo.eval(Sel_Street + ".datafile")
 '               If Msg <> item And item <> "" Then
 '                  TblList.AddItem item
 '                  Msg = item
  '              End If
  ''              mapinfo.do " fetch  next from  " & Sel_Street
  '              i = i + 1
   '          Wend


 Select Case Menu_Flag
        Case 2001
                 STREET_DATA.Caption = "分析数据表选择"
        Case 2002
                 STREET_DATA.Caption = " 统计数据表选择"
        Case 2003
                 STREET_DATA.Caption = "网调报告数据表选择"
 End Select

 On Error GoTo 0
             mapinfo.do " fetch  first from  static"

             TblList.Clear
             row = Val(mapinfo.eval("tableinfo(static,8)"))
             i = 0
             msg = ""
             While i < row
                item = mapinfo.eval("static.col2")
                msg1 = mapinfo.eval("static.col1")
                If item <> "" And msg1 = str_name Then
                   TblList.AddItem item
                End If
                msg = item
                mapinfo.do " fetch  next from  static"
                i = i + 1
             Wend

End Sub

Private Sub OK_Click()
 Dim StrFile    As String
 On Error Resume Next

 Select Case Menu_Flag
     Case 2001
            tblname = TblList.Text
            Unload Me
            If tblname <> "" Then
               StrFile = tblname + Mid(Sel_Street, 5, 4)
               mapinfo.do "select * from " & Sel_Street & " where col1 = " + Chr(34) + tblname + Chr(34) + " into  " & StrFile
               mapinfo.do "Add Map window FrontWindow() Layer  " & StrFile
            End If
     Case 2002
             tblname = TblList.Text
             Unload Me

             static_dis.Show 1
     Case 2003
            tblname = TblList.Text
            Unload Me
            Call My_Report
 End Select
End Sub

