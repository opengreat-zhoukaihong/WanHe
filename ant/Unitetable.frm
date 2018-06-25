VERSION 5.00
Begin VB.Form UniteTable 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "合并文件"
   ClientHeight    =   2865
   ClientLeft      =   3450
   ClientTop       =   2850
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Unitetable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   4125
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   525
      TabIndex        =   5
      Top             =   2475
      Width           =   2190
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   2880
      TabIndex        =   2
      Top             =   825
      Width           =   1080
   End
   Begin VB.CommandButton OK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   2880
      TabIndex        =   1
      Top             =   420
      Width           =   1080
   End
   Begin VB.ListBox TblList 
      Height          =   1680
      Left            =   525
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   405
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "到文件："
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   4
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "添加文件："
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   3
      Top             =   135
      Width           =   900
   End
End
Attribute VB_Name = "UniteTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim item As String
    Dim i As Integer
    On Error Resume Next
    mapinfo.do "jj=1"
    TblList.Clear
    TableNum = mapinfo.eval("NumTables()")
    i = 0
    While i < TableNum
       item = mapinfo.eval("tableinfo(jj,1)")
       Select Case UCase(item)
          'Case "CELL", "BASE", "BASE_ADD", "STREET", "AREA", "LANDMARK", "WATER", "MOUNTAIN", "REPETER", "COMPBASE", "COMPCELL", "PUBLIC", "VIP", "POST", "USER_1", "USER_2", "USER_3", "BLOCK", "TOWN"
          Case "GSMCELL", "DCSCELL", "CELL", "BASE", "BASE_ADD", "STREET", "AREA", "LANDMARK", "WATER", "MOUNTAIN", "REPETER", "COMPBASE", "COMPCELL", "PUBLIC", "VIP", "POST", "USER_1", "USER_2", "USER_3", "BLOCK", "TOWN", "DUPLABEL", "DUPLICATE"
          Case Else
              TblList.AddItem item
              Combo1.AddItem item
       End Select
       mapinfo.do " jj=jj+1 "
       i = i + 1
    Wend
    Combo1.Text = Combo1.List(1)
End Sub

Private Sub OK_Click()
   Dim i As Integer, j As Integer
   Dim FirstFile As String
   Dim firstCols As Integer
   Dim MyCommand As String
   
   On Error Resume Next
   
   Me.Hide
   FirstFile = Trim(Combo1.Text)
   If FirstFile <> "" And TblList.SelCount > 0 Then
      firstCols = mapinfo.eval("tableinfo(" & FirstFile & " ,4)")
      For i = 0 To TblList.ListCount - 1
          If TblList.Selected(i) Then
             If firstCols = 59 Then
                If mapinfo.eval("tableinfo(" & TblList.List(i) & " ,4)") = 59 Then
                   mapinfo.do "Insert Into " & FirstFile & " ( COL1, COL2, COL3, COL4, COL5, COL6, COL7, COL8, COL9, COL10, COL11, COL12, COL13, COL14, COL15, COL16, COL17, COL18, COL19, COL20, COL21, COL22, COL23, COL24, COL25, COL26, COL27, COL28, COL29, COL30, COL31, COL32, COL33, COL34, COL35, COL36, COL37, COL38, COL39, COL40, COL41, COL42, COL43, COL44, COL45, COL46, COL47, COL48, COL49, COL50, COL51, COL52, COL53, COL54, COL55, COL56, COL57, COL58, COL59) Select COL1, COL2, COL3, COL4, COL5, COL6, COL7, COL8, COL9, COL10, COL11, COL12, COL13, COL14, COL15, COL16, COL17, COL18, COL19, COL20, COL21, COL22, COL23, COL24, COL25, COL26, COL27, COL28, COL29, COL30, COL31, COL32, COL33, COL34, COL35, COL36, COL37, COL38, COL39, COL40, COL41, COL42, COL43, COL44, COL45, COL46, COL47, COL48, COL49, COL50, COL51, COL52, COL53, COL54, COL55, COL56, COL57, COL58, COL59 From " & TblList.List(i)
                End If
             ElseIf firstCols = 76 Then
                If mapinfo.eval("tableinfo(" & TblList.List(i) & " ,4)") = 76 Then
                   MyCommand = "Insert Into " & FirstFile & " (COL1,COL2,COL3,COL4,COL5,COL6,COL7,COL8,COL9,COL10,COL11,COL12,COL13,COL14,COL15,COL16,COL17,COL18,COL19,COL20,COL21,COL22,COL23,COL24,COL25,COL26,COL27,COL28,COL29,COL30,COL31,COL32,COL33,COL34,COL35,COL36,COL37,COL38,COL39,COL40,COL41,COL42,COL43,COL44,COL45,COL46,COL47,COL48,COL49,COL50,COL51,COL52,COL53,COL54, COL55, COL56, COL57, COL58, COL59, COL60, COL61, COL62, COL63, COL64, COL65, COL66, COL67, COL68, COL69, COL70, COL71, COL72, COL73, COL74, COL75, COL76) Select COL1, COL2, COL3, COL4, COL5, COL6, COL7, COL8, COL9, COL10, COL11, COL12, COL13, COL14, COL15, COL16, COL17, COL18, COL19, COL20, COL21, COL22, COL23, COL24, COL25, COL26, COL27, COL28, COL29, COL30, COL31, COL32, COL33, COL34, COL35, COL36, COL37, COL38, COL39, COL40, COL41, COL42, COL43, COL44, COL45, COL46, COL47, COL48, COL49, COL50, COL51, COL52, COL53, COL54, COL55, COL56, COL57, COL58, COL59 "
                   MyCommand = MyCommand & " , COL60, COL61, COL62, COL63, COL64, COL65, COL66, COL67, COL68, COL69, COL70, COL71, COL72, COL73, COL74, COL75, COL76 From " & TblList.List(i)
                   mapinfo.do MyCommand
                End If
             ElseIf firstCols = 88 Then
                If mapinfo.eval("tableinfo(" & TblList.List(i) & " ,4)") = 88 Then
                   MyCommand = "Insert Into " & FirstFile & " (COL1,COL2,COL3,COL4,COL5,COL6,COL7,COL8,COL9,COL10,COL11,COL12,COL13,COL14,COL15,COL16,COL17,COL18,COL19,COL20,COL21,COL22,COL23,COL24,COL25,COL26,COL27,COL28,COL29,COL30,COL31,COL32,COL33,COL34,COL35,COL36,COL37,COL38,COL39,COL40,COL41,COL42,COL43,COL44,COL45,COL46,COL47,COL48,COL49,COL50,COL51,COL52,COL53,COL54, COL55, COL56, COL57, COL58, COL59, COL60, COL61, COL62, COL63, COL64, COL65, COL66, COL67, COL68, COL69, COL70, COL71, COL72, COL73, COL74, COL75, COL76,COL77,COL78,COL79,COL80,COL81, COL82, COL83, COL84, COL85, COL86, COL87, COL88) Select COL1, COL2, COL3, COL4, COL5, COL6, COL7, COL8, COL9, COL10, COL11, COL12, COL13, COL14, COL15, COL16, COL17, COL18, COL19, COL20, COL21, COL22, COL23, COL24, COL25, COL26, COL27, COL28, COL29, COL30, COL31, COL32, COL33, COL34, COL35, COL36, COL37, COL38, COL39, COL40, COL41, COL42, COL43, COL44, COL45, COL46, COL47, COL48, COL49, COL50, COL51, COL52, COL53, COL54, COL55, COL56, COL57, COL58, COL59 "
                   MyCommand = MyCommand & " , COL60, COL61, COL62, COL63, COL64, COL65, COL66, COL67, COL68, COL69, COL70, COL71, COL72, COL73, COL74, COL75, COL76,COL77,COL78,COL79,COL80,COL81, COL82, COL83, COL84, COL85, COL86, COL87, COL88 From " & TblList.List(i)
                   mapinfo.do MyCommand
                End If
             ElseIf firstCols = 150 Then
                If mapinfo.eval("tableinfo(" & TblList.List(i) & " ,4)") = 150 Then
                   MyCommand = "Insert Into " & FirstFile & " ("
                   For j = 1 To 149
                       MyCommand = MyCommand & "COL" & Format(j) & ","
                   Next
                   MyCommand = MyCommand & "COL150) Select "
                   For j = 1 To 149
                       MyCommand = MyCommand & "COL" & Format(j) & ","
                   Next
                   MyCommand = MyCommand & "COL150 From " & TblList.List(i)
                   mapinfo.do MyCommand
                End If
             End If
          End If
      Next
   End If
   mapinfo.do "commit table " & FirstFile
   Unload Me
End Sub
