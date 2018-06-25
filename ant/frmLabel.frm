VERSION 5.00
Begin VB.Form frmLabel 
   Caption         =   "参数标注"
   ClientHeight    =   3525
   ClientLeft      =   3690
   ClientTop       =   3165
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4005
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option2 
      Caption         =   "左右"
      Height          =   240
      Left            =   1155
      TabIndex        =   5
      Top             =   3150
      Width           =   720
   End
   Begin VB.OptionButton Option1 
      Caption         =   "上下"
      Height          =   240
      Left            =   195
      TabIndex        =   4
      Top             =   3150
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton OK 
      Caption         =   "&O 确认"
      Height          =   320
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   1080
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&C 取消"
      Height          =   320
      Left            =   2760
      TabIndex        =   1
      Top             =   885
      Width           =   1080
   End
   Begin VB.ListBox List1 
      Height          =   2160
      Left            =   210
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择参数："
      Height          =   180
      Index           =   1
      Left            =   195
      TabIndex        =   6
      Top             =   195
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "标注位置："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   2805
      Width           =   900
   End
End
Attribute VB_Name = "frmLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    On Error Resume Next
    For i = 1 To Val(mapinfo.eval("tableinfo( " & tblname & " ,4)"))
        List1.AddItem mapinfo.eval("Columninfo( " & tblname & ",COL" & Format(i) & ", 1)")
    Next
    
End Sub

Private Sub List1_Click()
    On Error Resume Next
    If List1.SelCount > 2 And List1.Selected(List1.ListIndex) Then
       List1.Selected(List1.ListIndex) = False
    End If
End Sub

Private Sub OK_Click()
    Dim LableField1 As String, LableField2 As String
    Dim i As Integer
    Dim Layers As Integer
    Dim Mymsg As String
    
    On Error Resume Next
    For i = 1 To List1.ListCount
        If List1.Selected(i - 1) Then
           If LableField1 = "" Then
              LableField1 = List1.List(i - 1)
           Else
              LableField2 = List1.List(i - 1)
           End If
        End If
    Next
    If List1.SelCount = 1 Then
       Unload Me
       For i = 1 To Val(mapinfo.eval("NumTables()"))
           If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "DUPLABEL" Or UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "DUPLICATE" Then
              mapinfo.do "close table DupLabel"
           End If
       Next
       mapinfo.do "set map redraw off"
       mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With " & LableField1 & " Auto On Overlap Off Duplicates On Position Above Auto On Offset 2"
       mapinfo.do "set map redraw on"
    ElseIf List1.SelCount = 2 Then
       mapinfo.do "select * from " & tblname & " into DupLabel"
       mapinfo.do "Add Map window FrontWindow() Layer DupLabel"
       Layers = mapinfo.eval("mapperinfo(frontwindow(),9)")
       Mymsg = "set map order "
       For i = 2 To Layers
           Mymsg = Mymsg + Format(i) + ","
       Next
       Mymsg = Mymsg + "1"
       mapinfo.do Mymsg
       
       If Option1.Value = True Then
          Unload Me
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,8421376,16777215) With " & LableField1 & " Auto On Overlap Off Duplicates On Position Above Auto On Offset 10"
          mapinfo.do "set map redraw on"
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer ""DupLabel"" Label Visibility Font (""Arial"",257,8,16711935,16777215) With " & LableField2 & " Auto On Overlap Off Duplicates On Position Below Auto On Offset 10"
          mapinfo.do "set map redraw on"
       Else
          Unload Me
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer " + Chr(34) + tblname + Chr(34) + " Label Visibility Font (""Arial"",257,8,255,16777215) With " & LableField1 & " Auto On Overlap Off Duplicates On Position Left Auto On Offset 10"
          mapinfo.do "set map redraw on"
          mapinfo.do "set map redraw off"
          mapinfo.do "Set Map Layer ""DupLabel"" Label Visibility Font (""Arial"",257,8,16711935,16777215) With " & LableField2 & " Auto On Overlap Off Duplicates On Position Right Auto On Offset 10"
          mapinfo.do "set map redraw on"
       End If
    Else
       Unload Me
    End If
End Sub
