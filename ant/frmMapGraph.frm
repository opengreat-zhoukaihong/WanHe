VERSION 5.00
Begin VB.Form frmMapGraph 
   Caption         =   "参数统计图"
   ClientHeight    =   2580
   ClientLeft      =   3405
   ClientTop       =   4695
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMapGraph.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   172
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   330
End
Attribute VB_Name = "frmMapGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    MapGraphflag = True
End Sub

Private Sub Form_Resize()
    Dim MyWinId As Long
    Dim i As Integer
    Dim MyWinHwnd As Long
    Dim MyResult As Long

    On Error Resume Next
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 4 Then
            MyWinId = mapinfo.eval("windowid(" & i & ")")
            Exit For
        End If
    Next
    If MyWinId = 0 Then
       Exit Sub
    End If
    MyWinHwnd = mapinfo.eval("WindowInfo(" & MyWinId & ",12)")
    MyResult = MoveWindow(MyWinHwnd, 0, 0, ScaleWidth, ScaleHeight, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim MyTableNum As Integer, i As Integer

    On Error Resume Next
    MapGraphflag = False
FindAgain:
    'MyTableNum = mapinfo.eval("NumTables()")
    'For i = 1 To MyTableNum
    '    If Len(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) > 5 Then
    '       If Left(UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")), 5) = "QUERY" Or UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "SELECTION" Then
    '          mapinfo.Do "close table tableinfo(" & Format(i) & ",1)"
    '          GoTo FindAgain
    '       End If
    '    End If
    'Next

End Sub
