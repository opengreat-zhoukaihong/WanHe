VERSION 5.00
Begin VB.Form Legend_Frm 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Í¼Àý"
   ClientHeight    =   3450
   ClientLeft      =   1890
   ClientTop       =   1575
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Legend_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Dim Legend_Hwnd As Variant
    Dim i As Integer
    Dim WinId As Variant
    
    On Error Resume Next
            
    Legend_Hwnd = MapInfo.eval("windowinfo(1009,12)")
    'For i = 1 To MapInfo.eval("NumWindows()")
    '    If MapInfo.eval("windowinfo(" & MapInfo.eval("windowid(" & i & ")") & ",3)") = 1009 Then
    '       Legend_Hwnd = MapInfo.eval("windowid(" & i & ")")
    '    End If
    'Next
    'i = MoveWindow(Legend_Hwnd, 0, 0, ScaleWidth, ScaleHeight, 0)
    'MapInfo.do "set window " & Legend_Hwnd & " width " & ScaleWidth & " height " & ScaleHeight
End Sub
