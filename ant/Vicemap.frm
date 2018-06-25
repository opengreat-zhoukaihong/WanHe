VERSION 5.00
Begin VB.Form ViceMapForm 
   Caption         =   " "
   ClientHeight    =   4230
   ClientLeft      =   2055
   ClientTop       =   2595
   ClientWidth     =   6705
   Icon            =   "Vicemap.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
End
Attribute VB_Name = "ViceMapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    On Error Resume Next
    Dim mapHWnd, j As Long
    If MapForm.WindowState <> 1 Then
        If thereIsAMap Then
'            MessageBeep (2)
            On Error GoTo Go_OUT
            mapHWnd = Val(mapinfo.eval("WindowInfo(frontwindow(),12)"))
            On Error Resume Next
            j = MoveWindow(mapHWnd, 0, 0, ScaleWidth, ScaleHeight, 0)
        End If
    End If
Go_OUT:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mapHWnd   As Long
 '   mapHWnd = Val(mapinfo.Eval("WindowInfo(frontwindow(),12)"))
    On Error Resume Next
    If Map_No > 1 Then
       Map_No = Map_No - 1
    End If
    
'    mapinfo.do "close window " & mapHWnd
End Sub
