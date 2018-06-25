VERSION 5.00
Begin VB.Form MapForm 
   Caption         =   "  "
   ClientHeight    =   4080
   ClientLeft      =   2100
   ClientTop       =   1515
   ClientWidth     =   7170
   Icon            =   "Mapform.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   272
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   Visible         =   0   'False
End
Attribute VB_Name = "MapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
    Dim mapHWnd As Long, j As Long
    Dim i As Integer
    Dim MyLayout As Long
    
    On Error Resume Next
    If mapid = 0 Then
        For i = 1 To mapinfo.eval("NumWindows()")
            If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
               mapid = mapinfo.eval("windowid(" & i & ")")
               If mapid = mapinfo.eval("frontwindow()") Then
                  Exit For
               End If
            End If
        Next
    End If
    If Me.WindowState <> 1 And mapid > 0 Then
        If thereIsAMap Then
            mapHWnd = Val(mapinfo.eval("WindowInfo(" & mapid & ",12)"))
            If Not IsSetMyHook Then
               'Call SetMyHook(mapHWnd)
               IsSetMyHook = True
            End If
            j = MoveWindow(mapHWnd, 0, 0, ScaleWidth, ScaleHeight, 0)
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If IsSetMyHook Then
       'UnMyHook
       IsSetMyHook = False
    End If
    mapinfo.runmenucommand 104
    thereIsAMap = 0
End Sub
