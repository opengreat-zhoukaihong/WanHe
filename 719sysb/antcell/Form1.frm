VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   1665
   ClientTop       =   1530
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6690
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   660
      Left            =   1920
      TabIndex        =   0
      Top             =   2130
      Width           =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mapinfo As Variant

Private Sub Command1_Click()
    Dim MyResult As Boolean
    
    Set mapinfo = CreateObject("mapinfo.Application")
    mapinfo.do "Set Next Document Parent " & Form1.hWnd & " Style 1"
    mapinfo.do "open table ""C:\My Documents\antcell\map\cell.tab"""
    mapinfo.do "map from cell"
    mapinfo.do "open table ""C:\My Documents\antcell\map\area.tab"""
    mapinfo.do "map from area"
    MyResult = CreateObj(mapinfo, "ant")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mapinfo = Nothing
End Sub
