VERSION 5.00
Begin VB.Form CellGraph 
   Caption         =   "小区天线图象观察"
   ClientHeight    =   4635
   ClientLeft      =   4860
   ClientTop       =   3900
   ClientWidth     =   6975
   Icon            =   "Cellgrap.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4635
   ScaleWidth      =   6975
   Begin VB.Image My_Pic 
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "CellGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim msg ' Declare variables.

    On Error Resume Next    ' Set up error handling.
    msg = Gsm_Path + "\bmp\" + Bmp_Name
    My_Pic.Picture = LoadPicture(msg)  ' Load bitmap.
    If Err Then
        msg = "找不到图象文件" + Bmp_Name + "  !"
        MsgBox msg, 64, "提示" ' Display error message.
        Unload Me
    End If
End Sub

