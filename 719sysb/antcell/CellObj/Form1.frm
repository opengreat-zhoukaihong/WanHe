VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mapinfo As Variant
Dim wInt As Integer

  
Private Sub Command1_Click()
 
   Set mapinfo = CreateObject("MapInfo.Application")
    mapinfo.do "Set Application WIndow " & Form1.hWnd
    
    wInt = cell(mapinfo)
    MsgBox wInt
    'mapinfo.do "Run Menu Command ID 102"
    

   
End Sub

