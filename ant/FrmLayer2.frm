VERSION 5.00
Begin VB.Form FrmLayer2 
   Caption         =   "通话过程事件"
   ClientHeight    =   2955
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   3015
   Icon            =   "FrmLayer2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2955
   ScaleWidth      =   3015
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   2835
      IntegralHeight  =   0   'False
      ItemData        =   "FrmLayer2.frx":000C
      Left            =   0
      List            =   "FrmLayer2.frx":000E
      TabIndex        =   0
      Top             =   0
      Width           =   2910
   End
End
Attribute VB_Name = "FrmLayer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Me.Width = 3000
    Me.Height = 3500
    Me.Left = 10000 ' 4000
    Me.Top = 3600 '4500
    
    Width = List1.Width
    Height = List1.Height
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = ScaleWidth
    List1.Height = ScaleHeight

End Sub

Private Sub List1_Click()
    Dim SelectList As Integer
    
    On Error Resume Next
    SelectList = List1.ListIndex + 1

    
End Sub

