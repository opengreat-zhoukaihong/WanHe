VERSION 5.00
Begin VB.Form frmCallEvent 
   Caption         =   "通话过程事件"
   ClientHeight    =   2955
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   3015
   FillColor       =   &H00FF0000&
   ForeColor       =   &H000000FF&
   Icon            =   "frmCallEvent.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2955
   ScaleWidth      =   3015
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2835
      IntegralHeight  =   0   'False
      ItemData        =   "frmCallEvent.frx":000C
      Left            =   0
      List            =   "frmCallEvent.frx":000E
      TabIndex        =   0
      Top             =   0
      Width           =   2910
   End
End
Attribute VB_Name = "frmCallEvent"
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

