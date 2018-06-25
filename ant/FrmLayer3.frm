VERSION 5.00
Begin VB.Form FrmLayer3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "第三层信息"
   ClientHeight    =   3945
   ClientLeft      =   7185
   ClientTop       =   4140
   ClientWidth     =   3330
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLayer3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3330
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
      Height          =   3915
      IntegralHeight  =   0   'False
      ItemData        =   "FrmLayer3.frx":000C
      Left            =   -15
      List            =   "FrmLayer3.frx":000E
      TabIndex        =   0
      Top             =   -15
      Width           =   3300
   End
End
Attribute VB_Name = "FrmLayer3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Me.Width = 3400
    Me.Height = 3500
    Me.Left = 10000 '2000
    Me.Top = 1800 '4500
    
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
    If List1.ListIndex > -1 Then
            SelectList = List1.ListIndex + 1
     End If
End Sub

