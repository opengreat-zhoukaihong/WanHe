VERSION 5.00
Begin VB.Form Cope_Define 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�����ֻ�ѡ��"
   ClientHeight    =   2070
   ClientLeft      =   3090
   ClientTop       =   3060
   ClientWidth     =   3225
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Cope_def.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2070
   ScaleWidth      =   3225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   375
      TabIndex        =   2
      Top             =   120
      Width           =   2400
      Begin VB.OptionButton Option1 
         Caption         =   "M1 ������M2 ����"
         Height          =   240
         Left            =   315
         TabIndex        =   4
         Top             =   390
         Value           =   -1  'True
         Width           =   1770
      End
      Begin VB.OptionButton Option2 
         Caption         =   "M2 ������M1 ����"
         Height          =   240
         Left            =   315
         TabIndex        =   3
         Top             =   780
         Width           =   1755
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   320
      Left            =   1650
      TabIndex        =   1
      Top             =   1650
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   320
      Left            =   450
      TabIndex        =   0
      Top             =   1650
      Width           =   1080
   End
End
Attribute VB_Name = "Cope_Define"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    On Error Resume Next
    If Option1.Value = True Then
       M2_Local = False
    Else
       M2_Local = True
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    M2_Local = False
End Sub

