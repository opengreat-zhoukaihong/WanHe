VERSION 5.00
Begin VB.Form LacLegend 
   BackColor       =   &H8000000E&
   Caption         =   "Lac หตร๗"
   ClientHeight    =   2115
   ClientLeft      =   4080
   ClientTop       =   3480
   ClientWidth     =   1920
   Icon            =   "LacLegend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   1920
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   210
      Index           =   4
      Left            =   540
      Shape           =   1  'Square
      Top             =   1665
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1080"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   945
      TabIndex        =   4
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1080"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   945
      TabIndex        =   3
      Top             =   1305
      Width           =   420
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   210
      Index           =   3
      Left            =   540
      Shape           =   1  'Square
      Top             =   1305
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1080"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   945
      TabIndex        =   2
      Top             =   975
      Width           =   420
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   210
      Index           =   2
      Left            =   540
      Shape           =   1  'Square
      Top             =   975
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1080"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   945
      TabIndex        =   1
      Top             =   630
      Width           =   420
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   210
      Index           =   1
      Left            =   540
      Shape           =   1  'Square
      Top             =   630
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1080"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   945
      TabIndex        =   0
      Top             =   285
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000A&
      FillStyle       =   0  'Solid
      Height          =   210
      Index           =   0
      Left            =   540
      Shape           =   1  'Square
      Top             =   285
      Width           =   225
   End
End
Attribute VB_Name = "LacLegend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim LacRow As Integer, i As Integer
     
    On Error Resume Next
    mapinfo.do "SELECT lac FROM base group by lac order by lac desc into mytemp"
    LacRow = Val(mapinfo.eval("TABLEINFO(mytemp, 8)"))
    mapinfo.do "fetch first from mytemp"
    For i = 0 To LacRow - 1
        If mapinfo.eval("mytemp.lac") = 0 Then
            Shape1(i).FillColor = 0
            Shape1(i).BorderColor = 0
            Label1(i).Caption = "N/A"
        Else
            Shape1(i).FillColor = MyLacColor(i)
            Shape1(i).BorderColor = MyLacColor(i)
            Label1(i).Caption = mapinfo.eval("mytemp.lac")
        End If
        mapinfo.do "fetch next from mytemp"
    Next

End Sub
