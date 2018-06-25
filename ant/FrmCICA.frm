VERSION 5.00
Begin VB.Form FrmCICA 
   BackColor       =   &H80000009&
   Caption         =   "邻频载干比显示"
   ClientHeight    =   2235
   ClientLeft      =   7425
   ClientTop       =   1515
   ClientWidth     =   3540
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCICA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3540
   Begin VB.CommandButton OK 
      Cancel          =   -1  'True
      Caption         =   "关闭"
      Height          =   320
      Left            =   1770
      TabIndex        =   28
      Top             =   1830
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示 Full"
      Height          =   320
      Left            =   540
      TabIndex        =   27
      Top             =   1830
      Width           =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   3480
      TabIndex        =   0
      Top             =   0
      Width           =   3540
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
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
         Left            =   1440
         TabIndex        =   32
         Top             =   885
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
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
         Left            =   2040
         TabIndex        =   31
         Top             =   885
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
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
         Left            =   2520
         TabIndex        =   30
         Top             =   885
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
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
         Left            =   3090
         TabIndex        =   29
         Top             =   885
         Width           =   45
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   855
         TabIndex        =   26
         Top             =   1380
         Width           =   45
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   135
         TabIndex        =   25
         Top             =   1380
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   300
         TabIndex        =   24
         Top             =   1380
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   3
         Left            =   3090
         TabIndex        =   23
         Top             =   1395
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   2
         Left            =   3090
         TabIndex        =   22
         Top             =   1140
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   3
         Left            =   2520
         TabIndex        =   21
         Top             =   1395
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   2
         Left            =   2520
         TabIndex        =   20
         Top             =   1140
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   3
         Left            =   2040
         TabIndex        =   19
         Top             =   1395
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   2
         Left            =   2040
         TabIndex        =   18
         Top             =   1140
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   3
         Left            =   1425
         TabIndex        =   17
         Top             =   1395
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   2
         Left            =   1425
         TabIndex        =   16
         Top             =   1140
         Width           =   45
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   810
         TabIndex        =   15
         Top             =   1380
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   14
         Top             =   1380
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Top             =   645
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   2520
         TabIndex        =   12
         Top             =   405
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Rxlev"
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
         Left            =   2475
         TabIndex        =   11
         Top             =   120
         Width           =   465
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   1110
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   3090
         TabIndex        =   10
         Top             =   645
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   2040
         TabIndex        =   9
         Top             =   630
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   630
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   1440
         TabIndex        =   7
         Top             =   390
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   390
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   3090
         TabIndex        =   5
         Top             =   390
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "C/A"
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
         Left            =   3060
         TabIndex        =   4
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "BSIC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1950
         TabIndex        =   3
         Top             =   120
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "ARFCN"
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
         Left            =   1290
         TabIndex        =   2
         Top             =   105
         Width           =   570
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "BCCH±1±2"
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
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   1125
      End
   End
End
Attribute VB_Name = "FrmCICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rxle_neig1 As Integer
Dim rxle_neig2 As Integer
Dim rxle_neig3 As Integer
Dim rxle_neig4 As Integer
Dim bsic_neig1 As Integer
Dim bsic_neig2 As Integer
Dim bsic_neig3 As Integer
Dim bsic_neig4 As Integer
Dim myrxlev_f As Integer
Dim myrxlev_s As Integer
Dim mybsic_serv As Integer
Dim mybcch_serv As Integer
Dim MySelName As String

Private Sub Command1_Click()
    Dim MSXpos As Integer
    Dim MSYpos As Integer
    
    On Error Resume Next
    If Command1.Caption = "显示 Full" Then
         Command1.Caption = "显示 Sub"
            Picture1.Cls
            MSXpos = 150
            MSYpos = 1365 - rxle_neig1 * 6.1  '-2
            Picture1.Line (MSXpos, Line1.y1 - 10)-(MSXpos + 120, MSYpos), &HFF0000, BF
            MSXpos = 320 '370 - 50
            MSYpos = 1365 - rxle_neig2 * 6.1   '-1
            Picture1.Line (MSXpos, Line1.y1 - 10)-(MSXpos + 120, MSYpos), &HFF0000, BF
            MSXpos = 490
            MSYpos = 1365 - myrxlev_f * 6.1  '
            Picture1.Line (MSXpos, Line1.y1 - 10)-(MSXpos + 120, MSYpos), &H80000008, BF
            MSXpos = 660
            MSYpos = 1365 - rxle_neig3 * 6.1    '+1
            Picture1.Line (MSXpos, Line1.y1 - 10)-(MSXpos + 120, MSYpos), &HFF0000, BF
            MSXpos = 830
            MSYpos = 1365 - rxle_neig4 * 6.1   '+2
            Picture1.Line (MSXpos, Line1.y1 - 10)-(MSXpos + 120, MSYpos), &HFF0000, BF
            
            Label10.Caption = Format(mybcch_serv)
            'Label11.Caption = "+1"
            'Label12.Caption = "-1"
            'Label13.Caption = "-2"
            'Label14.Caption = "+2"
            Label6(0).Caption = Format(mybcch_serv - 2)
            Label7(0).Caption = Format(bsic_neig1)
            Label6(1).Caption = Format(mybcch_serv - 1)
            Label7(1).Caption = Format(bsic_neig2)
            Label6(2).Caption = Format(mybcch_serv + 1)
            Label7(2).Caption = Format(bsic_neig3)
            Label6(3).Caption = Format(mybcch_serv + 2)
            Label7(3).Caption = Format(bsic_neig4)
            
            Label6(4).Caption = Format(mybcch_serv)
            Label7(4).Caption = Format(mybsic_serv)
            Label9(4).Caption = Format(myrxlev_f)
                
                Label9(0).Caption = Format(rxle_neig1)
                Label9(1).Caption = Format(rxle_neig2)
                Label9(2).Caption = Format(rxle_neig3)
                Label9(3).Caption = Format(rxle_neig4)

                'If rxlev_f) <> 0 Then
                    If myrxlev_f - rxle_neig1 <= -18 Then
                        Label8(0).ForeColor = &HFF&
                    Else
                        Label8(0).ForeColor = &HC00000
                    End If
                    Label8(0).Caption = Format(myrxlev_f - rxle_neig1)
                    If myrxlev_f - rxle_neig2 <= -9 Then
                         Label8(1).ForeColor = &HFF&
                    Else
                          Label8(1).ForeColor = &HC00000
                    End If
                    Label8(1).Caption = Format(myrxlev_f - rxle_neig2)
                    If myrxlev_f - rxle_neig3 <= -9 Then
                         Label8(2).ForeColor = &HFF&
                    Else
                          Label8(2).ForeColor = &HC00000
                    End If
                    Label8(2).Caption = Format(myrxlev_f - rxle_neig3)
                    If myrxlev_f - rxle_neig4 <= -18 Then
                         Label8(3).ForeColor = &HFF&
                    Else
                          Label8(3).ForeColor = &HC00000
                    End If
                    Label8(3).Caption = Format(myrxlev_f - rxle_neig4)
                
              '  End If
    
    Else
       Command1.Caption = "显示 Full"
    
            Picture1.Cls
            MSXpos = 150
            MSYpos = 1365 - rxle_neig1 * 6.1  '-2
            Picture1.Line (MSXpos, Line1.y1 - 10)-(MSXpos + 120, MSYpos), &HFF0000, BF
            MSXpos = 320 '370 - 50
            MSYpos = 1365 - rxle_neig2 * 6.1   '-1
            Picture1.Line (MSXpos, Line1.y1 - 10)-(MSXpos + 120, MSYpos), &HFF0000, BF
            MSXpos = 490
            MSYpos = 1365 - myrxlev_s * 6.1 '
            Picture1.Line (MSXpos, Line1.y1 - 10)-(MSXpos + 120, MSYpos), &H80000008, BF
            MSXpos = 660
            MSYpos = 1365 - rxle_neig3 * 6.1  '+1
            Picture1.Line (MSXpos, Line1.y1 - 10)-(MSXpos + 120, MSYpos), &HFF0000, BF
            MSXpos = 830
            MSYpos = 1365 - rxle_neig4 * 6.1    '+2
            Picture1.Line (MSXpos, Line1.y1 - 10)-(MSXpos + 120, MSYpos), &HFF0000, BF
            
            Label10.Caption = Format(mybcch_serv)
            'Label11.Caption = "+1"
            'Label12.Caption = "-1"
            'Label13.Caption = "-2"
            'Label14.Caption = "+2"
            Label6(0).Caption = Format(mybcch_serv - 2)
            Label7(0).Caption = Format(bsic_neig1)
            Label6(1).Caption = Format(mybcch_serv - 1)
            Label7(1).Caption = Format(bsic_neig2)
            Label6(2).Caption = Format(mybcch_serv + 1)
            Label7(2).Caption = Format(bsic_neig3)
            Label6(3).Caption = Format(mybcch_serv + 2)
            Label7(3).Caption = Format(bsic_neig4)

            Label6(4).Caption = Format(mybcch_serv)
            Label7(4).Caption = Format(mybsic_serv)
            Label9(4).Caption = Format(myrxlev_s)
            
                Label9(0).Caption = Format(rxle_neig1)
                Label9(1).Caption = Format(rxle_neig2)
                Label9(2).Caption = Format(rxle_neig3)
                Label9(3).Caption = Format(rxle_neig4)

'                If myrxlev_s <> 0 Then
                    If myrxlev_s - rxle_neig1 <= -18 Then
                        Label8(0).ForeColor = &HFF&
                    Else
                        Label8(0).ForeColor = &HC00000
                    End If
                    Label8(0).Caption = Format(myrxlev_s - rxle_neig1)
                    If myrxlev_s - rxle_neig2 <= -9 Then
                         Label8(1).ForeColor = &HFF&
                    Else
                          Label8(1).ForeColor = &HC00000
                    End If
                    Label8(1).Caption = Format(myrxlev_s - rxle_neig2)
                    If myrxlev_s - rxle_neig3 <= -9 Then
                         Label8(2).ForeColor = &HFF&
                    Else
                          Label8(2).ForeColor = &HC00000
                    End If
                    Label8(2).Caption = Format(myrxlev_s - rxle_neig3)
                    If myrxlev_s - rxle_neig4 <= -18 Then
                         Label8(3).ForeColor = &HFF&
                    Else
                          Label8(3).ForeColor = &HC00000
                    End If
                    Label8(3).Caption = Format(myrxlev_s - rxle_neig4)
                
'                End If
    End If
  
End Sub

Private Sub Form_Activate()
    
    Command1_Click
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    MySelName = mapinfo.eval("selectioninfo(2)")
    Picture1.Line (60, 1365)-(750, 1365), &H808080
    rxle_neig1 = mapinfo.eval("selection.rxle_neig1")
    rxle_neig2 = mapinfo.eval("selection.rxle_neig2")
    rxle_neig3 = mapinfo.eval("selection.rxle_neig3")
    rxle_neig4 = mapinfo.eval("selection.rxle_neig4")
    bsic_neig1 = mapinfo.eval("selection.bsic_neig1")
    bsic_neig2 = mapinfo.eval("selection.bsic_neig2")
    bsic_neig3 = mapinfo.eval("selection.bsic_neig3")
    bsic_neig4 = mapinfo.eval("selection.bsic_neig4")
    myrxlev_f = mapinfo.eval("selection.rxlev_f")
    myrxlev_s = mapinfo.eval("selection.rxlev_s")
    mybcch_serv = mapinfo.eval("selection.bcch_serv")
    mybsic_serv = mapinfo.eval("selection.bsic_serv")
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    mapinfo.do "close table " & MySelName

End Sub

Private Sub OK_Click()
    On Error Resume Next
    Unload Me
End Sub
