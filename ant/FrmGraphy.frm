VERSION 5.00
Begin VB.Form FrmGraphy 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Í¼ÐÎÃèÊö"
   ClientHeight    =   4380
   ClientLeft      =   1080
   ClientTop       =   1695
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGraphy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   662
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   7080
      TabIndex        =   3
      Top             =   30
      Width           =   1080
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "FER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   9
         Left            =   45
         TabIndex        =   16
         Top             =   3810
         UseMnemonic     =   0   'False
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   8
         Left            =   810
         TabIndex        =   15
         Top             =   450
         UseMnemonic     =   0   'False
         Width           =   120
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H00C0C000&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   1
         Left            =   585
         Top             =   495
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   7
         Left            =   270
         TabIndex        =   14
         Top             =   450
         UseMnemonic     =   0   'False
         Width           =   120
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00009E9E&
         FillColor       =   &H00009E9E&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   0
         Left            =   45
         Top             =   495
         Width           =   150
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         Caption         =   "TxPwr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   255
         Left            =   30
         TabIndex        =   12
         Top             =   3255
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Ncell->TCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   210
         Index           =   6
         Left            =   15
         TabIndex        =   11
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Ncell->BCCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   210
         Index           =   12
         Left            =   15
         TabIndex        =   10
         Top             =   0
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "TA "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   5
         Left            =   30
         TabIndex        =   9
         Top             =   3525
         UseMnemonic     =   0   'False
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "RxQual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   4
         Left            =   30
         TabIndex        =   8
         Top             =   3030
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "RxLev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   3
         Left            =   30
         TabIndex        =   7
         Top             =   2445
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "RxLev 20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   6
         Top             =   1905
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "RxLev 40"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   5
         Top             =   1365
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "RxLev 60"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00585858&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   825
         Width           =   780
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4410
      Left            =   0
      ScaleHeight     =   290
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   466
      TabIndex        =   0
      Top             =   0
      Width           =   7050
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   1
         Left            =   7110
         ScaleHeight     =   273
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1260
         TabIndex        =   13
         Top             =   0
         Width           =   18900
         Begin VB.Line Line1 
            BorderColor     =   &H80000001&
            Index           =   1
            Visible         =   0   'False
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   290
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4095
         Index           =   0
         Left            =   -15
         ScaleHeight     =   273
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1310
         TabIndex        =   2
         Top             =   0
         Width           =   19650
         Begin VB.Line Line1 
            BorderColor     =   &H80000001&
            Index           =   0
            Visible         =   0   'False
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   290
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   420
         Left            =   0
         Max             =   844
         TabIndex        =   1
         Top             =   4095
         Width           =   7005
      End
   End
End
Attribute VB_Name = "FrmGraphy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Dim i As Integer
        
    On Error Resume Next
    Me.Width = 10050
    Me.Height = 4785
    Me.Left = 0
    Me.Top = -30
    CurrentP = 0
    JustStart = True
    CurrentLeft = -1
    AheadP = 0
    ScrollOffset = 0
    
    For i = 0 To 1
        Picture3(i).Line (0, 208)-(Picture3(i).Width, 208), &H808080
        Picture3(i).Line (0, 169)-(Picture3(i).Width, 169), &H808080
        Picture3(i).Line (0, 133)-(Picture3(i).Width, 133), &H808080
        Picture3(i).Line (0, 97)-(Picture3(i).Width, 97), &H808080
        Picture3(i).Line (0, 61)-(Picture3(i).Width, 61), &H808080
        Picture3(i).Line (0, 13)-(Picture3(i).Width, 13), &HC000C0
        Picture3(i).Line (0, 246)-(Picture3(i).Width, 246), &H808080
    Next
    If SysSetting.IsdBm = "1" Then
       Label1(0).Caption = "-50 dBm"
       Label1(1).Caption = "-70 dBm"
       Label1(2).Caption = "-90 dBm"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    FrmMain.MnuGraphical.Checked = False
    
End Sub

Private Sub HScroll1_Change()
    
    On Error Resume Next
    If JustStart Then
       Picture3(AheadP).Left = -HScroll1.Value - ScrollOffset - 1
    Else
       Picture3(AheadP).Left = -HScroll1.Value - ScrollOffset - 1
       Picture3((AheadP + 1) Mod 2).Left = Picture3(AheadP).Left + Picture3(AheadP).Width - 1
    End If
    
End Sub

Private Sub HScroll1_Scroll()
    
    On Error Resume Next
    If JustStart Then
       Picture3(AheadP).Left = -HScroll1.Value - ScrollOffset - 1
    Else
       Picture3(AheadP).Left = -HScroll1.Value - ScrollOffset - 1
       Picture3((AheadP + 1) Mod 2).Left = Picture3(AheadP).Left + Picture3(AheadP).Width - 1
    End If

End Sub


Private Sub Picture3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    On Error Resume Next
    If Button = 2 Then
       PopupMenu FrmMain.MnuGraphControl
    ElseIf Button = 1 Then
       If JustStart Then
          If OldPosX >= X Then
             Call DisplayValue(X, Index)
          Else
             MyHideLine
          End If
       Else
          If Index = AheadP Then
             Call DisplayValue(X + 1, Index)
          Else
             If OldPosX >= X Then
                Call DisplayValue(X + 1, Index)
             Else
                MyHideLine
             End If
          End If
       End If
    End If

End Sub

Private Sub DisplayValue(MyPos As Single, MyIndex As Integer)
    Dim i As Integer
    
    On Error Resume Next
    Label3(0).Caption = Trim(SampleData(Int(MyPos / 2 + 0.5)).num_frame)
    Label6.Caption = Trim(SampleData(Int(MyPos / 2 + 0.5)).time)
    If Trim(SampleData(Int(MyPos / 2 + 0.5)).FER) <> "" Then
        Label8.Caption = Trim(SampleData(Int(MyPos / 2 + 0.5)).FER) & "%"
    Else
        Label8.Caption = ""
    End If
    For i = 1 To 8
        If i = 6 Then
           If Label2(5).Caption = "RxQual:" Then
              Label3(i).Caption = Trim(SampleData(Int(MyPos / 2 + 0.5)).FieldCol(i))
           ElseIf Label2(5).Caption = "BER:" Then
                If Trim(SampleData(Int(MyPos / 2 + 0.5)).FieldCol(i)) = "0" Then
                    Label3(i).Caption = "0.14%"
                ElseIf Trim(SampleData(Int(MyPos / 2 + 0.5)).FieldCol(i)) = "1" Then
                    Label3(i).Caption = "0.28%"
                ElseIf Trim(SampleData(Int(MyPos / 2 + 0.5)).FieldCol(i)) = "2" Then
                    Label3(i).Caption = "0.57%"
                ElseIf Trim(SampleData(Int(MyPos / 2 + 0.5)).FieldCol(i)) = "3" Then
                    Label3(i).Caption = "1.13%"
                ElseIf Trim(SampleData(Int(MyPos / 2 + 0.5)).FieldCol(i)) = "4" Then
                    Label3(i).Caption = "2.26%"
                ElseIf Trim(SampleData(Int(MyPos / 2 + 0.5)).FieldCol(i)) = "5" Then
                    Label3(i).Caption = "4.53%"
                ElseIf Trim(SampleData(Int(MyPos / 2 + 0.5)).FieldCol(i)) = "6" Then
                    Label3(i).Caption = "9.05%"
                ElseIf Trim(SampleData(Int(MyPos / 2 + 0.5)).FieldCol(i)) = "7" Then
                    Label3(i).Caption = "18.10%"
                End If
                
           End If
        Else
            Label3(i).Caption = Trim(SampleData(Int(MyPos / 2 + 0.5)).FieldCol(i))
        End If
    Next
    Label4(0).Caption = Trim(SampleData(Int(MyPos / 2 + 0.5)).lon)
    Label4(1).Caption = Trim(SampleData(Int(MyPos / 2 + 0.5)).lat)
    If Label1(0).Caption = "-50 dBm" And Trim(Label3(5).Caption) <> "" Then
       Label3(5).Caption = Format(Val(Label3(5).Caption) - 110)
    End If
    Line1(MyIndex).x1 = MyPos
    Line1(MyIndex).X2 = MyPos
    If Line1(MyIndex).Visible = False Then
       Line1(MyIndex).Visible = True
       If Line1((MyIndex + 1) Mod 2).Visible = True Then
          Line1((MyIndex + 1) Mod 2).Visible = False
       End If
    End If
    
End Sub

Private Sub MyHideLine()
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To 1
        If Line1(i).Visible = True Then
           Line1(i).Visible = False
        End If
    Next
    For i = 0 To 8
        Label3(i).Caption = ""
    Next
    Label4(0).Caption = ""
    Label4(1).Caption = ""
    Label6.Caption = ""
    Label8.Caption = ""
    
End Sub
