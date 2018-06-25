VERSION 5.00
Begin VB.Form Ncell_graph 
   BackColor       =   &H80000005&
   Caption         =   "Neighbouring Cells"
   ClientHeight    =   3855
   ClientLeft      =   7275
   ClientTop       =   1605
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ncellgra.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4590
   Begin VB.CommandButton Command1 
      Caption         =   "显示 Full"
      Height          =   320
      Left            =   1185
      TabIndex        =   37
      Top             =   3435
      Width           =   1080
   End
   Begin VB.CommandButton OK 
      Cancel          =   -1  'True
      Caption         =   "关闭"
      Height          =   320
      Left            =   2415
      TabIndex        =   26
      Top             =   3435
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Tx_Power"
      Height          =   180
      Index           =   23
      Left            =   2310
      TabIndex        =   60
      Top             =   2790
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "5"
      Height          =   180
      Index           =   9
      Left            =   3735
      TabIndex        =   59
      Top             =   3135
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "15"
      Height          =   180
      Index           =   8
      Left            =   3285
      TabIndex        =   58
      Top             =   3135
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "0"
      Height          =   180
      Index           =   7
      Left            =   3960
      TabIndex        =   57
      Top             =   3135
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "19"
      Height          =   180
      Index           =   6
      Left            =   3075
      TabIndex        =   56
      Top             =   3135
      Width           =   180
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H0000C0C0&
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   8
      Left            =   3165
      Top             =   2790
      Width           =   840
   End
   Begin VB.Line Line5 
      Index           =   2
      X1              =   3150
      X2              =   4005
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line6 
      Index           =   14
      X1              =   3327
      X2              =   3327
      Y1              =   3015
      Y2              =   3105
   End
   Begin VB.Line Line6 
      Index           =   9
      X1              =   3769
      X2              =   3769
      Y1              =   3015
      Y2              =   3105
   End
   Begin VB.Line Line6 
      Index           =   8
      X1              =   3990
      X2              =   3990
      Y1              =   3015
      Y2              =   3105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "C2"
      Height          =   180
      Index           =   4
      Left            =   2835
      TabIndex        =   55
      Top             =   270
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   20
      Left            =   2865
      TabIndex        =   54
      Top             =   525
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   19
      Left            =   2865
      TabIndex        =   53
      Top             =   780
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   18
      Left            =   2865
      TabIndex        =   52
      Top             =   1020
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   17
      Left            =   2865
      TabIndex        =   51
      Top             =   1260
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   16
      Left            =   2865
      TabIndex        =   50
      Top             =   1515
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   15
      Left            =   2865
      TabIndex        =   49
      Top             =   1770
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   14
      Left            =   2865
      TabIndex        =   48
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "C1"
      Height          =   180
      Index           =   5
      Left            =   2475
      TabIndex        =   47
      Top             =   270
      Width           =   180
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   13
      Left            =   2505
      TabIndex        =   46
      Top             =   525
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   12
      Left            =   2505
      TabIndex        =   45
      Top             =   780
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   11
      Left            =   2505
      TabIndex        =   44
      Top             =   1020
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   10
      Left            =   2505
      TabIndex        =   43
      Top             =   1260
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   9
      Left            =   2505
      TabIndex        =   42
      Top             =   1515
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   8
      Left            =   2505
      TabIndex        =   41
      Top             =   1770
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   7
      Left            =   2505
      TabIndex        =   40
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Cell"
      Height          =   180
      Index           =   3
      Left            =   720
      TabIndex        =   39
      Top             =   510
      Width           =   360
   End
   Begin VB.Line Line3 
      Index           =   21
      X1              =   4260
      X2              =   4260
      Y1              =   285
      Y2              =   360
   End
   Begin VB.Line Line3 
      Index           =   20
      X1              =   4050
      X2              =   4050
      Y1              =   285
      Y2              =   360
   End
   Begin VB.Line Line3 
      Index           =   19
      X1              =   3870
      X2              =   3870
      Y1              =   285
      Y2              =   360
   End
   Begin VB.Line Line3 
      Index           =   18
      X1              =   3510
      X2              =   3510
      Y1              =   285
      Y2              =   360
   End
   Begin VB.Line Line3 
      Index           =   17
      X1              =   3705
      X2              =   3705
      Y1              =   285
      Y2              =   360
   End
   Begin VB.Line Line3 
      Index           =   14
      X1              =   3330
      X2              =   3330
      Y1              =   285
      Y2              =   360
   End
   Begin VB.Line Line6 
      Index           =   7
      X1              =   4005
      X2              =   4005
      Y1              =   2610
      Y2              =   2700
   End
   Begin VB.Line Line6 
      Index           =   6
      X1              =   3885
      X2              =   3885
      Y1              =   2610
      Y2              =   2700
   End
   Begin VB.Line Line6 
      Index           =   5
      X1              =   3765
      X2              =   3765
      Y1              =   2610
      Y2              =   2700
   End
   Begin VB.Line Line6 
      Index           =   4
      X1              =   3645
      X2              =   3645
      Y1              =   2610
      Y2              =   2700
   End
   Begin VB.Line Line6 
      Index           =   3
      X1              =   3525
      X2              =   3525
      Y1              =   2610
      Y2              =   2700
   End
   Begin VB.Line Line6 
      Index           =   2
      X1              =   3405
      X2              =   3405
      Y1              =   2610
      Y2              =   2700
   End
   Begin VB.Line Line6 
      Index           =   1
      X1              =   3285
      X2              =   3285
      Y1              =   2610
      Y2              =   2700
   End
   Begin VB.Line Line6 
      Index           =   0
      X1              =   3150
      X2              =   3150
      Y1              =   3015
      Y2              =   3105
   End
   Begin VB.Line Line5 
      Index           =   1
      X1              =   3150
      X2              =   4140
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   3165
      X2              =   4020
      Y1              =   2685
      Y2              =   2685
   End
   Begin VB.Line Line3 
      Index           =   16
      X1              =   3150
      X2              =   3150
      Y1              =   285
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "3"
      Height          =   165
      Index           =   22
      Left            =   1860
      TabIndex        =   38
      Top             =   3030
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "30"
      Height          =   180
      Index           =   21
      Left            =   3630
      TabIndex        =   36
      Top             =   90
      Width           =   180
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   7
      Left            =   3165
      Top             =   2415
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   6
      Left            =   3165
      Top             =   2040
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   5
      Left            =   3165
      Top             =   1785
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   4
      Left            =   3165
      Top             =   1530
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   3
      Left            =   3165
      Top             =   1275
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   2
      Left            =   3165
      Top             =   1035
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BorderColor     =   &H00008000&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   1
      Left            =   3165
      Top             =   795
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   150
      Index           =   0
      Left            =   3165
      Top             =   510
      Width           =   300
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   6
      Left            =   2115
      TabIndex        =   35
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   6
      Left            =   1695
      TabIndex        =   34
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   6
      Left            =   1290
      TabIndex        =   33
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Ncell6"
      Height          =   180
      Index           =   20
      Left            =   540
      TabIndex        =   32
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Ncell5"
      Height          =   180
      Index           =   19
      Left            =   540
      TabIndex        =   31
      Top             =   1770
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Ncell4"
      Height          =   180
      Index           =   18
      Left            =   540
      TabIndex        =   30
      Top             =   1515
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Ncell3"
      Height          =   180
      Index           =   17
      Left            =   540
      TabIndex        =   29
      Top             =   1260
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Ncell2"
      Height          =   180
      Index           =   16
      Left            =   540
      TabIndex        =   28
      Top             =   1020
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Ncell1"
      Height          =   180
      Index           =   15
      Left            =   540
      TabIndex        =   27
      Top             =   765
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "0"
      Height          =   180
      Index           =   14
      Left            =   3015
      TabIndex        =   25
      Top             =   2610
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "7"
      Height          =   180
      Index           =   13
      Left            =   4050
      TabIndex        =   24
      Top             =   2625
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "0"
      Height          =   180
      Index           =   12
      Left            =   3120
      TabIndex        =   23
      Top             =   90
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "63"
      Height          =   180
      Index           =   11
      Left            =   4200
      TabIndex        =   22
      Top             =   90
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "RxQual_s"
      Height          =   180
      Index           =   10
      Left            =   2310
      TabIndex        =   21
      Top             =   2415
      Width           =   720
   End
   Begin VB.Line Line2 
      X1              =   3150
      X2              =   3150
      Y1              =   3015
      Y2              =   360
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3150
      X2              =   4260
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   5
      Left            =   2115
      TabIndex        =   20
      Top             =   1770
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   4
      Left            =   2115
      TabIndex        =   19
      Top             =   1515
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   3
      Left            =   2115
      TabIndex        =   18
      Top             =   1260
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   2
      Left            =   2115
      TabIndex        =   17
      Top             =   1020
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   1
      Left            =   2115
      TabIndex        =   16
      Top             =   780
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   0
      Left            =   2115
      TabIndex        =   15
      Top             =   510
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   5
      Left            =   1695
      TabIndex        =   14
      Top             =   1770
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   4
      Left            =   1695
      TabIndex        =   13
      Top             =   1515
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   3
      Left            =   1695
      TabIndex        =   12
      Top             =   1260
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   2
      Left            =   1695
      TabIndex        =   11
      Top             =   1020
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   1
      Left            =   1695
      TabIndex        =   10
      Top             =   765
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   0
      Left            =   1695
      TabIndex        =   9
      Top             =   510
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   5
      Left            =   1290
      TabIndex        =   8
      Top             =   1770
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   4
      Left            =   1290
      TabIndex        =   7
      Top             =   1515
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   3
      Left            =   1290
      TabIndex        =   6
      Top             =   1260
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   2
      Left            =   1290
      TabIndex        =   5
      Top             =   1005
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   1
      Left            =   1290
      TabIndex        =   4
      Top             =   765
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   0
      Left            =   1290
      TabIndex        =   3
      Top             =   510
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "RxLev"
      Height          =   180
      Index           =   2
      Left            =   1950
      TabIndex        =   2
      Top             =   270
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "BSIC"
      Height          =   180
      Index           =   1
      Left            =   1530
      TabIndex        =   1
      Top             =   270
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "BCCH"
      Height          =   180
      Index           =   0
      Left            =   1065
      TabIndex        =   0
      Top             =   270
      Width           =   360
   End
End
Attribute VB_Name = "Ncell_graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyFullFlag As Boolean
Dim CurrentRow As Integer
Dim MyRows As Variant
Dim MySelName As String
Dim SelectRxlev_s As Variant, SelectRxlev_f As Variant, SelectRxQual_s As Variant, SelectRxQual_f As Variant
Dim SelectTx_power As Variant

Private Sub Command1_Click()
    On Error Resume Next
    If MyFullFlag = True Then
       Label4(0).Caption = SelectRxlev_s
       If Label1(10).Caption = "[IDLE] RxQual_f" Then
           Label1(10).Caption = "[IDLE] RxQual_s"
       Else
           If Val(SelectRxQual_s) > 3 Then
               Shape1(7).BorderColor = &HFF&
               Shape1(7).FillColor = &HFF&
           Else
               Shape1(7).BorderColor = &HFF0000
               Shape1(7).FillColor = &HFF0000
           End If
           Shape1(7).Width = Val(SelectRxQual_s) * 120
           Label1(10).Caption = "RxQual_s"
       End If
       'Label1(3).Caption = "RxLev_s"
       Command1.Caption = "显示 Full"
       MyFullFlag = False
    Else
       Label4(0).Caption = SelectRxlev_f
       If Label1(10).Caption = "[IDLE] RxQual_s" Then
           Label1(10).Caption = "[IDLE] RxQual_f"
       Else
           If Val(SelectRxQual_f) > 3 Then
               Shape1(7).BorderColor = &HFF&
               Shape1(7).FillColor = &HFF&
           Else
               Shape1(7).BorderColor = &HFF0000
               Shape1(7).FillColor = &HFF0000
           End If
           Shape1(7).Width = Val(SelectRxQual_f) * 120
           Label1(10).Caption = "RxQual_f"
       End If
      ' Label1(3).Caption = "RxLev_f"
       Command1.Caption = "显示 Sub"
       MyFullFlag = True
    End If
    If Val(Label4(0).Caption) > 63 Then
       Shape1(0).Width = 63 * 18
    Else
       Shape1(0).Width = Val(Label4(0).Caption) * 18
    End If
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    
    On Error Resume Next
    If CurrentRow >= MyRows Then
       Exit Sub
    End If
    mapinfo.do "fetch next from " & SelTbl
    Label2(0).Caption = mapinfo.eval("selection.bcch_serv")
    Label3(0).Caption = mapinfo.eval("selection.bsic_serv")
    If MyFullFlag = True Then
       Label4(0).Caption = mapinfo.eval("selection.rxlev_f")
    Else
       Label4(0).Caption = mapinfo.eval("selection.rxlev_s")
    End If
    For i = 1 To 6
        Label2(i).Caption = mapinfo.eval("selection.bcch_n" & i)
        Label3(i).Caption = mapinfo.eval("selection.bsic_n" & i)
        Label4(i).Caption = mapinfo.eval("selection.rxlev_n" & i)
    Next
    If Val(Label4(0).Caption) > 63 Then
       Shape1(0).Height = 63 * 18
    Else
       Shape1(0).Height = Val(Label4(0).Caption) * 18
    End If
    Shape1(0).Top = 2205 - Shape1(0).Height
    For i = 1 To 6
        If Val(Label4(i).Caption) > 63 Then
           Shape1(i).Height = 63 * 18
        Else
           Shape1(i).Height = Val(Label4(i).Caption) * 18
        End If
        Shape1(i).Top = 2205 - Shape1(i).Height
    Next
    If MyFullFlag = True Then
       If Val(mapinfo.eval("selection.rxqual_f")) > 3 Then
          Shape1(7).BorderColor = &HFF&
          Shape1(7).FillColor = &HFF&
       Else
          Shape1(7).BorderColor = &HFF0000
          Shape1(7).FillColor = &HFF0000
       End If
       Shape1(7).Height = Val(mapinfo.eval("selection.rxqual_f")) * 120
    Else
       If Val(mapinfo.eval("selection.rxqual_s")) > 3 Then
          Shape1(7).BorderColor = &HFF&
          Shape1(7).FillColor = &HFF&
       Else
          Shape1(7).BorderColor = &HFF0000
          Shape1(7).FillColor = &HFF0000
       End If
       Shape1(7).Height = Val(mapinfo.eval("selection.rxqual_s")) * 120
    End If
    Shape1(7).Top = 2205 - Shape1(7).Height
    CurrentRow = CurrentRow + 1
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    On Error Resume Next
    If CurrentRow <= 1 Then
       Exit Sub
    End If
    mapinfo.do "fetch prev from " & SelTbl
    Label2(0).Caption = mapinfo.eval("selection.bcch_serv")
    Label3(0).Caption = mapinfo.eval("selection.bsic_serv")
    If MyFullFlag = True Then
       Label4(0).Caption = mapinfo.eval("selection.rxlev_f")
    Else
       Label4(0).Caption = mapinfo.eval("selection.rxlev_s")
    End If
    For i = 1 To 6
        Label2(i).Caption = mapinfo.eval("selection.bcch_n" & i)
        Label3(i).Caption = mapinfo.eval("selection.bsic_n" & i)
        Label4(i).Caption = mapinfo.eval("selection.rxlev_n" & i)
    Next
    If Val(Label4(0).Caption) > 63 Then
       Shape1(0).Height = 63 * 18
    Else
       Shape1(0).Height = Val(Label4(0).Caption) * 18
    End If
    Shape1(0).Top = 2205 - Shape1(0).Height
    For i = 1 To 6
        If Val(Label4(i).Caption) > 63 Then
           Shape1(i).Height = 63 * 18
        Else
           Shape1(i).Height = Val(Label4(i).Caption) * 18
        End If
        Shape1(i).Top = 2205 - Shape1(i).Height
    Next
    If MyFullFlag = True Then
       If Val(mapinfo.eval("selection.rxqual_f")) > 3 Then
          Shape1(7).BorderColor = &HFF&
          Shape1(7).FillColor = &HFF&
       Else
          Shape1(7).BorderColor = &HFF0000
          Shape1(7).FillColor = &HFF0000
       End If
       Shape1(7).Height = Val(mapinfo.eval("selection.rxqual_f")) * 120
    Else
       If Val(mapinfo.eval("selection.rxqual_s")) > 3 Then
          Shape1(7).BorderColor = &HFF&
          Shape1(7).FillColor = &HFF&
       Else
          Shape1(7).BorderColor = &HFF0000
          Shape1(7).FillColor = &HFF0000
       End If
       Shape1(7).Height = Val(mapinfo.eval("selection.rxqual_s")) * 120
    End If
    Shape1(7).Top = 2205 - Shape1(7).Height
    CurrentRow = CurrentRow - 1

End Sub

Public Sub Form_Load()
    Dim i As Integer
    Dim mySelRows As Variant
    Dim CellName As String
    Dim NcellLon As Variant, NcellLat As Variant
    Dim MyCival As String
    
    On Error Resume Next
    NcellWinFlag = True
    MySelName = mapinfo.eval("selectioninfo(2)")
    MyRows = mapinfo.eval("tableinfo(" & SelTbl & ",8)")
    mySelRows = mapinfo.eval("searchpoint(" & mapid & ",selection.lon,selection.lat)")
    mySelRows = mapinfo.eval("SearchInfo(1, 2)")
'    mapinfo.Do "Fetch Rec " & mySelRows & " FROM " & SelTbl
    MyFullFlag = False
    CurrentRow = mySelRows
    If UCase(mapinfo.eval("Columninfo( " & SelTbl & ",COL30, 1)")) = "C1" Then
       Label4(13).Caption = Trim(mapinfo.eval("selection.c1"))
       Label4(20).Caption = Trim(mapinfo.eval("selection.c2"))
       For i = 1 To 6
           Label4(13 - i).Caption = Trim(mapinfo.eval("selection.c1_n" & Format(i)))
           Label4(20 - i).Caption = Trim(mapinfo.eval("selection.c2_n" & Format(i)))
       Next
    End If
    MyCival = mapinfo.eval("selection.ci_serv")
    Label2(0).Caption = mapinfo.eval("selection.bcch_serv")
    Label3(0).Caption = mapinfo.eval("selection.bsic_serv")
    'Label4(0).Caption = mapinfo.eval("selection.rxlev_f")
    Label4(0).Caption = mapinfo.eval("selection.rxlev_s")
    SelectRxlev_s = mapinfo.eval("selection.rxlev_s")
    SelectRxlev_f = mapinfo.eval("selection.rxlev_f")
    SelectRxQual_s = mapinfo.eval("selection.rxqual_s")
    SelectRxQual_f = mapinfo.eval("selection.rxqual_f")
    SelectTx_power = mapinfo.eval("selection.tx_power")
    NcellLon = mapinfo.eval("selection.lon")
    NcellLat = mapinfo.eval("selection.lat")
    If Val(mapinfo.eval("selection.rxqual_s")) > 3 Then
       Shape1(7).BorderColor = &HFF&
       Shape1(7).FillColor = &HFF&
    Else
       Shape1(7).BorderColor = &HFF0000
       Shape1(7).FillColor = &HFF0000
    End If
    If Trim(mapinfo.eval("selection.rxqual_s")) = "" Or mapinfo.eval("selection.rxqual_s") = 9 Then
       Shape1(7).Width = 0
       Label1(10).Caption = "[IDLE] RxQual_s"
    Else
       Shape1(7).Width = Val(mapinfo.eval("selection.rxqual_s")) * 120
       Label1(10).Caption = "RxQual_s"
    End If
    For i = 1 To 6
        If Val(mapinfo.eval("selection.bcch_n" & i)) = 0 Then
           Label2(i).Caption = ""
           Label3(i).Caption = ""
           Label4(i).Caption = ""
           Label1(i + 14).Caption = ""
        Else
           Label2(i).Caption = mapinfo.eval("selection.bcch_n" & i)
           Label3(i).Caption = mapinfo.eval("selection.bsic_n" & i)
           Label4(i).Caption = mapinfo.eval("selection.rxlev_n" & i)
           If Val(Label3(i).Caption) = 99 Then
              Shape1(i).FillColor = &HFF&
              Shape1(i).BorderColor = &HFF&
           Else
              Shape1(i).FillColor = 32768
              Shape1(i).BorderColor = 32768
           End If
        End If
    Next
    If SelectTx_power = "" Then
       Shape1(8).Width = 0
       Label1(23).Caption = "[IDLE] Tx_Power"
    Else
       Shape1(8).Width = (840 / 19) * (19 - SelectTx_power)
       Label1(23).Caption = "Tx_Power"
    End If
    If Val(Label4(0).Caption) > 63 Then
       Shape1(0).Width = 63 * 18
    Else
       Shape1(0).Width = Val(Label4(0).Caption) * 18
    End If
    For i = 1 To 6
        If Val(Label4(i).Caption) > 63 Then
           Shape1(i).Width = 63 * 18
        Else
           Shape1(i).Width = Val(Label4(i).Caption) * 18
        End If
    Next
    'mapinfo.do "close table " & MySelName
    Call SearchCellName(0, 0, 0, 0, CellName, MyCival, "")
    If Trim(CellName) <> "" Then
       Label1(3).Caption = CellName
    Else
       Label1(3).Caption = "Cell"
    End If
    For i = 1 To 6
        'If Val(Label3(i).Caption) > 0 And Val(Label3(i).Caption) <> 99 Then
        If Val(Label3(i).Caption) <> 99 Then
           Call SearchCellName(Val(Label3(i).Caption), Val(Label2(i).Caption), NcellLon, NcellLat, CellName, "", Label1(3).Caption)
           If Trim(CellName) <> "" Then
              Label1(i + 14).Caption = CellName
           Else
              Label1(i + 14).Caption = "Ncell" & Format(i)
           End If
        Else
           Label1(i + 14).Caption = "Ncell" & Format(i)
        End If
    Next
'    mapinfo.do "close table mytemp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    NcellWinFlag = False
End Sub

Private Sub OK_Click()
    On Error Resume Next
    mapinfo.do "close table " & MySelName
    Unload Me
End Sub
