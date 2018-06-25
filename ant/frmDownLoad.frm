VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDownLoad 
   Caption         =   "智能导入"
   ClientHeight    =   6585
   ClientLeft      =   1695
   ClientTop       =   1800
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDownLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8700
   Begin VB.Frame Frame1 
      Caption         =   "小区库参数关连"
      Height          =   3825
      Left            =   225
      TabIndex        =   3
      Top             =   60
      Width           =   8235
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IntegralHeight  =   0   'False
         ItemData        =   "frmDownLoad.frx":000C
         Left            =   6495
         List            =   "frmDownLoad.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   3405
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "经纬度格式:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   35
         Left            =   5475
         TabIndex        =   76
         Top             =   3465
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   6510
         TabIndex        =   75
         Top             =   3120
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell16:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   34
         Left            =   5745
         TabIndex        =   74
         Top             =   3120
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   6510
         TabIndex        =   73
         Top             =   2820
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell15:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   33
         Left            =   5745
         TabIndex        =   72
         Top             =   2835
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   6510
         TabIndex        =   71
         Top             =   2535
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell14:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   32
         Left            =   5745
         TabIndex        =   70
         Top             =   2550
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   6510
         TabIndex        =   69
         Top             =   2250
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell13:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   31
         Left            =   5745
         TabIndex        =   68
         Top             =   2265
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   6510
         TabIndex        =   67
         Top             =   1965
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell12:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   30
         Left            =   5745
         TabIndex        =   66
         Top             =   1980
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   6510
         TabIndex        =   65
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell11:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   29
         Left            =   5745
         TabIndex        =   64
         Top             =   1695
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   6510
         TabIndex        =   63
         Top             =   1395
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell10:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   28
         Left            =   5745
         TabIndex        =   62
         Top             =   1410
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   6510
         TabIndex        =   61
         Top             =   1110
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell9:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   27
         Left            =   5850
         TabIndex        =   60
         Top             =   1125
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   6510
         TabIndex        =   59
         Top             =   825
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell8:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   26
         Left            =   5850
         TabIndex        =   58
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   6510
         TabIndex        =   57
         Top             =   540
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell7:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   25
         Left            =   5850
         TabIndex        =   56
         Top             =   555
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   6510
         TabIndex        =   55
         Top             =   255
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell6:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   24
         Left            =   5850
         TabIndex        =   54
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   3975
         TabIndex        =   53
         Top             =   3405
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell5:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   23
         Left            =   3315
         TabIndex        =   52
         Top             =   3420
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   3975
         TabIndex        =   51
         Top             =   3120
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell4:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   22
         Left            =   3315
         TabIndex        =   50
         Top             =   3135
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   3975
         TabIndex        =   49
         Top             =   2835
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell3:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   21
         Left            =   3315
         TabIndex        =   48
         Top             =   2850
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   3975
         TabIndex        =   47
         Top             =   2550
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell2:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   20
         Left            =   3315
         TabIndex        =   46
         Top             =   2565
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   3975
         TabIndex        =   45
         Top             =   2265
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ncell1:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   19
         Left            =   3315
         TabIndex        =   44
         Top             =   2280
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   3975
         TabIndex        =   43
         Top             =   1980
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Length:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   18
         Left            =   3270
         TabIndex        =   42
         Top             =   1995
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   3975
         TabIndex        =   41
         Top             =   1695
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "BaseType:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   17
         Left            =   3015
         TabIndex        =   40
         Top             =   1710
         Width           =   870
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   3975
         TabIndex        =   39
         Top             =   1410
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lat:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   16
         Left            =   3585
         TabIndex        =   38
         Top             =   1425
         Width           =   300
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   3975
         TabIndex        =   37
         Top             =   1125
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lon:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   15
         Left            =   3525
         TabIndex        =   36
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3975
         TabIndex        =   35
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   14
         Left            =   3420
         TabIndex        =   34
         Top             =   855
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3975
         TabIndex        =   33
         Top             =   555
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ant_Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   13
         Left            =   3090
         TabIndex        =   32
         Top             =   570
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3975
         TabIndex        =   31
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ant_Gain:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   12
         Left            =   3090
         TabIndex        =   30
         Top             =   285
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   1410
         TabIndex        =   29
         Top             =   3420
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Max_Tx_Ms:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   11
         Left            =   330
         TabIndex        =   28
         Top             =   3420
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   1410
         TabIndex        =   27
         Top             =   3120
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ant_Heigh:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   10
         Left            =   420
         TabIndex        =   26
         Top             =   3135
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   1410
         TabIndex        =   25
         Top             =   2835
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Max_Tx_Bts:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   9
         Left            =   300
         TabIndex        =   24
         Top             =   2850
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1410
         TabIndex        =   23
         Top             =   2550
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Downtilt:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   8
         Left            =   615
         TabIndex        =   22
         Top             =   2565
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1410
         TabIndex        =   21
         Top             =   2265
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Non_Bcch:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   7
         Left            =   420
         TabIndex        =   20
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1410
         TabIndex        =   19
         Top             =   1980
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LAC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   6
         Left            =   930
         TabIndex        =   18
         Top             =   1995
         Width           =   390
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1410
         TabIndex        =   17
         Top             =   1695
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bearing:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   5
         Left            =   630
         TabIndex        =   16
         Top             =   1710
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1410
         TabIndex        =   15
         Top             =   1410
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bsic:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   4
         Left            =   915
         TabIndex        =   14
         Top             =   1425
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1410
         TabIndex        =   13
         Top             =   1125
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Arfcn:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   3
         Left            =   870
         TabIndex        =   12
         Top             =   1140
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1410
         TabIndex        =   11
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CI:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   2
         Left            =   1095
         TabIndex        =   10
         Top             =   855
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1410
         TabIndex        =   9
         Top             =   555
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bs_no:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   1
         Left            =   735
         TabIndex        =   8
         Top             =   570
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1410
         TabIndex        =   5
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cell_name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   285
         Width           =   960
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      DragIcon        =   "frmDownLoad.frx":0010
      Height          =   320
      Left            =   3300
      TabIndex        =   2
      Top             =   6210
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      DragIcon        =   "frmDownLoad.frx":0162
      Height          =   320
      Left            =   4500
      TabIndex        =   1
      Top             =   6210
      Width           =   1080
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      DragIcon        =   "frmDownLoad.frx":02B4
      Height          =   1485
      Left            =   225
      TabIndex        =   0
      Top             =   3975
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   2619
      _Version        =   327680
      Rows            =   5
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      HighLight       =   2
      SelectionMode   =   2
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      MouseIcon       =   "frmDownLoad.frx":06F6
      OLEDropMode     =   1
   End
   Begin VB.Label Label4 
      Caption         =   "·如果表格中的中文出现乱码，请先退出，然后把要导入的文件格式存为Excel 5.0/95再重新导入。"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   5865
      Width           =   8025
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "·请用鼠标拖动各列并把它放到相应的小区库参数旁。"
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   5580
      Width           =   4320
   End
   Begin VB.Menu MnuPopMenu 
      Caption         =   "PopMenu"
      Visible         =   0   'False
      Begin VB.Menu MnuClear 
         Caption         =   "清空"
      End
      Begin VB.Menu MnuDefine 
         Caption         =   "定义"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmDownLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentLabel As Integer
Dim MyCombine(34) As String
Dim XLSRows As Long

Private Sub Command1_Click()
    Dim MyDatabase As Database
    Dim MyRecordset As Recordset
    Dim hFreefile As Integer
    Dim CellData As NewCell1800
    Dim i As Long
    Dim j As Integer, n As Integer
    Dim ImportFile As String
    Dim strFind As String
    Dim MyRecord As Record
    Dim MyNewCellName As String, MyNewBs_no As String, MyNewCi As String
    Dim MyLon As Variant, MyLat As Variant
    Dim strTmp1 As String, strTmp2 As String, strTmp3 As String
    Dim MyCellnameTmp As String
    
    Dim k As Integer, HH As Integer
    Dim CellRows As Long
    Dim MyCellColor As Long
    Dim AntennaL As Single
    Dim old_name As String, o_name As String
    Dim BASETYPE As String
    Dim BaseColor As Long
    Dim bcch(3), BSIC(3), ci(3), MyDir(3), other(4) As String
    Dim Lac As String, bs_name As String, bs_no As String
    Dim LeeFinds As Integer
    Dim row As Long, LacRow As Long
    
    
    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\gsm.dat"
    hFreefile = FreeFile
    Open Gsm_FileName For Binary As #hFreefile
    Get #hFreefile, 1, MyRecord
    Close #hFreefile
    
    Set MyDatabase = OpenDatabase(Gsm_Path & "\map", False, False, "FoxPro 2.5;")
    Set MyRecordset = MyDatabase.OpenRecordset("SELECT * " & "FROM cell ORDER BY bs_no ", dbOpenDynaset)
    mapinfo.do "fetch first from CellTemp"
    mapinfo.do "fetch next from CellTemp"
    For i = 1 To XLSRows - 1
        MyNewCellName = GetFieldValue(0)
        MyNewBs_no = GetFieldValue(1)
        MyNewCi = GetFieldValue(2)
        If MyRecord.exchange = 0 Then
           If MyNewBs_no = "$" Then
              If MyNewCellName = "$" Then
                 Exit For
              End If
              If MyNewCellName <> "" Then
                 MyRecordset.FindFirst "cell_name = """ & MyNewCellName & """"
              End If
           Else
              If MyNewBs_no <> "" Then
                 MyRecordset.FindFirst "bs_no = """ & MyNewBs_no & """"
              End If
           End If
        Else
           If MyNewCi = "$" Then
              If MyNewCellName = "$" Then
                 Exit For
              End If
              If MyNewCellName <> "" Then
                 MyRecordset.FindFirst "cell_name = """ & MyNewCellName & """"
              End If
           Else
              If MyNewCi <> "" Then
                 MyRecordset.FindFirst "ci = """ & MyNewCi & """"
              End If
           End If
        End If
        If MyRecordset.NoMatch Then
           If MyNewCellName <> "" Then
              If Asc(Right(MyNewCellName, 1)) >= 48 And Asc(Right(MyNewCellName, 1)) <= 57 Then
                 MyRecordset.FindFirst "cell_name = """ & MyNewCellName & """"
              Else
                 MyRecordset.AddNew
              End If
           End If
           If MyRecordset.NoMatch Then
              MyRecordset.AddNew
           Else
              MyRecordset.Edit
           End If
        Else
           MyRecordset.Edit
        End If
        For j = 0 To 34
            If GetFieldValue(j) <> "$" Then
               Select Case j
                   Case 0
                        MyCellnameTmp = GetFieldValue(j)
                        If LenB(MyCellnameTmp) > 20 Then
                            MyCellnameTmp = LeftB(MyCellnameTmp, 20)
                        End If
                        If Trim(MyCellnameTmp) = "" Then
                            strTmp1 = GetFieldValue(1)
                            If strTmp1 = "" Then
                                strTmp1 = GetFieldValue(2)
                                If strTmp1 = "" Then
                                    MyRecordset.Fields("cell_name").Value = "无"
                                Else
                                    MyRecordset.Fields("cell_name").Value = strTmp1
                                End If
                            Else
                                MyRecordset.Fields("cell_name").Value = strTmp1
                            End If
                        Else
                            MyRecordset.Fields("cell_name").Value = MyCellnameTmp
                        End If
                   Case 1
                        strTmp1 = GetFieldValue(j)
                        If Len(strTmp1) > 10 Then
                            strTmp1 = Left(strTmp1, 10)
                        End If
                        MyRecordset.Fields("bs_no").Value = strTmp1
                   Case 2
                        MyRecordset.Fields("ci").Value = GetFieldValue(j)
                   Case 3
                        MyRecordset.Fields("arfcn").Value = Val(GetFieldValue(j))
                   Case 4
                        MyRecordset.Fields("bsic").Value = Val(GetFieldValue(j))
                   Case 5
                        MyRecordset.Fields("bearing").Value = Val(Val(GetFieldValue(j)))
                   Case 6
                        MyRecordset.Fields("lac").Value = Val(GetFieldValue(j))
                   Case 7
                        strTmp1 = GetFieldValue(j)
                        strTmp1 = Trim(strTmp1)
                        If Len(strTmp1) > 0 Then
                            If Right(strTmp1, 1) = "," Then
                                strTmp1 = Left(strTmp1, Len(strTmp1) - 1)
                            ElseIf Left(strTmp1, 1) = "," Then
                                strTmp1 = Right(strTmp1, Len(strTmp1) - 1)
                            End If
                        End If
                        MyRecordset.Fields("non_bcch").Value = strTmp1
                   Case 8
                        MyRecordset.Fields("downtilt").Value = Val(GetFieldValue(j))
                   Case 9
                        MyRecordset.Fields("max_tx_ms").Value = GetFieldValue(j)
                   Case 10
                        MyRecordset.Fields("ant_heigh").Value = GetFieldValue(j)
                   Case 11
                        MyRecordset.Fields("max_tx_ms").Value = GetFieldValue(j)
                   Case 12
                        strTmp1 = GetFieldValue(j)
                        If Len(strTmp1) > 3 Then
                            strTmp1 = Left(strTmp1, 3)
                            If Right(strTmp1, 1) = "." Then
                                strTmp1 = Left(strTmp1, Len(strTmp1) - 1)
                            End If
                        End If
                        MyRecordset.Fields("ant_gain").Value = strTmp1
                   Case 13
                        strTmp1 = GetFieldValue(j)
                        If Len(strTmp1) > 15 Then
                            strTmp1 = Left(strTmp1, 15)
                        End If
                        MyRecordset.Fields("ant_type").Value = strTmp1
                   Case 14
                        strTmp1 = GetFieldValue(j)
                        If Len(strTmp1) > 8 Then
                            strTmp1 = Left(strTmp1, 8)
                        End If
                        MyRecordset.Fields("time").Value = strTmp1
                   Case 15
                        MyLon = GetFieldValue(j)
                   Case 16
                        MyLat = GetFieldValue(j)
                   Case 17
                        MyRecordset.Fields("basetype").Value = GetFieldValue(j)
                   Case 18
                        MyRecordset.Fields("length").Value = GetFieldValue(j)
                   Case Else
                        strTmp1 = GetFieldValue(j)
                        If Len(strTmp1) > 10 Then
                            strTmp1 = Left(strTmp1, 10)
                        End If
                        MyRecordset.Fields("ncell" & Format(j - 18)).Value = strTmp1
               End Select
            Else
               If j = 17 Then
                  If Val(GetFieldValue(3)) < 125 Then
                     MyRecordset.Fields("basetype").Value = "0"
                  Else
                     MyRecordset.Fields("basetype").Value = "3"
                  End If
               End If
            End If
        Next
        strTmp1 = ""
        strTmp2 = ""
        strTmp3 = ""
        Select Case Combo1.Text
            Case "度 分 秒", "度" & Chr(-24093) & "分" & "' " & "秒" & Chr(34)
                MyLon = Trim(MyLon)
                If Len(MyLon) = 0 Or Val(MyLon) = 0 Then
                   MyLon = 0
                   GoTo nonLon
                End If
                Do While Asc(Left(MyLon, 1)) > 57 Or Asc(Left(MyLon, 1)) < 48
                   MyLon = Right(MyLon, Len(MyLon) - 1)
                Loop
                If Len(MyLon) = 0 Then
                   MyLon = 0
                   GoTo nonLon
                End If
                Do While Asc(Left(MyLon, 1)) < 58 And Asc(Left(MyLon, 1)) > 47
                   strTmp1 = strTmp1 & Left(MyLon, 1)
                   MyLon = Right(MyLon, Len(MyLon) - 1)
                Loop
                Do While Asc(Left(MyLon, 1)) > 57 Or Asc(Left(MyLon, 1)) < 48
                   MyLon = Right(MyLon, Len(MyLon) - 1)
                Loop
                Do While Asc(Left(MyLon, 1)) < 58 And Asc(Left(MyLon, 1)) > 47
                   strTmp2 = strTmp2 & Left(MyLon, 1)
                   MyLon = Right(MyLon, Len(MyLon) - 1)
                Loop
                Do While Asc(Left(MyLon, 1)) > 57 Or Asc(Left(MyLon, 1)) < 48
                   MyLon = Right(MyLon, Len(MyLon) - 1)
                Loop
                Do While Asc(Left(MyLon, 1)) < 58 And Asc(Left(MyLon, 1)) > 47 Or Left(MyLon, 1) = "."
                   strTmp3 = strTmp3 & Left(MyLon, 1)
                   MyLon = Right(MyLon, Len(MyLon) - 1)
                   If Len(MyLon) = 0 Then
                      Exit Do
                   End If
                Loop
                MyLon = strTmp1 + Val(strTmp2) / 60 + Val(strTmp3) / 3600
nonLon:
                MyLat = Trim(MyLat)
                If Len(MyLat) = 0 Or Val(MyLat) = 0 Then
                   MyLat = 0
                   GoTo nonLat
                End If
                Do While Asc(Left(MyLat, 1)) > 57 Or Asc(Left(MyLat, 1)) < 48
                   MyLat = Right(MyLat, Len(MyLat) - 1)
                Loop
                If Len(MyLat) = 0 Then
                   MyLat = 0
                   GoTo nonLat
                End If
                strTmp1 = ""
                strTmp2 = ""
                strTmp3 = ""
                Do While Asc(Left(MyLat, 1)) < 58 And Asc(Left(MyLat, 1)) > 47
                   strTmp1 = strTmp1 & Left(MyLat, 1)
                   MyLat = Right(MyLat, Len(MyLat) - 1)
                Loop
                Do While Asc(Left(MyLat, 1)) > 57 Or Asc(Left(MyLat, 1)) < 48
                   MyLat = Right(MyLat, Len(MyLat) - 1)
                Loop
                Do While Asc(Left(MyLat, 1)) < 58 And Asc(Left(MyLat, 1)) > 47
                   strTmp2 = strTmp2 & Left(MyLat, 1)
                   MyLat = Right(MyLat, Len(MyLat) - 1)
                Loop
                Do While Asc(Left(MyLat, 1)) > 57 Or Asc(Left(MyLat, 1)) < 48
                   MyLat = Right(MyLat, Len(MyLat) - 1)
                Loop
                Do While Asc(Left(MyLat, 1)) < 58 And Asc(Left(MyLat, 1)) > 47 Or Left(MyLon, 1) = "."
                   strTmp3 = strTmp3 & Left(MyLat, 1)
                   MyLat = Right(MyLat, Len(MyLat) - 1)
                   If Len(MyLat) = 0 Then
                      Exit Do
                   End If
                Loop
                MyLat = strTmp1 + Val(strTmp2) / 60 + Val(strTmp3) / 3600
            Case "十进制格式"
        End Select
nonLat:
        If Val(MyLon) = 0 Then
            MyLon = 0
        End If
        If Val(MyLat) = 0 Then
            MyLat = 0
        End If
        MyRecordset.Fields("lon").Value = MyLon
        MyRecordset.Fields("lat").Value = MyLat
        MyRecordset.Update
        mapinfo.do "fetch next from celltemp"
    Next
    MyRecordset.Close
    MyDatabase.Close
    
        mapinfo.do "Register Table " + Chr(34) + Gsm_Path & "\map\cell.dbf" + Chr(34) + "Type ""DBF"" Into " + Chr(34) + Gsm_Path & "\map\cell.tab" + Chr(34)
        mapinfo.do "open table " + Chr(34) + Gsm_Path & "\map\cell.tab" + Chr(34)
        CellRows = mapinfo.eval("tableinfo(cell,8)")
        mapinfo.do "Create Map For cell CoordSys Earth Projection 1, 0"
        For i = 1 To CellRows
            If Val(mapinfo.eval("cell.length")) = 0 Then
               AntennaL = 0.002
            Else
               AntennaL = Val(mapinfo.eval("cell.length")) / 100000
            End If
'            mapinfo.do " x1 = cell.Lon + " & Format(AntennaL) & " * Sin(cell.bearing * 0.01745329252)" '  DEG_2_RAD)"
'            mapinfo.do " y1 = cell.Lat + " & Format(AntennaL) & " * Cos(cell.bearing * 0.01745329252)"  ' DEG_2_RAD)"
'            MyCellColor = MyCellRndColor(Val(mapinfo.eval("cell.arfcn")))
            If Val(mapinfo.eval("cell.basetype")) = 0 Then
               mapinfo.do " x1 = cell.Lon + " & Format(AntennaL) & " * Sin(cell.bearing * 0.01745329252)" '  DEG_2_RAD)"
               mapinfo.do " y1 = cell.Lat + " & Format(AntennaL) & " * Cos(cell.bearing * 0.01745329252)"  ' DEG_2_RAD)"
               If mapinfo.eval("cell.arfcn") > 124 Then
                  MyCellColor = MyCellRndColor(Val(mapinfo.eval("cell.arfcn")) Mod 124)
               Else
                  MyCellColor = MyCellRndColor(Val(mapinfo.eval("cell.arfcn")))
               End If
               mapinfo.do "Set Style Pen MakePen(1,60," & Format(MyCellColor) & ")"
               mapinfo.do "update cell  set Obj= CreateLine(x1,y1,cell.lon, cell.Lat)  where rowid=" & i
               If AntennaL = 0.002 Then
                  mapinfo.do "update cell  set LENGTH= ""200""  where rowid=" & i
               End If
            ElseIf Val(mapinfo.eval("cell.basetype")) = 3 Then
               mapinfo.do " x1 = cell.Lon + " & Format(AntennaL / 1.5) & " * Sin(cell.bearing * 0.01745329252)" '  DEG_2_RAD)"
               mapinfo.do " y1 = cell.Lat + " & Format(AntennaL / 1.5) & " * Cos(cell.bearing * 0.01745329252)" ' DEG_2_RAD)"
               If mapinfo.eval("cell.arfcn") > 124 Then
                  MyCellColor = MyCellRndColor(Val(mapinfo.eval("cell.arfcn")) Mod 124)
               Else
                  MyCellColor = MyCellRndColor(Val(mapinfo.eval("cell.arfcn")))
               End If
               mapinfo.do "Set Style Pen MakePen(1,60,0)"
               mapinfo.do "update cell  set Obj= CreateLine(x1,y1,cell.lon, cell.Lat)  where rowid=" & i
               If AntennaL = 0.002 Then
                  mapinfo.do "update cell  set LENGTH= ""200""  where rowid=" & i
               End If
            Else
               mapinfo.do "set style symbol MakeFontSymbol(59,16711680,12,""MapInfo Weather"",256,-Cell.bearing)"
               mapinfo.do "update cell set Obj= CreatePoint(cell.Lon,cell.Lat ) where rowid=" & i
            End If
            mapinfo.do "fetch next from cell"
        Next
        mapinfo.do "commit table cell"
    
    If Dir(Gsm_Path & "\map\base.tab", 0) = "" Then
        hDbfFile = FreeFile
        Open Gsm_Path & "\map\base.dbf" For Binary As #hDbfFile
        MakeBase1800File
        Close #hDbfFile
        mapinfo.do "Register Table " + Chr(34) + Gsm_Path & "\map\base.dbf" + Chr(34) + "Type ""DBF"" Into " + Chr(34) + Gsm_Path & "\map\base.tab" + Chr(34)
    End If
    mapinfo.do "open table " & Chr(34) + Gsm_Path & "\map\base.tab" + Chr(34)
    
    k = 1
    j = 1
    old_name = " "
    mapinfo.do "fetch first from cell"
    old_name = Trim(mapinfo.eval("cell.cell_name"))
    Do While mapinfo.eval("EOT(cell)") <> "T"
       If old_name <> "" Then Exit Do
       mapinfo.do "fetch next from cell"
       old_name = Trim(mapinfo.eval("cell.cell_name"))
    Loop
    If old_name = "" Then
       mapinfo.do "fetch first from base"
       mapinfo.do "delete from base"
       mapinfo.do "commit table base"
       mapinfo.do "pack table base Graphic Data Data Interactive  "
       GoTo no_cell
    End If
    Call getname(old_name)
    mapinfo.do "fetch First from base"
    
    mapinfo.do "delete from base"
    mapinfo.do "commit table Base"
    mapinfo.do "pack table base Graphic Data Data Interactive  "
    For i = 1 To 3
        ci(i) = " "
        MyDir(i) = "0"
        BSIC(i) = "0"
        bcch(i) = "0"
    Next i
   
   o_name = old_name
   While mapinfo.eval("EOT(cell)") <> "T"
         If old_name = o_name Then
cc:         bcch(k) = mapinfo.eval("cell.arfcn")
            ci(k) = mapinfo.eval("cell.ci")
            MyDir(k) = mapinfo.eval("cell.bearing")
            BSIC(k) = mapinfo.eval("cell.bsic")
            Lac = mapinfo.eval("cell.lac")
            bs_name = mapinfo.eval("cell.cell_name")
            bs_no = mapinfo.eval("cell.bs_no")
            BASETYPE = mapinfo.eval("cell.basetype")
            If BASETYPE = "" Then
               BASETYPE = "0"
            End If
            HH = 0
            For i = 1 To 4
                HH = i + 15
                other(i) = mapinfo.eval("cell.col" & HH)
            Next i
            mapinfo.do "x1 = cell.lon"
            mapinfo.do "y1 = cell.lat"
            k = k + 1
         Else
            LeeFinds = InStr(bs_no, Chr(0))
            If LeeFinds > 0 Then
               bs_no = Trim(Left(bs_no, LeeFinds - 1))
            End If
            Msg = "insert into  base  (col1,col2,col3,col4,col5,col6,col7,col8,col9,col10,col11,col12,col13,col14,col15,col16,col17,col18,col19) values ( "
            Msg = Msg + Chr(34) + old_name + Chr(34) + ","
            Msg = Msg + Chr(34) + bs_no + Chr(34) + ","
            Msg = Msg + bcch(1) + ","
            Msg = Msg + bcch(2) + ","
            Msg = Msg + bcch(3) + ","
            Msg = Msg + Chr(34) + ci(1) + Chr(34) + ","
            Msg = Msg + Chr(34) + ci(2) + Chr(34) + ","
            Msg = Msg + Chr(34) + ci(3) + Chr(34) + ","
            Msg = Msg + BSIC(1) + ","
            Msg = Msg + BSIC(2) + ","
            Msg = Msg + BSIC(3) + ","
            Msg = Msg + MyDir(1) + ","
            Msg = Msg + MyDir(2) + ","
            Msg = Msg + MyDir(3) + ","
            Msg = Msg + Chr(34) + Lac + Chr(34) + ","
            'msg = msg + Chr(34) + other(1) + Chr(34) + ","
            Msg = Msg + Chr(34) + " " + Chr(34) + ","
            Msg = Msg + Chr(34) + BASETYPE + Chr(34) + ","
            Msg = Msg + Chr(34) + " " + Chr(34) + ","
            Msg = Msg + Chr(34) + " " + Chr(34) + ")"
            mapinfo.do Msg
            Msg = "UPDATE  base  set lon = x1,lat = y1 where  rowid = " & j
            mapinfo.do Msg
            For i = 1 To 3
                ci(i) = " "
                MyDir(i) = "0"
                BSIC(i) = "0"
                bcch(i) = "0"
            Next i
            k = 1
            j = j + 1
            GoTo cc
         End If
         old_name = o_name
         Do While mapinfo.eval("EOT(cell)") <> "T"
            mapinfo.do "fetch next from cell"
            o_name = Trim(mapinfo.eval("cell.cell_name"))
            If o_name <> "" Then Exit Do
         Loop
         If o_name = "" Then GoTo exit_do
'         o_name = Mid(o_name, 1, 4)
         Call getname(o_name)

    Wend
exit_do:
           Msg = "insert into  base  (col1,col2,col3,col4,col5,col6,col7,col8,col9,col10,col11,col12,col13,col14,col15,col16,col17,col18,col19) values ( "
           Msg = Msg + Chr(34) + old_name + Chr(34) + ","
           Msg = Msg + Chr(34) + bs_no + Chr(34) + ","
           Msg = Msg + bcch(1) + ","
           Msg = Msg + bcch(2) + ","
           Msg = Msg + bcch(3) + ","
           Msg = Msg + Chr(34) + ci(1) + Chr(34) + ","
           Msg = Msg + Chr(34) + ci(2) + Chr(34) + ","
           Msg = Msg + Chr(34) + ci(3) + Chr(34) + ","
           Msg = Msg + BSIC(1) + ","
           Msg = Msg + BSIC(2) + ","
           Msg = Msg + BSIC(3) + ","
           Msg = Msg + MyDir(1) + ","
           Msg = Msg + MyDir(2) + ","
           Msg = Msg + MyDir(3) + ","
           Msg = Msg + Chr(34) + Lac + Chr(34) + ","
           Msg = Msg + Chr(34) + "" + Chr(34) + ","
           Msg = Msg + Chr(34) + BASETYPE + Chr(34) + ","
           Msg = Msg + Chr(34) + "" + Chr(34) + ","
           Msg = Msg + Chr(34) + "" + Chr(34) + ")"
           mapinfo.do Msg
           Msg = "UPDATE base set lon= x1,lat= y1 WHERE ROWID=" & j
           mapinfo.do Msg
           
    mapinfo.do "commit table Base"
    mapinfo.do "DROP MAP Base"
    mapinfo.do "Create Map For Base CoordSys Earth Projection 1, 0"

    i = 0
    'If Is_New = False Then
    '   Gsm_FileName = Gsm_Path + "\base_add.dbf"
    '   Gsm_File2 = Gsm_Path + "\map\base_add.dbf"
    '   Kill Gsm_File2
    '   FileCopy Gsm_FileName, Gsm_File2
    '   Gsm_FileName = Gsm_Path + "\map\base_add.tab"
    '   Kill Gsm_FileName
    '   mapinfo.Do "Register Table " + Chr(34) + Gsm_File2 + Chr(34) + "Type ""DBF"" Into " + Chr(34) + Gsm_FileName + Chr(34)
    'End If
    'mapinfo.Do "open table " + Chr(34) + Gsm_Path + "\map\base_add.tab" + Chr(34)
    row = Val(mapinfo.eval("TABLEINFO(Base, 8)"))
    mapinfo.do "SELECT lac FROM base where lac>0 group by lac order by lac desc into mytemp"
    LacRow = Val(mapinfo.eval("TABLEINFO(mytemp, 8)"))
    mapinfo.do "fetch first from Base"
    i = 1
    While i <= row
          mapinfo.do "fetch first from mytemp"
          If mapinfo.eval("base.lac") = 0 Then
             BaseColor = 0
          Else
             For k = 1 To LacRow
                 If mapinfo.eval("base.lac") = mapinfo.eval("mytemp.lac") Then
                    Exit For
                 End If
                 mapinfo.do "fetch next from mytemp"
             Next
             BaseColor = MyLacColor(k - 1)
          End If
'          msg = "base.bsic_1"
'          j = Val(mapinfo.eval(msg))
'          j = j * 12345678 + j * 876543
          'msg = "Set Style Symbol MakeFontSymbol(168," & j & ",12,""Symbol"",0,0)"
          'mapinfo.do "Set Style Symbol MakeFontSymbol(39," & j & ",12,""Wingdings 2"",256,0)"
         
          mapinfo.do "set style symbol MakeFontSymbol(39," & Format(BaseColor) & ",8,""MapInfo Cartographic"",0,0)"
          mapinfo.do "update Base  set Obj= CreatePoint(Lon,Lat ) where rowid=" & i
          old_name = mapinfo.eval("base.bs_name")
          Call getname(old_name)
          mapinfo.do "x1 = base.lon"
          mapinfo.do "y1 = base.lat"
             
          'msg = "insert into base_add (bs_name,address,lon,lat) values (" + Chr(34) + old_name + Chr(34) + "," + Chr(34) + Base_Address(i) + Chr(34) + ",x1,y1)"
          'mapinfo.Do msg
          'mapinfo.Do "fetch next from Base_add"
          mapinfo.do "fetch next from Base"
          i = i + 1
    Wend
        
'*********************************************************************************
no_cell:
    mapinfo.do "commit table base"
    mapinfo.do "close table base"
    
    mapinfo.do "close table cell"
    
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim XLSCols As Integer
    Dim i As Integer, j As Integer
    Dim MyFormatString As String, MyStringTemp As String
    
    On Error Resume Next
    mapinfo.do "Register Table " + Chr(34) + Gsm_FileName + Chr(34) + " TYPE XLS Into " + Chr(34) + Gsm_Path + "\CellTemp.tab" + Chr(34)
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\CellTemp.tab" + Chr(34)
    If Err Then
       'MsgBox "无法打开文件 " & Gsm_FileName & "或文件格式错误", 64, "提示"
       MsgBox "无法打开文件 " & Gsm_FileName & "或文件格式错误，" & Chr(10) & "请确定该文件是Excel 5.0/95格式并且只有一个工作表再做导入。", 64, "提示"
       Unload Me
       Exit Sub
    End If
    XLSRows = mapinfo.eval("tableinfo(CellTemp,8)")
    XLSCols = mapinfo.eval("tableinfo(CellTemp,4)")
    'if XLSRows<6 ...
    MSFlexGrid1.Cols = XLSCols
    For i = 1 To XLSCols
        MyStringTemp = mapinfo.eval("Columninfo( CellTemp,COL" & Format(i) & ", 1)")
        MyFormatString = MyFormatString & "^" & MyStringTemp & "|"
    Next
    MyFormatString = Left(MyFormatString, Len(MyFormatString) - 1)
    MSFlexGrid1.FormatString = MyFormatString
    For i = 0 To XLSCols - 1
        MSFlexGrid1.ColWidth(i) = 1000
    Next
    For i = 1 To 4
        MSFlexGrid1.row = i
        For j = 0 To XLSCols - 1
            MSFlexGrid1.col = j
            MSFlexGrid1.Text = mapinfo.eval("CellTemp.col" & Format(j + 1))
        Next
        mapinfo.do "fetch next from CellTemp"
    Next
    Combo1.AddItem "度 分 秒"
    Combo1.AddItem "度" & Chr(-24093) & "分" & "' " & "秒" & Chr(34)
    Combo1.AddItem "十进制格式"
    Combo1.Text = "度 分 秒"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mapinfo.do "close table CellTemp"
    Kill Gsm_Path + "\CellTemp.*"
End Sub

Private Sub Label2_Change(Index As Integer)
    On Error Resume Next
    If Len(Label2(Index).Caption) > 5 Then
       If Label2(Index).Width <> 645 * 2 Then
          Label2(Index).Width = 645 * 2
       End If
    ElseIf Label2(Index).Width > 645 Then
       Label2(Index).Width = 645
    End If
End Sub

Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim i As Integer
    Dim MyValue As String
    
    On Error Resume Next
    If Source.Name = "MSFlexGrid1" Then
       Source.col = Source.Tag
       Source.row = 0
       If Label2(Index).Caption <> "" Then
          If Index = 7 Then
             Label2(Index).Caption = Trim(Label2(Index).Caption) & "," & Source.Text
             MyCombine(Index) = MyCombine(Index) & " & "","" & " & "CellTemp.col" & Format(Source.Tag + 1)
          Else
             Label2(Index).Caption = Trim(Label2(Index).Caption) & " " & "+" & " " & Source.Text
             MyCombine(Index) = MyCombine(Index) & " & " & "CellTemp.col" & Format(Source.Tag + 1)
          End If
       Else
          For i = 0 To 34
              If Source.Text = Label2(i).Caption Then
                 MyValue = InputBox("请输入分界符：", "智能导入", "/")
                 If MyValue = "" Then
                    If Label2(Index).Caption <> "" Then
                        If Index = 7 Then
                           Label2(Index).Caption = Trim(Label2(Index).Caption) & "," & Source.Text
                           MyCombine(Index) = MyCombine(Index) & " & "","" & " & "CellTemp.col" & Format(Source.Tag + 1)
                        Else
                           Label2(Index).Caption = Trim(Label2(Index).Caption) & " " & "+" & " " & Source.Text
                           MyCombine(Index) = MyCombine(Index) & " & " & "CellTemp.col" & Format(Source.Tag + 1)
                        End If
                    Else
                        Label2(Index).Caption = Source.Text
                        MyCombine(Index) = "CellTemp.col" & Format(Source.Tag + 1)
                    End If
                 Else
                    Label2(Index).Caption = "Right(" & Source.Text & " , "" " & MyValue & " "" )"
                    Label2(i).Caption = "Left( " & Source.Text & " , "" " & MyValue & " "" )"
                    MyCombine(Index) = "Right$(CellTemp.col" & Format(Source.Tag + 1) & ",Len(CellTemp.col" & Format(Source.Tag + 1) & ")-Instr(1,CellTemp.col" & Format(Source.Tag + 1) & ",""" & MyValue & """))"
                    MyCombine(i) = "Left$(CellTemp.col" & Format(Source.Tag + 1) & ",Instr(1,CellTemp.col" & Format(Source.Tag + 1) & ",""" & MyValue & """)-1)"
                 End If
                 Exit For
              End If
          Next
          If Label2(Index).Caption = "" Then
             Label2(Index).Caption = Source.Text
             MyCombine(Index) = "CellTemp.col" & Format(Source.Tag + 1)
          End If
       End If
    End If
End Sub

Private Sub Label2_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

    On Error Resume Next
    If State = vbEnter Then
       Source.DragIcon = Command1.DragIcon
    ElseIf State = vbLeave Then
        Source.DragIcon = Command2.DragIcon
    End If

End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
       If Label2(Index).Caption = "" Then
          MnuClear.Enabled = False
          'MnuDefine.Enabled = False
       Else
          MnuClear.Enabled = True
       End If
       CurrentLabel = Index
       PopupMenu MnuPopMenu
    End If

End Sub

Private Sub MnuClear_Click()
    On Error Resume Next
    Label2(CurrentLabel).Caption = ""
    MyCombine(CurrentLabel) = ""
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    MSFlexGrid1.Tag = ""
    If MSFlexGrid1.MouseRow <> 0 Then Exit Sub
    MSFlexGrid1.Tag = str(MSFlexGrid1.MouseCol)
    MSFlexGrid1.DRAG 1
End Sub

Function GetFieldValue(MyIndex As Integer)
    Dim j As Integer
    Dim MyTemp As String
    
    On Error Resume Next
    If MyCombine(MyIndex) <> "" Then
       If j = 7 Then
          MyTemp = Trim(mapinfo.eval(MyCombine(MyIndex)))
          Do While Right(MyTemp, 1) = ","
             MyTemp = Left(MyTemp, Len(MyTemp) - 1)
          Loop
          GetFieldValue = MyTemp
       'ElseIf j = 0 Then
       '   Mytemp = Trim(mapinfo.eval(MyCombine(MyIndex)))
       '   GetFieldValue = Left(Mytemp, 21)
       ElseIf MyIndex = 17 Then
          MyTemp = Trim(mapinfo.eval(MyCombine(MyIndex)))
          GetFieldValue = Format(Val(MyTemp))
       Else
          GetFieldValue = Trim(mapinfo.eval(MyCombine(MyIndex)))
       End If
    Else
       GetFieldValue = "$"
    End If
End Function

Sub getname(MyName)
    Dim mychar As String
    Dim mycode As Integer, finds As Integer
    
    On Error Resume Next
    finds = InStr(MyName, Chr(0))
    If finds > 0 Then
       MyName = Left(MyName, finds - 1)
    End If
    MyName = Trim(MyName)
    If Len(MyName) > 0 Then
       mychar = Right(MyName, 1)
       mycode = Asc(mychar)
       'If mycode >= 65 And mycode <= 90 Or mycode >= 97 And mycode <= 122 Or mycode >= 48 And mycode <= 57 Then
       If mycode >= 48 And mycode <= 57 Then
          MyName = Left(MyName, Len(MyName) - 1)
          MyName = Trim(MyName)
       End If
    End If
End Sub

