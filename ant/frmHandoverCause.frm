VERSION 5.00
Begin VB.Form frmHandoverCause 
   Caption         =   "切换发生原因"
   ClientHeight    =   4440
   ClientLeft      =   1245
   ClientTop       =   3630
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHandoverCause.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "分析事件选择"
      Height          =   930
      Left            =   150
      TabIndex        =   43
      Top             =   240
      Width           =   7095
      Begin VB.OptionButton Option1 
         Caption         =   "全部切换事件"
         Height          =   225
         Index           =   1
         Left            =   4080
         TabIndex        =   45
         Top             =   420
         Width           =   1485
      End
      Begin VB.OptionButton Option1 
         Caption         =   "切换失败事件"
         Height          =   225
         Index           =   0
         Left            =   1185
         TabIndex        =   44
         Top             =   435
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      DragIcon        =   "frmHandoverCause.frx":000C
      Height          =   320
      Left            =   2490
      TabIndex        =   2
      Top             =   3975
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      DragIcon        =   "frmHandoverCause.frx":015E
      Height          =   320
      Left            =   3735
      TabIndex        =   1
      Top             =   3975
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "导致切换判定条件"
      Height          =   2415
      Left            =   150
      TabIndex        =   0
      Top             =   1305
      Width           =   7095
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   6540
         TabIndex        =   42
         Text            =   "1"
         Top             =   1680
         Width           =   315
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   5415
         TabIndex        =   40
         Text            =   "5"
         Top             =   1665
         Width           =   315
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   2835
         TabIndex        =   37
         Text            =   "90"
         Top             =   1680
         Width           =   435
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   1650
         MaxLength       =   1
         TabIndex        =   35
         Text            =   "4"
         Top             =   1695
         Width           =   330
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   6540
         TabIndex        =   33
         Text            =   "1"
         Top             =   1305
         Width           =   315
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   5415
         TabIndex        =   31
         Text            =   "5"
         Top             =   1290
         Width           =   315
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   2835
         TabIndex        =   28
         Text            =   "70"
         Top             =   1305
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   1650
         MaxLength       =   1
         TabIndex        =   26
         Text            =   "4"
         Top             =   1320
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   6540
         TabIndex        =   24
         Text            =   "1"
         Top             =   930
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   5415
         TabIndex        =   22
         Text            =   "5"
         Top             =   915
         Width           =   315
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   2835
         TabIndex        =   19
         Text            =   "104"
         Top             =   930
         Width           =   435
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1650
         MaxLength       =   1
         TabIndex        =   17
         Text            =   "4"
         Top             =   945
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6540
         TabIndex        =   15
         Text            =   "1"
         Top             =   540
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5415
         TabIndex        =   13
         Text            =   "5"
         Top             =   525
         Width           =   315
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2835
         TabIndex        =   10
         Text            =   "90"
         Top             =   540
         Width           =   435
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1650
         MaxLength       =   1
         TabIndex        =   8
         Text            =   "5"
         Top             =   555
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(DCS)≤"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   23
         Left            =   5820
         TabIndex        =   41
         Top             =   1740
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TxPower: (GSM)≤"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   22
         Left            =   3900
         TabIndex        =   39
         Top             =   1725
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "-dBm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   21
         Left            =   3315
         TabIndex        =   38
         Top             =   1710
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RxLev ＞"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   20
         Left            =   2130
         TabIndex        =   36
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RxQual ＜"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   19
         Left            =   825
         TabIndex        =   34
         Top             =   1755
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(DCS)≤"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   18
         Left            =   5820
         TabIndex        =   32
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TxPower: (GSM)≤"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   17
         Left            =   3900
         TabIndex        =   30
         Top             =   1335
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "-dBm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   3315
         TabIndex        =   29
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RxLev ＞"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   2130
         TabIndex        =   27
         Top             =   1350
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RxQual ＜"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   825
         TabIndex        =   25
         Top             =   1365
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(DCS)≤"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   5820
         TabIndex        =   23
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TxPower: (GSM)≤"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   3900
         TabIndex        =   21
         Top             =   975
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "-dBm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   3315
         TabIndex        =   20
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RxLev ＜"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   2130
         TabIndex        =   18
         Top             =   990
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RxQual ＜"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   825
         TabIndex        =   16
         Top             =   1005
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(DCS)≤"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   5820
         TabIndex        =   14
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TxPower: (GSM)≤"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   3900
         TabIndex        =   12
         Top             =   585
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "-dBm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   3315
         TabIndex        =   11
         Top             =   570
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RxLev ＞"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   2130
         TabIndex        =   9
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RxQual ≥"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   825
         TabIndex        =   7
         Top             =   615
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "拥塞："
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   6
         Top             =   1785
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "干扰："
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1395
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "场强："
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   4
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "误码："
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   3
         Top             =   645
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmHandoverCause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim MyRows As Integer, i As Integer
    Dim MyTxPwr1 As String, MyTxPwr2 As String
    Dim MyNumTable As Integer
    
    On Error Resume Next
    Me.Hide
    MyNumTable = mapinfo.eval("NumTables()")
    For i = 1 To MyNumTable
        If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "SELHANDOVER" Then
            mapinfo.do "close table SelHandOver"
            Exit For
        End If
    Next
    If Option1(0).Value Then
        mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= ""HANDOVER FAILURE"" into Mytemp"
    Else
        mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= ""HANDOVER COMPLETE"" or ucase$(MESSAGE)= ""HANDOVER COMMAND"" OR ucase$(MESSAGE)= ""HANDOVER FAILURE"" into Mytemp"
    End If
    MyRows = mapinfo.eval("tableinfo(mytemp,8)")
    If MyRows = 0 Then
       If Option1(0).Value Then
          MsgBox "该路段不存在切换失败", 64, "提示"
       Else
          MsgBox "该路段不存在切换事件", 64, "提示"
       End If
       mapinfo.do "close table mytemp"
       Unload Me
       Exit Sub
    End If
    mapinfo.do "commit table Mytemp as " + Chr(34) + Gsm_Path + "\User\SelHandOver.tab" + Chr(34)
    mapinfo.do "close table mytemp"
    mapinfo.do "open table " & Chr(34) + Gsm_Path + "\User\SelHandOver.tab" + Chr(34)
    MyRows = mapinfo.eval("tableinfo(SelHandOver,8)")
    mapinfo.do "fetch first from SelHandOver"
    For i = 1 To MyRows
        If mapinfo.eval("SelHandOver.bcch_serv") < 124 Then
           MyTxPwr1 = Text1(10).Text
           MyTxPwr2 = Text1(14).Text
        Else
           MyTxPwr1 = Text1(11).Text
           MyTxPwr2 = Text1(15).Text
        End If
        If Val(mapinfo.eval("SelHandOver.rxqual_s")) >= Val(Text1(0)) And mapinfo.eval("SelHandOver.rxlev_s") >= 110 - Val(Text1(1)) Then
           'mapinfo.do "Update SelHandOver Set Mark = ""WuMa"" Where RowID = " & Format(i)
           mapinfo.do "Update SelHandOver Set Mark = ""误码"" Where RowID = " & Format(i)
        ElseIf Val(mapinfo.eval("SelHandOver.rxqual_s")) < Val(Text1(4)) And mapinfo.eval("SelHandOver.rxlev_s") < 110 - Val(Text1(5)) Then
           'mapinfo.do "Update SelHandOver Set Mark = ""ChangQiang"" Where RowID = " & Format(i)
           mapinfo.do "Update SelHandOver Set Mark = ""场强"" Where RowID = " & Format(i)
        ElseIf Val(mapinfo.eval("SelHandOver.rxqual_s")) < Val(Text1(8)) And mapinfo.eval("SelHandOver.rxlev_s") > 110 - Val(Text1(9)) And mapinfo.eval("SelHandOver.tx_power") <= Val(MyTxPwr1) Then
           'mapinfo.do "Update SelHandOver Set Mark = ""YongSai"" Where RowID = " & Format(i)
           mapinfo.do "Update SelHandOver Set Mark = ""拥塞"" Where RowID = " & Format(i)
        ElseIf Val(mapinfo.eval("SelHandOver.rxqual_s")) < Val(Text1(12)) And mapinfo.eval("SelHandOver.rxlev_s") > 110 - Val(Text1(13)) And mapinfo.eval("SelHandOver.tx_power") <= Val(MyTxPwr2) Then
           'mapinfo.do "Update SelHandOver Set Mark = ""GanRao"" Where RowID = " & Format(i)
           mapinfo.do "Update SelHandOver Set Mark = ""干扰"" Where RowID = " & Format(i)
        Else
           mapinfo.do "Update SelHandOver Set Mark = ""N/A"" Where RowID = " & Format(i)
        End If
        mapinfo.do "fetch next from SelHandOver"
    Next
    mapinfo.do "Add Map Auto Layer ""SelHandOver"""
    mapinfo.do "shade window FrontWindow() SelHandOver With mark values ""误码"" Symbol (121,16744448,15,""Monotype Sorts"",256,0) ,""场强"" Symbol (72,10502399,15,""Monotype Sorts"",256,0) ,""拥塞"" Symbol (105,255,15,""Monotype Sorts"",256,0),""干扰"" Symbol (58,16711680,15,""Monotype Sorts"",256,0) ,""N/A"" Symbol (62,16711935,15,""Monotype Sorts"",256,0)"
    If legendid = 0 Then
       mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
       mapinfo.do "Create Legend From Window  Frontwindow()"
       legendid = mapinfo.eval("windowinfo(1009,12)")
    End If
    mapinfo.do "set legend window FrontWindow() Layer prev Title ""导致切换的可能原因 " & tblname & """ Font (""宋体"",0,9,0) Subtitle" + Chr(34) + USERNAME + Chr(34) + " Font (""宋体"",0,9,0) ascending off ranges Font (""宋体"",0,9,0) """" display off "
    '条件太多
    
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub
