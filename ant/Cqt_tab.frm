VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form CQT_Table 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "拨打记录表"
   ClientHeight    =   4545
   ClientLeft      =   3795
   ClientTop       =   1710
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Cqt_tab.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4545
   ScaleWidth      =   6030
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   4125
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   741
      ButtonWidth     =   1111
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "下一个"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "前一个"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "最后一个"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "第一个"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "保存"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "返回"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "未接通状态"
      Height          =   1080
      Left            =   3375
      TabIndex        =   35
      Top             =   2760
      Width           =   2430
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   1305
         TabIndex        =   39
         Top             =   270
         Width           =   420
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   1305
         TabIndex        =   38
         Top             =   645
         Width           =   420
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   6
         Left            =   1725
         TabIndex        =   36
         Top             =   270
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text2(10)"
         BuddyDispid     =   196611
         BuddyIndex      =   10
         OrigLeft        =   1755
         OrigTop         =   285
         OrigRight       =   1995
         OrigBottom      =   495
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   7
         Left            =   1725
         TabIndex        =   37
         Top             =   645
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text2(11)"
         BuddyDispid     =   196611
         BuddyIndex      =   11
         OrigLeft        =   1800
         OrigTop         =   615
         OrigRight       =   2040
         OrigBottom      =   825
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "掉话:"
         Height          =   180
         Index           =   11
         Left            =   765
         TabIndex        =   41
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "未接通:"
         Height          =   180
         Index           =   12
         Left            =   585
         TabIndex        =   40
         Top             =   675
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "接通状态"
      Height          =   2565
      Left            =   3375
      TabIndex        =   16
      Top             =   120
      Width           =   2430
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   0
         Left            =   1725
         TabIndex        =   29
         Top             =   360
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text2(4)"
         BuddyDispid     =   196611
         BuddyIndex      =   4
         OrigLeft        =   1935
         OrigTop         =   405
         OrigRight       =   2175
         OrigBottom      =   615
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   1305
         TabIndex        =   22
         Top             =   360
         Width           =   420
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   1305
         TabIndex        =   21
         Top             =   720
         Width           =   420
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   1305
         TabIndex        =   20
         Top             =   1080
         Width           =   420
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   1305
         TabIndex        =   19
         Top             =   1440
         Width           =   420
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   1305
         TabIndex        =   18
         Top             =   1815
         Width           =   420
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   1305
         TabIndex        =   17
         Top             =   2175
         Width           =   420
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   1
         Left            =   1725
         TabIndex        =   30
         Top             =   720
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text2(5)"
         BuddyDispid     =   196611
         BuddyIndex      =   5
         OrigLeft        =   1935
         OrigTop         =   825
         OrigRight       =   2175
         OrigBottom      =   1035
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   2
         Left            =   1725
         TabIndex        =   31
         Top             =   1080
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text2(6)"
         BuddyDispid     =   196611
         BuddyIndex      =   6
         OrigLeft        =   1935
         OrigTop         =   1230
         OrigRight       =   2175
         OrigBottom      =   1440
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   3
         Left            =   1725
         TabIndex        =   32
         Top             =   1440
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text2(7)"
         BuddyDispid     =   196611
         BuddyIndex      =   7
         OrigLeft        =   1980
         OrigTop         =   1620
         OrigRight       =   2220
         OrigBottom      =   1830
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   4
         Left            =   1725
         TabIndex        =   33
         Top             =   1815
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text2(8)"
         BuddyDispid     =   196611
         BuddyIndex      =   8
         OrigLeft        =   1965
         OrigTop         =   2100
         OrigRight       =   2205
         OrigBottom      =   2310
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   5
         Left            =   1725
         TabIndex        =   34
         Top             =   2175
         Width           =   240
         _ExtentX        =   476
         _ExtentY        =   450
         _Version        =   327680
         BuddyControl    =   "Text2(9)"
         BuddyDispid     =   196611
         BuddyIndex      =   9
         OrigLeft        =   1935
         OrigTop         =   2445
         OrigRight       =   2175
         OrigBottom      =   2655
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "正常通话:"
         Height          =   180
         Index           =   5
         Left            =   420
         TabIndex        =   28
         Top             =   390
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "噪音加带:"
         Height          =   180
         Index           =   6
         Left            =   420
         TabIndex        =   27
         Top             =   750
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "串音加带:"
         Height          =   180
         Index           =   7
         Left            =   405
         TabIndex        =   26
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "回音加带:"
         Height          =   180
         Index           =   8
         Left            =   405
         TabIndex        =   25
         Top             =   1470
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "无话音:"
         Height          =   180
         Index           =   9
         Left            =   570
         TabIndex        =   24
         Top             =   1845
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "单方通话:"
         Height          =   180
         Index           =   10
         Left            =   390
         TabIndex        =   23
         Top             =   2205
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3720
      Left            =   195
      TabIndex        =   1
      Top             =   120
      Width           =   3030
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   345
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1350
         TabIndex        =   3
         Top             =   3150
         Width           =   645
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1350
         TabIndex        =   2
         Top             =   2715
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "小区:"
         Height          =   180
         Index           =   0
         Left            =   780
         TabIndex        =   15
         Top             =   915
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "每线话务量:"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1245
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "拥塞率:"
         Height          =   180
         Index           =   2
         Left            =   585
         TabIndex        =   13
         Top             =   1635
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "掉话率:"
         Height          =   180
         Index           =   3
         Left            =   570
         TabIndex        =   12
         Top             =   2010
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "话音信道:"
         Height          =   180
         Index           =   4
         Left            =   435
         TabIndex        =   11
         Top             =   405
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label3"
         Height          =   180
         Index           =   0
         Left            =   1350
         TabIndex        =   10
         Top             =   900
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label3"
         Height          =   180
         Index           =   1
         Left            =   1350
         TabIndex        =   9
         Top             =   1245
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label3"
         Height          =   180
         Index           =   2
         Left            =   1350
         TabIndex        =   8
         Top             =   1620
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label3"
         Height          =   180
         Index           =   3
         Left            =   1350
         TabIndex        =   7
         Top             =   1995
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "拨打次数:"
         Height          =   180
         Index           =   2
         Left            =   405
         TabIndex        =   6
         Top             =   3180
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "测试时间:"
         Height          =   180
         Index           =   1
         Left            =   405
         TabIndex        =   5
         Top             =   2760
         Width           =   810
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4035
      Top             =   5055
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   26
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Cqt_tab.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Cqt_tab.frx":08E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Cqt_tab.frx":0EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Cqt_tab.frx":14D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Cqt_tab.frx":1AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Cqt_tab.frx":20BC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CQT_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim current_rec As Integer
Dim CQT_Save As Boolean, CQT_Change As Boolean
Dim use_name As String

Sub CQT_to_face()
    Dim i As Integer, j As Integer
    
    On Error Resume Next
    Label3(0).Caption = mapinfo.eval(use_name & ".col2")
    Label3(1).Caption = mapinfo.eval(use_name & ".col3")
    Label3(2).Caption = mapinfo.eval(use_name & ".col4")
    Label3(3).Caption = mapinfo.eval(use_name & ".col5")
    For i = 1 To 8
        j = i + 5
        Text2(i + 3).Text = mapinfo.eval(use_name & ".col" & j)
    Next
    Text1(2).Text = mapinfo.eval(use_name & ".col16")
    Text1(1).Text = mapinfo.eval(use_name & ".col17")
End Sub

Sub CQT_to_table()
    Dim col_value(1 To 8) As Integer
    Dim i As Integer, j As Integer
    Dim Mymsg As String
    
    On Error Resume Next
    For i = 1 To 8
        col_value(i) = Val(Text2(i + 3).Text)
    Next
    Mymsg = "update " + use_name + " set "
    For i = 1 To 8
        j = i + 5
        Mymsg = Mymsg + "col" & j & " = " + str(col_value(i)) + ","
    Next
    Mymsg = Mymsg + "col16 = " + str(Val(Text1(2).Text)) + ",col17=" + Chr(34) + Text1(1).Text + Chr(34) + "where rowid=" & current_rec
    mapinfo.do Mymsg
    mapinfo.do "commit table " & use_name
End Sub

Private Sub Combo1_Click()
    On Error Resume Next
    If CQT_Change = True Then
       CQT_to_table
       CQT_Change = False
    End If
    current_rec = Combo1.ListIndex + 1
    mapinfo.do "fetch rec " & current_rec & "from " & use_name
    CQT_to_face

End Sub

Private Sub Form_Load()
    Dim finds As Integer, i As Integer
    Dim all_rows, cell_name
    
    On Error Resume Next
    CQT_Save = False
    CQT_Change = False
    use_name = convert_filename(1)
    finds = InStr(use_name, "\")
    Do While finds > 0
       use_name = Right(use_name, Len(use_name) - finds)
       finds = InStr(use_name, "\")
    Loop
    use_name = Left(use_name, Len(use_name) - 4)
    If Asc(Left(use_name, 1)) > 47 And Asc(Left(use_name, 1)) < 58 Then
       use_name = "_" + use_name
    End If
    Gsm_FileName = Left(convert_filename(1), Len(convert_filename(1)) - 4) + ".dbf"
    Gsm_File2 = Left(convert_filename(1), Len(convert_filename(1)) - 4) + ".old"
    FileCopy Gsm_FileName, Gsm_File2
    mapinfo.do "fetch first from " & use_name
    all_rows = mapinfo.eval("TABLEINFO(" + use_name + ", 8)")
    For i = 1 To all_rows
        cell_name = mapinfo.eval(use_name + ".col1")
        Combo1.AddItem cell_name
        mapinfo.do "fetch next from " & use_name
    Next
    Combo1.ListIndex = 0
End Sub

Private Sub SS_first_Click()
    On Error Resume Next
    If CQT_Change = True Then
       CQT_to_table
       CQT_Change = False
    End If
    Combo1.ListIndex = 0
End Sub

Private Sub SS_last_Click()
    On Error Resume Next
    If CQT_Change = True Then
       CQT_to_table
       CQT_Change = False
    End If
    Combo1.ListIndex = Combo1.ListCount - 1
End Sub

Private Sub SS_next_Click()
    On Error Resume Next
    If CQT_Change = True Then
       CQT_to_table
       CQT_Change = False
    End If
    If Combo1.ListIndex < Combo1.ListCount - 1 Then
       Combo1.ListIndex = Combo1.ListIndex + 1
    End If
End Sub

Private Sub SS_prev_Click()
    On Error Resume Next
    If CQT_Change = True Then
       CQT_to_table
       CQT_Change = False
    End If
    If Combo1.ListIndex > 0 Then
       Combo1.ListIndex = Combo1.ListIndex - 1
    End If

End Sub

Private Sub SS_return_Click()
    On Error Resume Next
    If CQT_Change = True Then
       CQT_to_table
       CQT_Change = False
    End If
    mapinfo.do "close table " & use_name
    Gsm_File2 = Left(convert_filename(1), Len(convert_filename(1)) - 4) + ".old"
    Gsm_FileName = Left(convert_filename(1), Len(convert_filename(1)) - 4) + ".dbf"
    If CQT_Save = True Then
       If (MsgBox("保存所做的修改吗？", 33, "提示")) <> 1 Then
           FileCopy Gsm_File2, Gsm_FileName
       End If
    End If
    Kill Gsm_File2
    Unload Me
End Sub

Private Sub SS_save_Click()
    On Error Resume Next
    If CQT_Change = True Then
       CQT_to_table
       CQT_Change = False
    End If
    CQT_Save = False
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 46 Then
       CQT_Save = True
       CQT_Change = True
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    CQT_Save = True
    CQT_Change = True
End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 46 Then
       CQT_Save = True
       CQT_Change = True
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    CQT_Save = True
    CQT_Change = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error Resume Next
    Select Case Button.Index
       Case 2
            SS_next_Click
       Case 3
            SS_prev_Click
       Case 4
            SS_last_Click
       Case 5
            SS_first_Click
       Case 7
            SS_save_Click
       Case 8
            SS_return_Click
    End Select
End Sub

Private Sub UpDown1_DownClick(Index As Integer)
    On Error Resume Next
    CQT_Change = True
    CQT_Save = True
End Sub

Private Sub UpDown1_UpClick(Index As Integer)
    On Error Resume Next
    CQT_Change = True
    CQT_Save = True
End Sub
