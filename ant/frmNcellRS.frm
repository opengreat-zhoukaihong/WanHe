VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNcellRS 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "邻小区合理性统计"
   ClientHeight    =   5010
   ClientLeft      =   2805
   ClientTop       =   480
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNcellRS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9015
   Begin VB.CommandButton Command2 
      Caption         =   "保存结果"
      DragIcon        =   "frmNcellRS.frx":000C
      Height          =   320
      Left            =   3330
      TabIndex        =   17
      Top             =   4530
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1305
      Index           =   1
      Left            =   4530
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2985
      Width           =   4350
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   270
      Index           =   0
      Left            =   2085
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   570
      Width           =   5730
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭"
      DragIcon        =   "frmNcellRS.frx":015E
      Height          =   320
      Left            =   4590
      TabIndex        =   0
      Top             =   4530
      Width           =   1080
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1305
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   1320
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   2302
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "小区中文名"
         Object.Width           =   1552
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "BCCH"
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "BSIC"
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "RxLev"
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "解码度"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "测量数"
         Object.Width           =   794
      EndProperty
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1305
      Index           =   1
      Left            =   135
      TabIndex        =   15
      Top             =   2985
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   2302
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "小区中文名"
         Object.Width           =   1552
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "BCCH"
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "BSIC"
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "RxLev"
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "解码度"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "测量数"
         Object.Width           =   794
      EndProperty
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1305
      Index           =   2
      Left            =   4530
      TabIndex        =   16
      Top             =   1320
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   2302
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "小区中文名"
         Object.Width           =   1552
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "BCCH"
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "BSIC"
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "RxLev"
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "解码度"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "测量数"
         Object.Width           =   794
      EndProperty
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   5835
      Top             =   4485
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "ARFCN："
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   8
      Left            =   3390
      TabIndex        =   19
      Top             =   225
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Label2"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   3
      Left            =   4050
      TabIndex        =   18
      Top             =   225
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "最差邻小区（前三个）："
      Height          =   180
      Index           =   7
      Left            =   4515
      TabIndex        =   14
      Top             =   1080
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "最佳邻小区（前三个）："
      Height          =   180
      Index           =   6
      Left            =   135
      TabIndex        =   13
      Top             =   1080
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "最强场强邻小区（前三个）："
      Height          =   180
      Index           =   5
      Left            =   135
      TabIndex        =   12
      Top             =   2745
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "未出现邻小区ARFCN："
      Height          =   180
      Index           =   4
      Left            =   4545
      TabIndex        =   10
      Top             =   2745
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "网络定义邻小区ARFCN："
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   3
      Left            =   165
      TabIndex        =   8
      Top             =   585
      Width           =   1890
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Label2"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   2
      Left            =   5460
      TabIndex        =   7
      Top             =   225
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "LAC："
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   2
      Left            =   4980
      TabIndex        =   6
      Top             =   225
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Label2"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   1
      Left            =   2460
      TabIndex        =   5
      Top             =   225
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "CI："
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   1
      Left            =   2070
      TabIndex        =   4
      Top             =   225
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Label2"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   0
      Left            =   900
      TabIndex        =   3
      Top             =   225
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "小区名："
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   225
      Width           =   720
   End
End
Attribute VB_Name = "frmNcellRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DCSFlag As Boolean

Private Sub Command1_Click()
    
    On Error Resume Next
    Unload Me

End Sub

Private Sub Command2_Click()
    Dim hFreefile As Integer
    Dim MyTabName As String
    Dim MyOutString As String
    
    On Error Resume Next
                  FileDialog.filename = Trim(Label2(0).Caption)
                  FileDialog.Filter = "*.tab Files|*.txt"
                  FileDialog.DefaultExt = "*.txt"
                  FileDialog.Flags = &H80000
                  FileDialog.InitDir = Gsm_Path
                  FileDialog.CancelError = True
                  FileDialog.ShowSave
                  If Err Then
                      Err = 0
                      GoTo ErrExit
                  End If
                  If FileDialog.filename <> "" Then
                     GoTo NewSave
                        Select Case Me.Tag
                            Case 1
                                MyTabName = "NcellRSTemp_1"
                                mapinfo.do "commit table NcellRSTemp_1 as " + Chr(34) + FileDialog.filename + Chr(34)
                            Case 2
                                MyTabName = "NcellRSTemp_2"
                                mapinfo.do "commit table NcellRSTemp_2 as " + Chr(34) + FileDialog.filename + Chr(34)
                            Case 3
                                MyTabName = "NcellRSTemp_3"
                                mapinfo.do "commit table NcellRSTemp_3 as " + Chr(34) + FileDialog.filename + Chr(34)
                        End Select
                        Exit Sub
NewSave:
                        If Dir(FileDialog.filename, 0) <> "" Then
                            Kill FileDialog.filename
                        End If
                        hFreefile = FreeFile
                        Open FileDialog.filename For Binary As #hFreefile
                        MyOutString = "============ 邻小区合理性统计 ============" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        MyOutString = "             小区名：" & Label2(0) & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        MyOutString = "                 CI：" & Label2(1) & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        MyOutString = "              ARFCN：" & Label2(3) & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        MyOutString = "                LAC：" & Label2(2) & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        MyOutString = "网络定义邻小区ARFCN：" & Text1(0).Text & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        MyOutString = "  未出现邻小区ARFCN：" & Text1(1).Text & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        
                        MyOutString = "====最差邻小区（前三个）====" & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        MyOutString = "    小区中文名    BCCH   BSIC   Rxlev   解码度   测量数" & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        If ListView1(2).ListItems(1).Text = "" Then
                            MyOutString = "     " & String(14, " ") & ListView1(2).ListItems(1).SubItems(1) & String(7 - Len(ListView1(2).ListItems(1).SubItems(1)), " ") & ListView1(2).ListItems(1).SubItems(2) & String(7 - Len(ListView1(2).ListItems(1).SubItems(2)), " ") & ListView1(2).ListItems(1).SubItems(3) & String(8 - Len(ListView1(2).ListItems(1).SubItems(3)), " ") & ListView1(2).ListItems(1).SubItems(4) & String(9 - Len(ListView1(2).ListItems(1).SubItems(4)), " ") & ListView1(2).ListItems(1).SubItems(5) & Chr(13) & Chr(10)
                        Else
                            MyOutString = "     " & ListView1(2).ListItems(1).Text & String(15 - LenB(ListView1(2).ListItems(1).Text), " ") & ListView1(2).ListItems(1).SubItems(1) & String(7 - Len(ListView1(2).ListItems(1).SubItems(1)), " ") & ListView1(2).ListItems(1).SubItems(2) & String(7 - Len(ListView1(2).ListItems(1).SubItems(2)), " ") & ListView1(2).ListItems(1).SubItems(3) & String(8 - Len(ListView1(2).ListItems(1).SubItems(3)), " ") & ListView1(2).ListItems(1).SubItems(4) & String(9 - Len(ListView1(2).ListItems(1).SubItems(4)), " ") & ListView1(2).ListItems(1).SubItems(5) & Chr(13) & Chr(10)
                        End If
                        Put #hFreefile, , MyOutString
                        'MyOutString = "      " & ListView1(2).ListItems(2).Text & "    " & ListView1(2).ListItems(2).SubItems(1) & "    " & ListView1(2).ListItems(2).SubItems(2) & "   " & ListView1(2).ListItems(2).SubItems(3) & "   " & ListView1(2).ListItems(2).SubItems(4) & "   " & ListView1(2).ListItems(2).SubItems(5) & Chr(13) & Chr(10)
                        If ListView1(2).ListItems(2).Text = "" Then
                            MyOutString = "     " & String(14, " ") & ListView1(2).ListItems(2).SubItems(1) & String(7 - Len(ListView1(2).ListItems(2).SubItems(1)), " ") & ListView1(2).ListItems(2).SubItems(2) & String(7 - Len(ListView1(2).ListItems(2).SubItems(2)), " ") & ListView1(2).ListItems(2).SubItems(3) & String(8 - Len(ListView1(2).ListItems(2).SubItems(3)), " ") & ListView1(2).ListItems(2).SubItems(4) & String(9 - Len(ListView1(2).ListItems(2).SubItems(4)), " ") & ListView1(2).ListItems(2).SubItems(5) & Chr(13) & Chr(10)
                        Else
                            MyOutString = "     " & ListView1(2).ListItems(2).Text & String(15 - LenB(ListView1(2).ListItems(2).Text), " ") & ListView1(2).ListItems(2).SubItems(1) & String(7 - Len(ListView1(2).ListItems(2).SubItems(1)), " ") & ListView1(2).ListItems(2).SubItems(2) & String(7 - Len(ListView1(2).ListItems(2).SubItems(2)), " ") & ListView1(2).ListItems(2).SubItems(3) & String(8 - Len(ListView1(2).ListItems(2).SubItems(3)), " ") & ListView1(2).ListItems(2).SubItems(4) & String(9 - Len(ListView1(2).ListItems(2).SubItems(4)), " ") & ListView1(2).ListItems(2).SubItems(5) & Chr(13) & Chr(10)
                        End If
                        Put #hFreefile, , MyOutString
                        If ListView1(2).ListItems(3).Text = "" Then
                            MyOutString = "     " & String(14, " ") & ListView1(2).ListItems(3).SubItems(1) & String(7 - Len(ListView1(2).ListItems(3).SubItems(1)), " ") & ListView1(2).ListItems(3).SubItems(2) & String(7 - Len(ListView1(2).ListItems(3).SubItems(2)), " ") & ListView1(2).ListItems(3).SubItems(3) & String(8 - Len(ListView1(2).ListItems(3).SubItems(3)), " ") & ListView1(2).ListItems(3).SubItems(4) & String(9 - Len(ListView1(2).ListItems(3).SubItems(4)), " ") & ListView1(2).ListItems(3).SubItems(5) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        Else
                            MyOutString = "     " & ListView1(2).ListItems(3).Text & String(15 - LenB(ListView1(2).ListItems(3).Text), " ") & ListView1(2).ListItems(3).SubItems(1) & String(7 - Len(ListView1(2).ListItems(3).SubItems(1)), " ") & ListView1(2).ListItems(3).SubItems(2) & String(7 - Len(ListView1(2).ListItems(3).SubItems(2)), " ") & ListView1(2).ListItems(3).SubItems(3) & String(8 - Len(ListView1(2).ListItems(3).SubItems(3)), " ") & ListView1(2).ListItems(3).SubItems(4) & String(9 - Len(ListView1(2).ListItems(3).SubItems(4)), " ") & ListView1(2).ListItems(3).SubItems(5) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        End If
                        Put #hFreefile, , MyOutString
                        
                        MyOutString = "====最佳邻小区（前三个）====" & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        MyOutString = "    小区中文名    BCCH   BSIC   Rxlev   解码度   测量数" & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        If ListView1(0).ListItems(1).Text = "" Then
                            MyOutString = "     " & String(14, " ") & ListView1(0).ListItems(1).SubItems(1) & String(7 - Len(ListView1(0).ListItems(1).SubItems(1)), " ") & ListView1(0).ListItems(1).SubItems(2) & String(7 - Len(ListView1(0).ListItems(1).SubItems(2)), " ") & ListView1(0).ListItems(1).SubItems(3) & String(8 - Len(ListView1(0).ListItems(1).SubItems(3)), " ") & ListView1(0).ListItems(1).SubItems(4) & String(9 - Len(ListView1(0).ListItems(1).SubItems(4)), " ") & ListView1(0).ListItems(1).SubItems(5) & Chr(13) & Chr(10)
                        Else
                            MyOutString = "     " & ListView1(0).ListItems(1).Text & String(15 - LenB(ListView1(0).ListItems(1).Text), " ") & ListView1(0).ListItems(1).SubItems(1) & String(7 - Len(ListView1(0).ListItems(1).SubItems(1)), " ") & ListView1(0).ListItems(1).SubItems(2) & String(7 - Len(ListView1(0).ListItems(1).SubItems(2)), " ") & ListView1(0).ListItems(1).SubItems(3) & String(8 - Len(ListView1(0).ListItems(1).SubItems(3)), " ") & ListView1(0).ListItems(1).SubItems(4) & String(9 - Len(ListView1(0).ListItems(1).SubItems(4)), " ") & ListView1(0).ListItems(1).SubItems(5) & Chr(13) & Chr(10)
                        End If
                        Put #hFreefile, , MyOutString
                        'MyOutString = "      " & ListView1(0).ListItems(2).Text & "    " & ListView1(0).ListItems(2).SubItems(1) & "    " & ListView1(0).ListItems(2).SubItems(2) & "   " & ListView1(0).ListItems(2).SubItems(3) & "   " & ListView1(0).ListItems(2).SubItems(4) & "   " & ListView1(0).ListItems(2).SubItems(5) & Chr(13) & Chr(10)
                        If ListView1(0).ListItems(2).Text = "" Then
                            MyOutString = "     " & String(14, " ") & ListView1(0).ListItems(2).SubItems(1) & String(7 - Len(ListView1(0).ListItems(2).SubItems(1)), " ") & ListView1(0).ListItems(2).SubItems(2) & String(7 - Len(ListView1(0).ListItems(2).SubItems(2)), " ") & ListView1(0).ListItems(2).SubItems(3) & String(8 - Len(ListView1(0).ListItems(2).SubItems(3)), " ") & ListView1(0).ListItems(2).SubItems(4) & String(9 - Len(ListView1(0).ListItems(2).SubItems(4)), " ") & ListView1(0).ListItems(2).SubItems(5) & Chr(13) & Chr(10)
                        Else
                            MyOutString = "     " & ListView1(0).ListItems(2).Text & String(15 - LenB(ListView1(0).ListItems(2).Text), " ") & ListView1(0).ListItems(2).SubItems(1) & String(7 - Len(ListView1(0).ListItems(2).SubItems(1)), " ") & ListView1(0).ListItems(2).SubItems(2) & String(7 - Len(ListView1(0).ListItems(2).SubItems(2)), " ") & ListView1(0).ListItems(2).SubItems(3) & String(8 - Len(ListView1(0).ListItems(2).SubItems(3)), " ") & ListView1(0).ListItems(2).SubItems(4) & String(9 - Len(ListView1(0).ListItems(2).SubItems(4)), " ") & ListView1(0).ListItems(2).SubItems(5) & Chr(13) & Chr(10)
                        End If
                        Put #hFreefile, , MyOutString
                        If ListView1(0).ListItems(3).Text = "" Then
                            MyOutString = "     " & String(14, " ") & ListView1(0).ListItems(3).SubItems(1) & String(7 - Len(ListView1(0).ListItems(3).SubItems(1)), " ") & ListView1(0).ListItems(3).SubItems(2) & String(7 - Len(ListView1(0).ListItems(3).SubItems(2)), " ") & ListView1(0).ListItems(3).SubItems(3) & String(8 - Len(ListView1(0).ListItems(3).SubItems(3)), " ") & ListView1(0).ListItems(3).SubItems(4) & String(9 - Len(ListView1(0).ListItems(3).SubItems(4)), " ") & ListView1(0).ListItems(3).SubItems(5) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        Else
                            MyOutString = "     " & ListView1(0).ListItems(3).Text & String(15 - LenB(ListView1(0).ListItems(3).Text), " ") & ListView1(0).ListItems(3).SubItems(1) & String(7 - Len(ListView1(0).ListItems(3).SubItems(1)), " ") & ListView1(0).ListItems(3).SubItems(2) & String(7 - Len(ListView1(0).ListItems(3).SubItems(2)), " ") & ListView1(0).ListItems(3).SubItems(3) & String(8 - Len(ListView1(0).ListItems(3).SubItems(3)), " ") & ListView1(0).ListItems(3).SubItems(4) & String(9 - Len(ListView1(0).ListItems(3).SubItems(4)), " ") & ListView1(0).ListItems(3).SubItems(5) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        End If
                        Put #hFreefile, , MyOutString
                        
                        MyOutString = "==场强最强邻小区（前三个）==" & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        MyOutString = "    小区中文名    BCCH   BSIC   Rxlev   解码度   测量数" & Chr(13) & Chr(10)
                        Put #hFreefile, , MyOutString
                        If ListView1(1).ListItems(1).Text = "" Then
                            MyOutString = "     " & String(14, " ") & ListView1(1).ListItems(1).SubItems(1) & String(7 - Len(ListView1(1).ListItems(1).SubItems(1)), " ") & ListView1(1).ListItems(1).SubItems(2) & String(7 - Len(ListView1(1).ListItems(1).SubItems(2)), " ") & ListView1(1).ListItems(1).SubItems(3) & String(8 - Len(ListView1(1).ListItems(1).SubItems(3)), " ") & ListView1(1).ListItems(1).SubItems(4) & String(9 - Len(ListView1(1).ListItems(1).SubItems(4)), " ") & ListView1(1).ListItems(1).SubItems(5) & Chr(13) & Chr(10)
                        Else
                            MyOutString = "     " & ListView1(1).ListItems(1).Text & String(15 - LenB(ListView1(1).ListItems(1).Text), " ") & ListView1(1).ListItems(1).SubItems(1) & String(7 - Len(ListView1(1).ListItems(1).SubItems(1)), " ") & ListView1(1).ListItems(1).SubItems(2) & String(7 - Len(ListView1(1).ListItems(1).SubItems(2)), " ") & ListView1(1).ListItems(1).SubItems(3) & String(8 - Len(ListView1(1).ListItems(1).SubItems(3)), " ") & ListView1(1).ListItems(1).SubItems(4) & String(9 - Len(ListView1(1).ListItems(1).SubItems(4)), " ") & ListView1(1).ListItems(1).SubItems(5) & Chr(13) & Chr(10)
                        End If
                        Put #hFreefile, , MyOutString
                        'MyOutString = "      " & ListView1(1).ListItems(2).Text & "    " & ListView1(1).ListItems(2).SubItems(1) & "    " & ListView1(1).ListItems(2).SubItems(2) & "   " & ListView1(1).ListItems(2).SubItems(3) & "   " & ListView1(1).ListItems(2).SubItems(4) & "   " & ListView1(1).ListItems(2).SubItems(5) & Chr(13) & Chr(10)
                        If ListView1(1).ListItems(2).Text = "" Then
                            MyOutString = "     " & String(14, " ") & ListView1(1).ListItems(2).SubItems(1) & String(7 - Len(ListView1(1).ListItems(2).SubItems(1)), " ") & ListView1(1).ListItems(2).SubItems(2) & String(7 - Len(ListView1(1).ListItems(2).SubItems(2)), " ") & ListView1(1).ListItems(2).SubItems(3) & String(8 - Len(ListView1(1).ListItems(2).SubItems(3)), " ") & ListView1(1).ListItems(2).SubItems(4) & String(9 - Len(ListView1(1).ListItems(2).SubItems(4)), " ") & ListView1(1).ListItems(2).SubItems(5) & Chr(13) & Chr(10)
                        Else
                            MyOutString = "     " & ListView1(1).ListItems(2).Text & String(15 - LenB(ListView1(1).ListItems(2).Text), " ") & ListView1(1).ListItems(2).SubItems(1) & String(7 - Len(ListView1(1).ListItems(2).SubItems(1)), " ") & ListView1(1).ListItems(2).SubItems(2) & String(7 - Len(ListView1(1).ListItems(2).SubItems(2)), " ") & ListView1(1).ListItems(2).SubItems(3) & String(8 - Len(ListView1(1).ListItems(2).SubItems(3)), " ") & ListView1(1).ListItems(2).SubItems(4) & String(9 - Len(ListView1(1).ListItems(2).SubItems(4)), " ") & ListView1(1).ListItems(2).SubItems(5) & Chr(13) & Chr(10)
                        End If
                        Put #hFreefile, , MyOutString
                        If ListView1(1).ListItems(3).Text = "" Then
                            MyOutString = "     " & String(14, " ") & ListView1(1).ListItems(3).SubItems(1) & String(7 - Len(ListView1(1).ListItems(3).SubItems(1)), " ") & ListView1(1).ListItems(3).SubItems(2) & String(7 - Len(ListView1(1).ListItems(3).SubItems(2)), " ") & ListView1(1).ListItems(3).SubItems(3) & String(8 - Len(ListView1(1).ListItems(3).SubItems(3)), " ") & ListView1(1).ListItems(3).SubItems(4) & String(9 - Len(ListView1(1).ListItems(3).SubItems(4)), " ") & ListView1(1).ListItems(3).SubItems(5)
                        Else
                            MyOutString = "     " & ListView1(1).ListItems(3).Text & String(15 - LenB(ListView1(1).ListItems(3).Text), " ") & ListView1(1).ListItems(3).SubItems(1) & String(7 - Len(ListView1(1).ListItems(3).SubItems(1)), " ") & ListView1(1).ListItems(3).SubItems(2) & String(7 - Len(ListView1(1).ListItems(3).SubItems(2)), " ") & ListView1(1).ListItems(3).SubItems(3) & String(8 - Len(ListView1(1).ListItems(3).SubItems(3)), " ") & ListView1(1).ListItems(3).SubItems(4) & String(9 - Len(ListView1(1).ListItems(3).SubItems(4)), " ") & ListView1(1).ListItems(3).SubItems(5)
                        End If
                        Put #hFreefile, , MyOutString
                        
                        Close #hFreefile
                  End If
ErrExit:
                  FileDialog.CancelError = False
                  FileDialog.filename = ""
    
End Sub

Private Sub Form_Load()
    Dim itmX As ListItem
    Dim i As Integer, MyArrayCount As Integer, mm As Integer
    Dim MyRows As Long, j As Long
    Dim MyNcellBcch(31) As Integer, MyNcellBsic(31) As Integer, MyNcellRxlev(31) As Long
    Dim MyValidCount(31) As Long, MyInvalidCount(31) As Long
    Dim IsFind As Boolean
    Dim MyNCellName As String
    Dim DefineBcch As String
    Dim MyHexString As String
    Dim MyHexTemp As String
    Dim Mybcchtmp As Integer, Mybsictmp As Integer
    Dim MyTableNum As Integer
    Dim CellIsOpen As Boolean
    Dim NRSelection As String
    Dim NcellRSTemp As String
    
    On Error Resume Next
    Select Case CurrentNcellRS
        Case 1
            NRSelection = "NRSelection_1"
            NcellRSTemp = "NcellRSTemp_1"
        Case 2
            NRSelection = "NRSelection_2"
            NcellRSTemp = "NcellRSTemp_2"
        Case 3
            NRSelection = "NRSelection_3"
            NcellRSTemp = "NcellRSTemp_3"
    End Select
    MyTableNum = mapinfo.eval("NumTables()")
    For i = 1 To MyTableNum
        If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "CELL" Then
           CellIsOpen = True
           Exit For
        End If
    Next
    If CellIsOpen Then
        Call SearchCellName(0, 0, 0, 0, MyNRSelCellName, MyNRSelCellCI, "")
        mapinfo.do "x2=x1"
        mapinfo.do "y2=y1"
    Else
        MyNRSelCellName = ""
        mapinfo.do "x2=0"
        mapinfo.do "y2=0"
    End If
    Label2(0).Caption = MyNRSelCellName
    Label2(1).Caption = MyNRSelCellCI
    Label2(2).Caption = MyNRSelCellLac
    Label2(3).Caption = Format(MyNRSelCellBcch)
    
'    mapinfo.do "select * from " & NRSelection & " where bcch_serv=" & MyNRSelCellBcch & " and bsic_serv=" & MyNRSelCellBsic & " into Linmytemp"
'    If mapinfo.eval("tableinfo(Linmytemp,8)") = 0 Then
'        mapinfo.do "select bcch_serv,bsic_serv,count(*) from " & NRSelection & " group by bcch_serv,bsic_serv 3 desc into Linmytemp"
'        'If mapinfo.eval("tableinfo(Linmytemp,8)") > 1 Then
'            Mybcchtmp = mapinfo.eval("Linmytemp.bcch_serv")
'            Mybsictmp = mapinfo.eval("" & NRSelection & ".bsic_serv")
'            'mapinfo.do "select * from " & NRSelection & " where bcch_serv=" & Mybcchtmp & " and bsic_serv=" & Mybsictmp & " into " & NRSelection & ""
'            mapinfo.do "select * from " & NRSelection & " where bcch_serv=" & Mybcchtmp & " and bsic_serv=" & Mybsictmp & " and not(bcch_n1=" & Mybcchtmp & " and bsic_n1=" & Mybsictmp & ") and not(bcch_n2=" & Mybcchtmp & " and bsic_n2=" & Mybsictmp & ") and not(bcch_n3=" & Mybcchtmp & " and bsic_n3=" & Mybsictmp & ") and not(bcch_n4=" & Mybcchtmp & " and bsic_n4=" & Mybsictmp & ") and not(bcch_n5=" & Mybcchtmp & " and bsic_n5=" & Mybsictmp & ") and not(bcch_n6=" & Mybcchtmp & " and bsic_n6=" & Mybsictmp & ") into " & NRSelection & ""
'        'End If
'    Else
'        'mapinfo.do "select * from " & NRSelection & " where bcch_serv=" & MyNRSelCellBcch & " and bsic_serv=" & MyNRSelCellBsic & " into " & NRSelection & ""
'        mapinfo.do "select * from " & NRSelection & " where bcch_serv=" & MyNRSelCellBcch & " and bsic_serv=" & MyNRSelCellBsic & " and not(bcch_n1=" & MyNRSelCellBcch & " and bsic_n1=" & MyNRSelCellBsic & ") and not(bcch_n2=" & MyNRSelCellBcch & " and bsic_n2=" & MyNRSelCellBsic & ") and not(bcch_n3=" & MyNRSelCellBcch & " and bsic_n3=" & MyNRSelCellBsic & ") and not(bcch_n4=" & MyNRSelCellBcch & " and bsic_n4=" & MyNRSelCellBsic & ") and not(bcch_n5=" & MyNRSelCellBcch & " and bsic_n5=" & MyNRSelCellBsic & ") and not(bcch_n6=" & MyNRSelCellBcch & " and bsic_n6=" & MyNRSelCellBsic & ") into " & NRSelection & ""
'    End If
    
    mapinfo.do "select bcch_serv,hex_string from " & NRSelection & " where message=""SYSTEM INFORMATION TYPE 5"" into SYSTEM5"
    If mapinfo.eval("tableinfo(SYSTEM5,8)") = 0 Then
        DefineBcch = ""
    Else
        If mapinfo.eval("SYSTEM5.bcch_serv") > 511 Then
            DCSFlag = True
        End If
        MyHexString = mapinfo.eval("SYSTEM5.hex_string")
        If InStr(UCase(MyHexString), "06 1D") > 0 Then
            MyHexString = Right(MyHexString, Len(MyHexString) - InStr(UCase(MyHexString), "06 1D") + 1 - 6)
            MyHexTemp = ""
            For i = 1 To Len(MyHexString)
                If Mid(MyHexString, i, 1) <> " " Then
                    MyHexTemp = MyHexTemp & Mid(MyHexString, i, 1)
                End If
            Next
            DefineBcch = System_ReadArfcn(MyHexTemp)
            Text1(0).Text = DefineBcch
        End If
    End If
    mapinfo.do "close table SYSTEM5"
    MyArrayCount = 0
    For i = 1 To 6
        mapinfo.do "select bcch_n" & Format(i) & ",bsic_n" & Format(i) & ",Avg(rxlev_n" & Format(i) & ") ,Count(*) from " & NRSelection & " where bcch_n" & Format(i) & "<>0 group by bcch_n" & Format(i) & ",bsic_n" & Format(i) & " order by bcch_n" & Format(i) & ",bsic_n" & Format(i) & " into LinMytemp"
        MyRows = mapinfo.eval("tableinfo(LinMytemp,8)")
        mapinfo.do "fetch first from LinMytemp"
        For j = 1 To MyRows
            IsFind = False
            If MyArrayCount > 0 Then
                For mm = 0 To MyArrayCount - 1
                    If mapinfo.eval("LinMytemp.col1") = MyNcellBcch(mm) Then
                        If mapinfo.eval("LinMytemp.col2") = 99 Or mapinfo.eval("LinMytemp.col2") = MyNcellBsic(mm) Then
                            MyNcellRxlev(mm) = (MyNcellRxlev(mm) * (MyInvalidCount(mm) + MyValidCount(mm)) + mapinfo.eval("LinMytemp.col3") * mapinfo.eval("LinMytemp.col4")) / (mapinfo.eval("LinMytemp.col4") + MyInvalidCount(mm) + MyValidCount(mm))
                            If mapinfo.eval("LinMytemp.col2") = 99 Then
                                MyInvalidCount(mm) = MyInvalidCount(mm) + mapinfo.eval("LinMytemp.col4")
                            Else
                                MyValidCount(mm) = MyValidCount(mm) + mapinfo.eval("LinMytemp.col4")
                            End If
                            IsFind = True
                            Exit For
                        End If
                    End If
                Next
            End If
            If MyArrayCount = 0 Or Not IsFind Then
                MyNcellBcch(MyArrayCount) = mapinfo.eval("LinMytemp.col1")
                If mapinfo.eval("LinMytemp.col2") = 99 Then
                    MyInvalidCount(MyArrayCount) = mapinfo.eval("LinMytemp.col4")
                Else
                    MyValidCount(MyArrayCount) = mapinfo.eval("LinMytemp.col4")
                    MyNcellBsic(MyArrayCount) = mapinfo.eval("LinMytemp.col2")
                End If
                MyNcellRxlev(MyArrayCount) = mapinfo.eval("LinMytemp.col3")
                MyArrayCount = MyArrayCount + 1
            End If
            mapinfo.do "fetch next from LinMytemp"
        Next
    Next
    mapinfo.do "Create Table """ & NcellRSTemp & """ (Bcch Decimal(3,0),Bsic Decimal(3,0),Rxlev Decimal(3,0),Vaild Decimal(3,0),Invaild Decimal(3,0),validity Decimal(3,0)) file " & Chr(34) & Gsm_Path & "\" & NcellRSTemp & ".tab" & Chr(34) & " TYPE NATIVE Charset " + Chr(34) + "WindowsSimpChinese" + Chr(34)
    mapinfo.do "open table """ & Gsm_Path & "\" & NcellRSTemp & ".tab"""
    For i = 0 To 31
        If MyNcellBcch(i) = 0 Then
            Exit For
        End If
        mapinfo.do "insert into " & NcellRSTemp & " (col1,col2,col3,col4,col5,col6) values (" & Format(MyNcellBcch(i)) & "," & Format(MyNcellBsic(i)) & "," & Format(MyNcellRxlev(i)) & "," & Format(MyValidCount(i)) & "," & Format(MyInvalidCount(i)) & "," & Format(MyValidCount(i) / (MyValidCount(i) + MyInvalidCount(i)) * 100) & " )"
    Next
    mapinfo.do "commit table " & NcellRSTemp
    mapinfo.do "select * from " & NcellRSTemp & " order by Rxlev desc into LinMytemp"
    MyRows = mapinfo.eval("tableinfo(LinMytemp,8)")
    If MyRows > 3 Then
        MyRows = 3
    End If
    For i = 1 To MyRows
        If CellIsOpen Then
            Call SearchCellName(mapinfo.eval("LinMytemp.bsic"), mapinfo.eval("LinMytemp.bcch"), mapinfo.eval("x2"), mapinfo.eval("y2"), MyNCellName, "", "")
        Else
            MyNCellName = ""
        End If
        Set itmX = ListView1(1).ListItems.ADD(, , CStr(MyNCellName))
        itmX.SubItems(1) = mapinfo.eval("LinMytemp.bcch")
        If mapinfo.eval("LinMytemp.bsic") = 0 Then
            itmX.SubItems(2) = "**"
        Else
            itmX.SubItems(2) = mapinfo.eval("LinMytemp.bsic")
        End If
        itmX.SubItems(3) = mapinfo.eval("LinMytemp.rxlev")
        itmX.SubItems(4) = mapinfo.eval("LinMytemp.validity")
        itmX.SubItems(5) = mapinfo.eval("LinMytemp.Vaild") + mapinfo.eval("LinMytemp.invaild")
        mapinfo.do "fetch next from LinMytemp"
    Next
    'mapinfo.do "select * from " & NcellRSTemp & " order by validity desc into LinMytemp"
    mapinfo.do "select * from " & NcellRSTemp & " order by validity desc,rxlev desc into LinMytemp"
    MyRows = mapinfo.eval("tableinfo(LinMytemp,8)")
    If MyRows > 3 Then
        MyRows = 3
    End If
    For i = 1 To MyRows
        If CellIsOpen Then
            Call SearchCellName(mapinfo.eval("LinMytemp.bsic"), mapinfo.eval("LinMytemp.bcch"), mapinfo.eval("x2"), mapinfo.eval("y2"), MyNCellName, "", "")
        Else
            MyNCellName = ""
        End If
        Set itmX = ListView1(0).ListItems.ADD(, , CStr(MyNCellName))
        itmX.SubItems(1) = mapinfo.eval("LinMytemp.bcch")
        If mapinfo.eval("LinMytemp.bsic") = 0 Then
            itmX.SubItems(2) = "**"
        Else
            itmX.SubItems(2) = mapinfo.eval("LinMytemp.bsic")
        End If
        itmX.SubItems(3) = mapinfo.eval("LinMytemp.rxlev")
        itmX.SubItems(4) = mapinfo.eval("LinMytemp.validity")
        itmX.SubItems(5) = mapinfo.eval("LinMytemp.Vaild") + mapinfo.eval("LinMytemp.invaild")
        mapinfo.do "fetch next from LinMytemp"
    Next
        
    mapinfo.do "select * from " & NcellRSTemp & " order by validity,rxlev into LinMytemp"
    MyRows = mapinfo.eval("tableinfo(LinMytemp,8)")
    If MyRows > 3 Then
        MyRows = 3
    End If
    For i = 1 To MyRows
        If CellIsOpen Then
            Call SearchCellName(mapinfo.eval("LinMytemp.bsic"), mapinfo.eval("LinMytemp.bcch"), mapinfo.eval("x2"), mapinfo.eval("y2"), MyNCellName, "", "")
        Else
            MyNCellName = ""
        End If
        Set itmX = ListView1(2).ListItems.ADD(, , CStr(MyNCellName))
        itmX.SubItems(1) = mapinfo.eval("LinMytemp.bcch")
        If mapinfo.eval("LinMytemp.bsic") = 0 Then
            itmX.SubItems(2) = "**"
        Else
            itmX.SubItems(2) = mapinfo.eval("LinMytemp.bsic")
        End If
        itmX.SubItems(3) = mapinfo.eval("LinMytemp.rxlev")
        itmX.SubItems(4) = mapinfo.eval("LinMytemp.validity")
        itmX.SubItems(5) = mapinfo.eval("LinMytemp.Vaild") + mapinfo.eval("LinMytemp.invaild")
        mapinfo.do "fetch next from LinMytemp"
    Next
    mapinfo.do "close table LinMytemp"
'    mapinfo.do "close table " & NcellRSTemp & ""
'    Kill Gsm_Path & "\" & NcellRSTemp & ".*"
    'mapinfo.do "close table " & NRSelection & ""
    If DefineBcch <> "" Then
        For i = 0 To 31
            If MyNcellBcch(i) = 0 Then
                Exit For
            End If
            If InStr(DefineBcch, Format(MyNcellBcch(i)) & " ") > 0 Then
                DefineBcch = Left(DefineBcch, InStr(DefineBcch, Format(MyNcellBcch(i)) & " ") - 1) & Right(DefineBcch, Len(DefineBcch) - InStr(DefineBcch, Format(MyNcellBcch(i)) & " ") - Len(Format(MyNcellBcch(i))) - 1)
            End If
        Next
        Text1(1).Text = DefineBcch
    End If
    If CurrentNcellRS > 1 Then
        mapinfo.do "close table selection"
        Exit Sub
    End If
    mapinfo.do "Add Map window FrontWindow() Layer " & NRSelection & ""
    mapinfo.do "shade window FrontWindow() " & NRSelection & " With RXLEV_s ignore 0 ranges apply all use all Symbol (39,65280,8,""MapInfo Cartographic"",0,0) 120: 35 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) , 35: 25 Symbol (39,8388736,8,""MapInfo Cartographic"",0,0) ,25: 15 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,15: 0 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0) "
    If legendid = 0 Then
        mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
        mapinfo.do "Create Legend From Window  Frontwindow()"
        legendid = mapinfo.eval("windowinfo(1009,12)")
    End If
    mapinfo.do "set legend window FrontWindow() Layer prev display on shades off symbols on lines off count on Title " + Chr(34) + "邻小区合理性统计 " + tblname + Chr(34) + " Font (""宋体"",0,9,0) Subtitle" + Chr(34) + "小区 " & MyNRSelCellName & " 覆盖" + Chr(34) + " Font (""宋体"",0,9,255) ascending on ranges Font (""宋体"",0,9,0) ""其余全部"" display off ,""0 至 15 (-110至-95dBm)"" display on ,""15 至 25 (-95至-85dBm)"" display on ,""25 至 35 (-85至-75dBm)"" display on ,""35 以上 (大于-75dBm)"" display on"
    mapinfo.do "close table selection"
End Sub

Function System_ReadArfcn(HexArfcn As String) As String
    Dim i As Integer, j As Integer
    Dim ReturnStr As String
    Dim ArfcnValue As Integer
    Dim Mystring As String
    Dim k As Integer
    Dim times As Integer
    
    On Error Resume Next
    If Len(HexArfcn) < 15 Then
        Exit Function
    End If
    ReturnStr = ""
       If Not DCSFlag Then
            For i = 16 To 1 Step -1
                ArfcnValue = CDbl("&H" & Mid(HexArfcn, (i - 1) * 2 + 1, 2))
                If ArfcnValue <> 0 Then
                    For j = 1 To 8
                        If (Int(ArfcnValue / (2 ^ (j - 1))) And 1) = 1 Then
                            ReturnStr = ReturnStr & Format(8 * (16 - i) + j) & "  "
                        End If
                        If Int(ArfcnValue / (2 ^ (j - 1))) = 0 Or (i = 1 And j = 4) Then
                            Exit For
                        End If
                    Next
                End If
            Next
            System_ReadArfcn = ReturnStr
       Else
'*************
            ArfcnValue = CDbl("&H" & Mid(HexArfcn, 3, 2))
            ArfcnValue = 512 + ArfcnValue * 2
            ReturnStr = Format(ArfcnValue)
            Mystring = Trim(Mid(HexArfcn, 5, 28))
            For i = 1 To 28
                For j = 4 To 1 Step -1
                    If j = 4 Then
                        k = 8
                    ElseIf j = 3 Then
                        k = 4
                    ElseIf j = 2 Then
                        k = 2
                    ElseIf j = 1 Then
                        k = 1
                    End If
                    times = times + 1
                    If (CDbl("&H" & Mid(Mystring, i, 1)) And k) / k = 1 Then
                        If i = 1 And k = 8 Then
                            ReturnStr = ArfcnValue
                        Else
                            ReturnStr = ReturnStr & "  " & Format(ArfcnValue - 1 + times) & " "
                        End If
                    Else
                        If i = 1 And k = 8 Then
                            ReturnStr = ArfcnValue
                        End If
                    End If
                Next j
            Next i
 
        System_ReadArfcn = ReturnStr
       End If

End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Select Case Me.Tag
        Case 1
            mapinfo.do "close table NcellRSTemp_1"
            Kill Gsm_Path & "\NcellRSTemp_1.*"
        Case 2
            mapinfo.do "close table NcellRSTemp_2"
            Kill Gsm_Path & "\NcellRSTemp_2.*"
        Case 3
            mapinfo.do "close table NcellRSTemp_3"
            Kill Gsm_Path & "\NcellRSTemp_3.*"
    End Select

End Sub


