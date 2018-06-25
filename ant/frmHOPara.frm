VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmHOPara 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000E&
   Caption         =   "切换前后参数显示"
   ClientHeight    =   5295
   ClientLeft      =   4590
   ClientTop       =   1290
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHOPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4845
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "切换前后参数对比"
      Height          =   2715
      Left            =   390
      TabIndex        =   3
      Top             =   2010
      Width           =   4065
      Begin VB.Shape Shape1 
         BorderColor     =   &H00008000&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   3
         Left            =   2790
         Top             =   2370
         Width           =   450
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00008000&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   2
         Left            =   2790
         Top             =   2130
         Width           =   450
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00008000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   1
         Left            =   2790
         Top             =   1890
         Width           =   770
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00008000&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   150
         Index           =   0
         Left            =   2790
         Top             =   1650
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00004040&
         Height          =   180
         Index           =   15
         Left            =   2010
         TabIndex        =   30
         Top             =   2370
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00004040&
         Height          =   180
         Index           =   14
         Left            =   2010
         TabIndex        =   29
         Top             =   2130
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00004040&
         Height          =   180
         Index           =   13
         Left            =   2010
         TabIndex        =   28
         Top             =   1875
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00004040&
         Height          =   180
         Index           =   12
         Left            =   2010
         TabIndex        =   27
         Top             =   1635
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00004040&
         Height          =   180
         Index           =   11
         Left            =   2010
         TabIndex        =   26
         Top             =   1380
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00004040&
         Height          =   180
         Index           =   10
         Left            =   2010
         TabIndex        =   25
         Top             =   1125
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00004040&
         Height          =   180
         Index           =   9
         Left            =   2010
         TabIndex        =   24
         Top             =   870
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00004040&
         Height          =   180
         Index           =   8
         Left            =   2010
         TabIndex        =   23
         Top             =   615
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   7
         Left            =   1290
         TabIndex        =   22
         Top             =   2370
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   6
         Left            =   1290
         TabIndex        =   21
         Top             =   2130
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   5
         Left            =   1290
         TabIndex        =   20
         Top             =   1875
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   4
         Left            =   1290
         TabIndex        =   19
         Top             =   1635
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   3
         Left            =   1290
         TabIndex        =   18
         Top             =   1380
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   2
         Left            =   1290
         TabIndex        =   17
         Top             =   1125
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   1
         Left            =   1290
         TabIndex        =   16
         Top             =   870
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   0
         Left            =   1290
         TabIndex        =   15
         Top             =   615
         Width           =   540
      End
      Begin VB.Line Line2 
         X1              =   2790
         X2              =   2790
         Y1              =   510
         Y2              =   2550
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1110
         X2              =   3630
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "改善/恶化"
         Height          =   180
         Index           =   11
         Left            =   2760
         TabIndex        =   14
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "后"
         Height          =   180
         Index           =   10
         Left            =   2100
         TabIndex        =   13
         Top             =   270
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "前"
         Height          =   180
         Index           =   9
         Left            =   1425
         TabIndex        =   12
         Top             =   270
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Tx_Power："
         Height          =   180
         Index           =   8
         Left            =   285
         TabIndex        =   11
         Top             =   2370
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "RxQual："
         Height          =   180
         Index           =   7
         Left            =   465
         TabIndex        =   10
         Top             =   1875
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "TA："
         Height          =   180
         Index           =   6
         Left            =   825
         TabIndex        =   9
         Top             =   2130
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "RxLev："
         Height          =   180
         Index           =   5
         Left            =   555
         TabIndex        =   8
         Top             =   1620
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "LAC："
         Height          =   180
         Index           =   4
         Left            =   735
         TabIndex        =   7
         Top             =   1380
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "BSIC："
         Height          =   180
         Index           =   3
         Left            =   630
         TabIndex        =   6
         Top             =   870
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "CI："
         Height          =   180
         Index           =   2
         Left            =   825
         TabIndex        =   5
         Top             =   1125
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "BCCH："
         Height          =   180
         Index           =   1
         Left            =   630
         TabIndex        =   4
         Top             =   615
         Width           =   540
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "关闭"
      DragIcon        =   "frmHOPara.frx":000C
      Height          =   320
      Left            =   1950
      TabIndex        =   0
      Top             =   4890
      Width           =   1080
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1545
      Left            =   375
      TabIndex        =   1
      Top             =   375
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   2725
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "切换类型"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "源小区"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "目标小区"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "切换事件列表："
      Height          =   180
      Index           =   0
      Left            =   345
      TabIndex        =   2
      Top             =   120
      Width           =   1260
   End
End
Attribute VB_Name = "frmHOPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyDatabase As Database
Dim MyRecordset As Recordset

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer, MyRows As Integer
    Dim itmX As ListItem
    Dim MyTableNum As Integer
    Dim CellIsOpen As Boolean
    Dim HOType As String
    Dim MyCellName As String
    Dim MyCival As String
    Dim j As Integer
    Dim MystrTmp As String
    Dim MyCell_S As String
    
    On Error Resume Next
    MyTableNum = mapinfo.eval("NumTables()")
    For i = 1 To MyTableNum
        If UCase(mapinfo.eval("tableinfo(" & Format(i) & ",1)")) = "CELL" Then
           CellIsOpen = True
           Exit For
        End If
    Next
    MyRows = mapinfo.eval("tableinfo(HOParameter,8)")
    mapinfo.do "fetch first from HOParameter"
    For i = 1 To MyRows
        MystrTmp = mapinfo.eval("HOParameter.mark1")
        If Left(MystrTmp, 3) = "HSC" Then
            HOType = "切换成功"
        Else
            HOType = "切换失败"
        End If
        For j = 1 To 3
            MystrTmp = Right(MystrTmp, Len(MystrTmp) - InStr(MystrTmp, ","))
        Next
        MyCival = Left(MystrTmp, InStr(MystrTmp, ",") - 1)
        If CellIsOpen Then
            Call SearchCellName(0, 0, 0, 0, MyCell_S, MyCival, "")
            If Trim(MyCell_S) = "" Then
                MyCell_S = MyCival
            End If
        Else
            MyCell_S = MyCival
        End If
'***************************切换失败时显示切换尝试载频
        If HOType = "切换失败" Then
            MystrTmp = mapinfo.eval("HOParameter.mark2")
            If InStr(MystrTmp, ";") > 0 Then
                MystrTmp = Trim(Right(MystrTmp, Len(MystrTmp) - InStr(MystrTmp, ";")))
                MyCellName = MystrTmp & "(尝试载频)"
            Else
                MyCellName = ""
            End If
'***************************切换失败时显示切换尝试载频
        Else
            For j = 1 To 8
                MystrTmp = Right(MystrTmp, Len(MystrTmp) - InStr(MystrTmp, ","))
            Next
            MyCival = Left(MystrTmp, InStr(MystrTmp, ",") - 1)
            If CellIsOpen Then
                Call SearchCellName(0, 0, 0, 0, MyCellName, MyCival, "")
                If MyCellName = "" Then
                    MyCellName = MyCival
                End If
            Else
                MyCellName = MyCival
            End If
        End If
        
        Set itmX = ListView1.ListItems.ADD(, , CStr(HOType))
        itmX.SubItems(1) = MyCell_S
        itmX.SubItems(2) = MyCellName
        mapinfo.do "fetch next from HOParameter"
    Next
    mapinfo.do "Add Map window FrontWindow() Layer HOParameter"
    mapinfo.do "shade window FrontWindow() HOParameter with left$(mark1,3) values " + Chr(34) + "HFC" + Chr(34) + " Symbol (""hand_f.bmp"",16776960,24,0)," + Chr(34) + "HSC" + Chr(34) + " Symbol (""hand_c.bmp"",16776960,24,0)"
    If legendid = 0 Then
        mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
        mapinfo.do "Create Legend From Window  Frontwindow()"
        legendid = mapinfo.eval("windowinfo(1009,12)")
    End If
    mapinfo.do "set legend window FrontWindow() Layer prev Title " + Chr(34) + "切换前后参数显示 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) Subtitle" + Chr(34) + "切换点显示" + Chr(34) + " Font(""宋体"",0,9,255) ascending off ranges Font(""宋体"",0,9,0) """" display off,""HANDOVER FAILURE"" display on,""HANDOVER COMPLETE"" display on"

    Call ListView1_ItemClick(ListView1.ListItems(1))
    
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
    Dim CurrentList As Integer
    Dim i As Integer
    Dim MystrTmp As String
    
    On Error Resume Next
    CurrentList = ListView1.SelectedItem.Index
    mapinfo.do "fetch rec " & Format(CurrentList) & " from HOParameter"
    mapinfo.do "select * from HOParameter where rowid = " & Format(CurrentList)
    MystrTmp = mapinfo.eval("HOParameter.mark1")
    For i = 1 To 12
        MystrTmp = Right(MystrTmp, Len(MystrTmp) - InStr(MystrTmp, ","))
        If i = 12 Then
            Label2(i - 1) = MystrTmp
        Else
            Label2(i - 1) = Left(MystrTmp, InStr(MystrTmp, ",") - 1)
        End If
    Next
    MystrTmp = mapinfo.eval("HOParameter.mark2")
    For i = 1 To 4
        MystrTmp = Right(MystrTmp, Len(MystrTmp) - InStr(MystrTmp, ","))
        If i = 4 Then
            Label2(i + 11) = MystrTmp
        Else
            Label2(i + 11) = Left(MystrTmp, InStr(MystrTmp, ",") - 1)
        End If
    Next
    For i = 5 To 7
        If Abs(Val(Label2(i + 8)) - Val(Label2(i))) <= 7 Then
            Shape1(i - 4).Width = Abs(Val(Label2(i + 8)) - Val(Label2(i))) * 110
        Else
            Shape1(i - 4).Width = 7 * 110
        End If
        If Val(Label2(i + 8)) - Val(Label2(i)) > 0 Then
            Shape1(i - 4).BorderColor = &HFF&
            Shape1(i - 4).FillColor = &HFF&
        Else
            Shape1(i - 4).BorderColor = &H8000&
            Shape1(i - 4).FillColor = &H8000&
        End If
    Next
    If Abs(Val(Label2(4)) - Val(Label2(12))) <= 15 Then
        Shape1(0).Width = Abs(Val(Label2(4)) - Val(Label2(12))) * 51
    Else
        Shape1(0).Width = 15 * 51
    End If
    If Val(Label2(12)) - Val(Label2(4)) < 0 Then
        Shape1(0).BorderColor = &HFF&
        Shape1(0).FillColor = &HFF&
    Else
        Shape1(0).BorderColor = &H8000&
        Shape1(0).FillColor = &H8000&
    End If
    
End Sub
