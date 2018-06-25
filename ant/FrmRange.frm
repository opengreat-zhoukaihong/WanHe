VERSION 5.00
Begin VB.Form FrmRange 
   Caption         =   "RxLev 统计范围定义"
   ClientHeight    =   2580
   ClientLeft      =   1875
   ClientTop       =   1560
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3120
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   4
      Text            =   "0"
      Top             =   1725
      Width           =   420
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   2370
      TabIndex        =   3
      Text            =   "17"
      Top             =   1740
      Width           =   420
   End
   Begin VB.ListBox List1 
      Height          =   1500
      ItemData        =   "FrmRange.frx":030A
      Left            =   180
      List            =   "FrmRange.frx":0311
      TabIndex        =   2
      Top             =   135
      Width           =   2730
   End
   Begin VB.CommandButton C_Cancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   320
      Left            =   1620
      TabIndex        =   1
      Top             =   2190
      Width           =   1080
   End
   Begin VB.CommandButton C_OK 
      Caption         =   "确定"
      Height          =   320
      Left            =   405
      TabIndex        =   0
      Top             =   2190
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "最小值："
      Height          =   180
      Left            =   255
      TabIndex        =   6
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "最大值："
      Height          =   180
      Left            =   1665
      TabIndex        =   5
      Top             =   1785
      Width           =   720
   End
   Begin VB.Menu MyPopUp 
      Caption         =   "No"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu MnuAdd 
         Caption         =   "增加"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "删除"
      End
      Begin VB.Menu Mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSystem 
         Caption         =   "系统默认值"
      End
      Begin VB.Menu MnuClient 
         Caption         =   "上次设定值"
      End
   End
End
Attribute VB_Name = "FrmRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyListIndex As Integer
Dim ModiFlag As Boolean

Sub Setting()
    On Error Resume Next
    
    List1.Clear
    List1.AddItem " 27-63   (-83<=dBm<-47)"
    List1.AddItem " 17-27   (-93<=dBm<-83)"
    List1.AddItem "  0-17   (-110<=dBm<-93)"
    List1.ListIndex = 0
End Sub

Private Sub C_Cancel_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub C_OK_Click()
    Dim i As Integer
    Dim RangeValue As String
    Dim ListString As String
    Dim SetFile As String
    Dim FileNumber As Integer
    Dim OutString1 As String, OutString2 As String
    
    On Error Resume Next
    SetFile = Gsm_Path + "\user\RxLevs.dat"
    If dir(SetFile, 0) <> "" Then
       Kill SetFile
    End If
    FileNumber = FreeFile
    Open SetFile For Binary As #FileNumber
    RangeNum = 0
    For i = 1 To List1.ListCount
        ListString = List1.List(i - 1)
        If Trim(ListString) <> "" Then
           RangeValue = Trim(Left(ListString, InStr(ListString, "(") - 1))
           RxLevRange(1, i) = Trim(Left(RangeValue, InStr(RangeValue, "-") - 1))
           RxLevRange(2, i) = Trim(Right(RangeValue, Len(RangeValue) - InStr(RangeValue, "-")))
           RangeNum = RangeNum + 1
           OutString1 = RxLevRange(1, i) & "," & RxLevRange(2, i)
           Put #FileNumber, , OutString1
           OutString1 = Chr(13)
           OutString2 = Chr(10)
           Put #FileNumber, , OutString1
           Put #FileNumber, , OutString2
           
        End If
    Next
    Close #FileNumber
    If List1.ListCount < 16 Then
       For i = List1.ListCount + 1 To 16
           RxLevRange(1, i) = ""
           RxLevRange(2, i) = ""
       Next
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim SetFile
    
    On Error Resume Next
    ModiFlag = False
    SetFile = Gsm_Path + "\user\RxLevs.dat"
    If dir(SetFile, 0) <> "" Then
       ClientSetting
    Else
       Setting
    End If
End Sub

Private Sub List1_Click()
    Dim ListString As String
    Dim RangeValue As String
    
    On Error Resume Next
    ListString = List1.List(List1.ListIndex)
    If Trim(ListString) = "" Then
       Text1.Text = ""
       Text2.Text = ""
       Exit Sub
    End If
    RangeValue = Trim(Left(ListString, InStr(ListString, "(") - 1))
    Text1.Text = Left(RangeValue, InStr(RangeValue, "-") - 1)
    Text2.Text = Trim(Right(RangeValue, Len(RangeValue) - InStr(RangeValue, "-")))
    
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 45 Then
       If List1.ListCount < 16 Then
          MnuAdd_Click
       End If
    ElseIf KeyCode = 46 Then
       If List1.ListCount > 1 Then
          MnuDelete_Click
       End If
    End If
       
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SetFile
    
    On Error Resume Next
    If Button = 2 Then
       SetFile = Gsm_Path + "\user\RxLevs.dat"
       If dir(SetFile, 0) = "" Then
          MnuClient.Enabled = False
       Else
          MnuClient.Enabled = True
       End If
       If List1.ListCount = 1 Then
          MnuDelete.Enabled = False
          MnuAdd.Enabled = True
       Else
          MnuDelete.Enabled = True
          If List1.ListCount >= 16 Then
             MnuAdd.Enabled = False
          Else
             MnuAdd.Enabled = True
          End If
       End If
       PopupMenu MyPopUp
    End If

End Sub

Private Sub MnuAdd_Click()
    On Error Resume Next
    List1.AddItem " "
    List1.ListIndex = List1.ListCount - 1
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub MnuClient_Click()
    On Error Resume Next
    ClientSetting
End Sub

Private Sub MnuDelete_Click()
    Dim RemoveIndex As Integer
    Dim ListString As String
    Dim RangeValue As String
    
    On Error Resume Next
    If List1.ListCount = 1 Then
       Exit Sub
    End If
    RemoveIndex = List1.ListIndex
    If RemoveIndex = List1.ListCount - 1 Then
       List1.ListIndex = List1.ListCount - 2
    Else
       List1.ListIndex = RemoveIndex + 1
    End If
    ListString = List1.List(List1.ListIndex)
    If Trim(ListString) <> "" Then
       RangeValue = Trim(Left(ListString, InStr(ListString, "(") - 1))
       Text1.Text = Left(RangeValue, InStr(RangeValue, "-") - 1)
       Text2.Text = Trim(Right(RangeValue, InStr(RangeValue, "-") - 1))
    Else
       Text1.Text = ""
       Text2.Text = ""
    End If
    List1.RemoveItem RemoveIndex
End Sub

Private Sub MnuSystem_Click()
    On Error Resume Next
    Setting
End Sub

Private Sub Text1_Change()
    On Error Resume Next
    MyListIndex = List1.ListIndex
    ModiFlag = True
End Sub

Private Sub Text2_Change()
    On Error Resume Next
    MyListIndex = List1.ListIndex
    ModiFlag = True
End Sub

Private Sub Text2_LostFocus()
    Dim ListString As String
    Dim MaxValue As String
    Dim RightString As String

    On Error Resume Next
    
    If Not ModiFlag Then
       Exit Sub
    End If
    ModiFlag = False
    ListString = List1.List(MyListIndex)
    If Trim(ListString) = "" Then
       If Trim(Text2.Text) <> "" And Trim(Text1.Text) <> "" Then
          ShowList
       End If
    Else
       If Trim(Text2.Text) <> "" Then
          ShowList
          Exit Sub
          MaxValue = Format(Val(Trim(Text2.Text)))
          If Len(MaxValue) = 1 Then
             MaxValue = " " & MaxValue
          End If
          RightString = Right(ListString, Len(ListString) - InStr(ListString, "(") + 1)
          ListString = Left(ListString, InStr(ListString, "-") - 1) & "-" & MaxValue
          If Len(ListString) < 9 Then
             ListString = ListString & String(9 - Len(ListString), " ")
          End If
          ListString = ListString & Left(RightString, InStr(RightString, "<-") + 1) & Format(110 - Val(MaxValue)) & ")"
          List1.List(MyListIndex) = ListString
       End If
    End If

End Sub

Private Sub Text1_LostFocus()
    Dim ListString As String
    Dim MinValue As String
    Dim RightString As String
    
    On Error Resume Next
    If Not ModiFlag Then
       Exit Sub
    End If
    ModiFlag = False
    ListString = Trim(List1.List(MyListIndex))
    If ListString = "" Then
       If Trim(Text2.Text) <> "" And Trim(Text1.Text) <> "" Then
          ShowList
       End If
    Else
       If Trim(Text1.Text) <> "" Then
          ShowList
          Exit Sub
          MinValue = Format(Val(Trim(Text1.Text)))
          If Len(MinValue) = 1 Then
             MinValue = " " & MinValue
          End If
          RightString = Right(ListString, Len(ListString) - InStr(ListString, "("))
          ListString = RTrim(" " & MinValue & "-" & Left(Right(ListString, Len(ListString) - InStr(ListString, "-")), InStr(Right(ListString, Len(ListString) - InStr(ListString, "-")), "(") - 1))
          If Len(ListString) < 9 Then
             ListString = ListString & String(9 - Len(ListString), " ")
          End If
          ListString = ListString & "(-" & Format(110 - Val(MinValue)) & Right(RightString, Len(RightString) - InStr(RightString, "<=") + 1)
          List1.List(MyListIndex) = ListString
       End If
    End If


End Sub

Sub ShowList()
    Dim ListString As String
    Dim MinValue As String, MaxValue As String
              
    On Error Resume Next
    If Val(Trim(Text1.Text)) < Val(Trim(Text2.Text)) Then
       MinValue = Format(Val(Trim(Text1.Text)))
       MaxValue = Format(Val(Trim(Text2.Text)))
    Else
       MinValue = Format(Val(Trim(Text2.Text)))
       MaxValue = Format(Val(Trim(Text1.Text)))
    End If
    If Len(MinValue) = 1 Then
       MinValue = " " & MinValue
    End If
    'MaxValue = Format(Val(Trim(Text2.Text)))
    If Len(MaxValue) = 1 Then
       MaxValue = " " & MaxValue
    End If
    ListString = " " & MinValue & "-" & MaxValue
    If Len(ListString) < 9 Then
       ListString = ListString & String(9 - Len(ListString), " ")
    End If
    ListString = ListString & "(-" & Format(110 - Val(MinValue)) & "<=dBm<-" & Format(110 - Val(MaxValue)) & ")"
    List1.List(MyListIndex) = ListString
    ListSort
End Sub

Sub ClientSetting()
    Dim SetFile As String
    Dim FileNumber As Integer
    Dim RangeString As String
    Dim MinValue As String, MaxValue As String
    Dim ListString As String
    
    On Error Resume Next
    SetFile = Gsm_Path + "\user\RxLevs.dat"
    If dir(SetFile, 0) <> "" Then
       FileNumber = FreeFile
       List1.Clear
       Open SetFile For Input As #FileNumber
       Do While Not EOF(FileNumber)
          Line Input #FileNumber, RangeString
          MinValue = Left(RangeString, InStr(RangeString, ",") - 1)
          MaxValue = Right(RangeString, Len(RangeString) - InStr(RangeString, ","))
          If Len(MinValue) = 1 Then
             MinValue = " " & MinValue
          End If
          If Len(MaxValue) = 1 Then
             MaxValue = " " & MaxValue
          End If
          ListString = " " & MinValue & "-" & MaxValue
          If Len(ListString) < 9 Then
             ListString = ListString & String(9 - Len(ListString), " ")
          End If
          ListString = ListString & "(-" & Format(110 - Val(MinValue)) & "<=dBm<-" & Format(110 - Val(MaxValue)) & ")"
          List1.AddItem ListString
       Loop
       Close #FileNumber
       List1.ListIndex = 0
    End If
End Sub

Sub ListSort()
    Dim i As Integer, j As Integer
    Dim SortRange(1 To 2, 1 To 16) As String
    Dim RangeValue As String, ListString As String
    Dim SortTemp As String, TempNum As Integer
    Dim SortNum As Integer
        
    On Error Resume Next
    If List1.ListCount = 1 Then
       Exit Sub
    End If
    SortNum = 0
    For i = 1 To List1.ListCount
        ListString = List1.List(i - 1)
        If Trim(ListString) <> "" Then
           RangeValue = Trim(Left(ListString, InStr(ListString, "(") - 1))
           SortRange(1, i) = Trim(Left(RangeValue, InStr(RangeValue, "-") - 1))
           SortRange(2, i) = ListString
           SortNum = SortNum + 1
        End If
    Next
    If SortNum < 2 Then
       Exit Sub
    End If
    List1.Clear
    For i = 1 To SortNum
        SortTemp = SortRange(1, 1)
        TempNum = 1
        For j = 1 To SortNum
            If Val(SortTemp) < Val(SortRange(1, j)) Then
               SortTemp = SortRange(1, j)
               TempNum = j
            End If
        Next
        List1.AddItem SortRange(2, TempNum)
        SortRange(1, TempNum) = "-999"
    Next
    List1.ListIndex = 0
End Sub
