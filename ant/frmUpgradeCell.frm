VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmUpgradeCell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "小区库存升级"
   ClientHeight    =   3300
   ClientLeft      =   3780
   ClientTop       =   1110
   ClientWidth     =   4515
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpgradeCell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3280.702
   ScaleMode       =   0  'User
   ScaleWidth      =   4515
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   2970
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7550
            MinWidth        =   7550
            Text            =   "本次操作将对所选目录下的所有小区库进行升级"
            TextSave        =   "本次操作将对所选目录下的所有小区库进行升级"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   225
      Width           =   2790
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   3
      Top             =   2550
      Width           =   2805
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   330
      Left            =   3255
      TabIndex        =   2
      Top             =   1020
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   345
      Left            =   3255
      TabIndex        =   1
      Top             =   600
      Width           =   1065
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   270
      TabIndex        =   0
      Top             =   585
      Width           =   2790
   End
End
Attribute VB_Name = "frmUpgradeCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CellArrayIndex As Integer
Dim MyCellPath() As String

Private Sub Command1_Click()
    Dim SelectPath As String
    Dim MyName As String
    Dim i As Integer
    Dim CellHeadData As ScanHead
    
    On Error Resume Next
    SelectPath = Trim(Text1.Text)
    MyName = Dir(SelectPath & "\", vbDirectory)
    If MyName = "" Then
       MsgBox "不存在路径 " & SelectPath, 64, "小区库升级"
    Else
       ReDim MyCellPath(10) As String
       CellArrayIndex = 0
       GetCellPath (SelectPath)
       For i = 0 To UBound(MyCellPath)
           If MyCellPath(i) = "" Then
              Exit For
           Else
              Menu_Flag = 0
              'Data_Convert.Show 1
              If Not (CMIsCDD And UCase(MyCellPath(i)) = UCase(Gsm_Path & "\map\cell.dbf")) Then
                 hDbfFile = FreeFile
                 Open MyCellPath(i) For Binary As #hDbfFile
                 Get #hDbfFile, , CellHeadData
                 Close #hDbfFile
                 If CellHeadData.RecordLen <> (35 + 1) * 32 + 1 And CellHeadData.RecordLen <> 336 + 5 Then
                    Call Data_Convert.UpdateCell(MyCellPath(i))
                 End If
              End If
           End If
       Next
    End If
    
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Dir1_Change()
    On Error Resume Next
    Text1.Text = Dir1.path
    
End Sub

Private Sub Drive1_Change()
    On Error GoTo DriveHandler
    Dir1.path = Drive1.Drive
    Exit Sub

DriveHandler:
     MsgBox "无法读取磁盘驱动器 " + Drive1.Drive + Chr(10) + Chr(10) + "请检查驱动器的门是否已关好！   ", 64, "打开文件"
    'Drive1.Drive = "c:"
    Drive1.Drive = Dir1.path
    Exit Sub

End Sub

Private Sub Form_Load()
    On Error Resume Next
    Drive1.Drive = Gsm_Path
    Dir1.path = Gsm_Path
    Text1.Text = Dir1.path
    
End Sub

Sub GetCellPath(strSelectPath As String)
    Dim MySearchPath As String
    Dim MyDirName As String
    Dim DirectoryName() As String
    Dim i As Integer, DirArrayIndex As Integer
    
    On Error Resume Next
    ReDim DirectoryName(10) As String
    DirArrayIndex = 0
    MySearchPath = strSelectPath
    MyDirName = Dir(MySearchPath & "\", vbDirectory)
    Do While MyDirName <> ""
       If MyDirName <> "." And MyDirName <> ".." Then
          If (GetAttr(MySearchPath & "\" & MyDirName) And vbDirectory) = vbDirectory Then
             DirectoryName(DirArrayIndex) = MySearchPath & "\" & MyDirName
             DirArrayIndex = DirArrayIndex + 1
             If DirArrayIndex > UBound(DirectoryName) - 1 Then
                ReDim Preserve DirectoryName(UBound(DirectoryName) + 10) As String
             End If
          Else
             If UCase(MyDirName) = "CELL.DBF" Then
                MyCellPath(CellArrayIndex) = MySearchPath & "\" & MyDirName
                CellArrayIndex = CellArrayIndex + 1
                If CellArrayIndex > UBound(MyCellPath) - 1 Then
                   ReDim Preserve MyCellPath(UBound(MyCellPath) + 10) As String
                End If
             End If
          End If
       End If
       MyDirName = Dir
    Loop
    For i = 0 To UBound(DirectoryName)
        If DirectoryName(i) = "" Then
           Exit For
        End If
        GetCellPath (DirectoryName(i))
    Next
    
End Sub

