VERSION 5.00
Begin VB.Form SysChange 
   BackColor       =   &H00C0C0C0&
   Caption         =   "切换选择"
   ClientHeight    =   2640
   ClientLeft      =   2715
   ClientTop       =   2970
   ClientWidth     =   4050
   BeginProperty Font 
      Name            =   "System"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Xiqi.frx":0000
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2640
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "完全切换"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   270
      TabIndex        =   4
      Top             =   2250
      Value           =   -1  'True
      Width           =   1080
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "更换基站"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1575
      TabIndex        =   3
      Top             =   2250
      Width           =   1050
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   180
      TabIndex        =   0
      Top             =   210
      Width           =   2460
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   2835
      TabIndex        =   2
      Top             =   615
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   2835
      TabIndex        =   1
      Top             =   225
      Width           =   1080
   End
End
Attribute VB_Name = "SysChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim np_name(1 To 1000) As String * 30
Dim np_path(1 To 1000) As String * 150
Dim kk As Integer
Dim my_np As np

Private Sub Command1_Click()
    Dim get_path As String
    Dim f_sour As String
    Dim f_dest As String
    Dim path_tmp As String
    Dim leefind As Integer
    Dim mypath As String
    Dim CellHeadData As ScanHead
    
    On Error Resume Next
    get_path = Trim(np_path(List1.ListIndex + 1))
    leefind = InStr(get_path, Chr(0))
    If leefind > 0 Then
       get_path = Trim(Left(get_path, leefind - 1))
    End If
    If get_path = "" Then GoTo again
    If Right(get_path, 1) = "\" Then get_path = Left(get_path, Len(get_path) - 1)
    get_path = get_path + "\"
    path_tmp = get_path
    If Option1.Value = True Then
       get_path = get_path + "cell.*"
       Err = 0
       f_sour = Dir(get_path, 0)
       If f_sour = "" Then
          dd = MsgBox("不存在路径 " + Trim(np_path(List1.ListIndex + 1)), 48, "复制文件")
          GoTo again
       End If
       f_dest = f_sour
       f_dest = Gsm_Path + "\map\" + f_dest
       f_sour = path_tmp + f_sour
       FileCopy f_sour, f_dest
       
       Do
          f_sour = Dir
          If f_sour = "" Then GoTo end_xu
          f_dest = f_sour
          f_dest = Gsm_Path + "\map\" + f_dest
          f_sour = path_tmp + f_sour
          FileCopy f_sour, f_dest
       Loop
end_xu:
       get_path = path_tmp + "base.*"
       Err = 0
       f_sour = Dir(get_path, 0)
       If f_sour = "" Then
          dd = MsgBox("不存在路径 " + Trim(np_path(List1.ListIndex + 1)), 48, "复制文件")
          GoTo again
       End If
       f_dest = f_sour
       f_dest = Gsm_Path + "\map\" + f_dest
       f_sour = path_tmp + f_sour
       FileCopy f_sour, f_dest
       
       Do
          f_sour = Dir
          If f_sour = "" Then GoTo end_li
          f_dest = f_sour
          f_dest = Gsm_Path + "\map\" + f_dest
          f_sour = path_tmp + f_sour
          FileCopy f_sour, f_dest
       Loop
    Else
       Gsm_FileName = Gsm_Path + "\map\*.*"
       If Dir(Gsm_FileName, 0) <> "" Then
          
          MyDir = Gsm_Path + "\map\Map1"
          If Dir(MyDir, 16) = "" Then
             MkDir MyDir
          Else
             Gsm_File2 = Gsm_Path + "\map\Map1\*.*"
             If Dir(Gsm_File2, 0) <> "" Then
                Kill Gsm_File2
             End If
          End If
          
          mypath = Gsm_Path + "\map\"
Gsm_File2 = Dir(mypath, vbDirectory)   ' Retrieve the first entry.
Do While Gsm_File2 <> ""   ' Start the loop.
    ' Ignore the current directory and the encompassing directory.
    If Gsm_File2 <> "." And Gsm_File2 <> ".." Then
        ' Use bitwise comparison to make sure MyName is a directory.
       If (GetAttr(mypath & Gsm_File2) And vbDirectory) <> vbDirectory Then
          FileCopy Gsm_Path + "\map\" + Gsm_File2, Gsm_Path + "\map\Map1\" + Gsm_File2
       End If
    End If
    Gsm_File2 = Dir    ' Get next entry.
Loop
          
          Kill Gsm_FileName
       End If
       get_path = get_path + "*.*"
       Err = 0
       f_sour = Dir(get_path, 0)
       If f_sour = "" Then
          dd = MsgBox("不存在路径 " + Trim(np_path(List1.ListIndex + 1)), 48, "复制文件")
          GoTo again
       End If
       f_dest = f_sour
       f_dest = Gsm_Path + "\map\" + f_dest
       f_sour = path_tmp + f_sour
       FileCopy f_sour, f_dest
       Do
          f_sour = Dir
          If f_sour = "" Then Exit Do
          f_dest = f_sour
          f_dest = Gsm_Path + "\map\" + f_dest
          f_sour = path_tmp + f_sour
          FileCopy f_sour, f_dest
       Loop
    End If
end_li:
    Unload Me
    i = restore_street
    
       If Dir(Gsm_Path + "\map\cell.dbf", 0) <> "" Then
          hDbfFile = FreeFile
          Open Gsm_Path + "\map\cell.dbf" For Binary As #hDbfFile
          Get #hDbfFile, , CellHeadData
          Close #hDbfFile
          If CellHeadData.RecordLen <> (35 + 1) * 32 + 1 And CellHeadData.RecordLen <> 336 + 5 Then
             If (MsgBox("系统检测到你的基站库结构是旧的，" + Chr(10) + "如果不更新有些功能将不能正常使用。" + Chr(10) + Chr(10) + "想现在就更新吗？", 36, "提示")) = 6 Then
                UpdateFileName = Gsm_Path + "\map\cell.dbf"
                Menu_Flag = 9999
                Data_Convert.Show 1
             End If
          End If
       End If
    
    Call MDIMain.OPen_All_Map_Click
again:
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\switch.dat"
    If Dir(Gsm_FileName) <> "" Then
       Open Gsm_FileName For Binary As #1
       i = 1
       kk = 0
       If FileLen(Gsm_FileName) > 0 Then
          Do While i * 180 <= FileLen(Gsm_FileName)
             Get #1, , my_np
             np_name(i) = my_np.Name
             np_path(i) = my_np.path
             List1.AddItem my_np.Name
             i = i + 1
             kk = kk + 1
          Loop
          List1.ListIndex = 0
       End If
       Close
    End If

End Sub
