VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form SysDefine 
   BackColor       =   &H00C0C0C0&
   Caption         =   "资源定义"
   ClientHeight    =   3390
   ClientLeft      =   2970
   ClientTop       =   2370
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "System"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Xidi.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3390
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   2970
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   741
      ButtonWidth     =   1111
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "增加"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "删除"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "修改"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "返回"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1110
      TabIndex        =   1
      Top             =   495
      Width           =   2895
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1110
      TabIndex        =   2
      Top             =   870
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1110
      TabIndex        =   0
      Top             =   120
      Width           =   2340
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   300
      Top             =   1845
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   26
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Xidi.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Xidi.frx":091C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Xidi.frx":0F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Xidi.frx":1540
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "路径:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   585
      TabIndex        =   4
      Top             =   525
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "资源名称:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   165
      Width           =   810
   End
End
Attribute VB_Name = "SysDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim my_np(1 To 800) As np
Dim ad_fr As Boolean, xu As Boolean, xu_press As Boolean
Dim np_num As Integer, xu_num As Integer
Dim kk As Integer
Dim np_new(1 To 200) As np
Dim fp_np_new As Integer
Dim hFreeFile As Integer

Private Sub Command1_Click()    'Add
    Dim MystrTmp As String
    Dim i As Integer
    
    On Error Resume Next
    If ad_fr = True Then
       ad_fr = False
       If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" Then
          np_new(fp_np_new).Name = Trim(Text1.Text)
          np_new(fp_np_new).path = Trim(Text2.Text)
          fp_np_new = fp_np_new + 1
          List1.AddItem Trim(Text1.Text), 0
       End If
    End If
    If xu = True Then
       xu = False
       If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" Then
          If xu_num > fp_np_new - 1 Then
              my_np(xu_num - (fp_np_new - 1)).Name = Trim(Text1.Text)
              my_np(xu_num - (fp_np_new - 1)).path = Trim(Text2.Text)
          Else
              np_new(fp_np_new - xu_num).Name = Trim(Text1.Text)
              np_new(fp_np_new - xu_num).path = Trim(Text2.Text)
          End If
          List1.RemoveItem xu_num - 1
          List1.AddItem Trim(Text1.Text), xu_num - 1
          
          'List1.Clear
          'For i = fp_np_new - 1 To 1 Step -1
          'For i = 1 To fp_np_new
          '    MystrTmp = np_new(i).Name
          '    List1.AddItem Trim(MystrTmp)
          'Next
          'For i = 1 To kk
          '    MystrTmp = my_np(i).Name
          '    List1.AddItem Trim(MystrTmp)
          'Next
       End If
       List1.ListIndex = 0
    End If
    Text1.Text = ""
    'Text2.Text = ""
    Text2.Text = Gsm_Path & "\map\"
    ad_fr = True
    Text1.SetFocus
    
End Sub

Private Sub Command2_Click()   'Delete
    Dim i As Integer
    
    On Error Resume Next
    If ad_fr = True Then
       ad_fr = False
       List1.ListIndex = 0
       Exit Sub
    End If
    If xu = True Then
       xu = False
       List1.ListIndex = 0
       Exit Sub
    End If
    If List1.ListCount > 0 Then
       'If np_num = 0 Then
          If np_num > fp_np_new - 1 Then
             For i = np_num To kk
                 my_np(i).Name = my_np(i + 1).Name
                 my_np(i).path = my_np(i + 1).path
             Next
             kk = kk - 1
          Else
             For i = np_num To fp_np_new - 1
                 np_new(i).Name = np_new(i + 1).Name
                 np_new(i).path = np_new(i + 1).path
             Next
             fp_np_new = fp_np_new - 1
          End If
          List1.RemoveItem np_num - 1
          If List1.ListCount = 0 Then
             Text1.Text = ""
             Text2.Text = ""
          Else
             List1.ListIndex = 0
          End If
    End If
End Sub

Private Sub Command3_Click()   'Edit
    Dim i As Integer
    Dim MystrTmp As String
    
    On Error Resume Next
    If ad_fr = True Then
       ad_fr = False
       If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" Then
          np_new(fp_np_new).Name = Trim(Text1.Text)
          np_new(fp_np_new).path = Trim(Text2.Text)
          fp_np_new = fp_np_new + 1
          List1.AddItem Trim(Text1.Text), 0
       End If
    End If
    If xu = True Then
       xu = False
       If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" Then
          If xu_num > fp_np_new - 1 Then
              my_np(xu_num - (fp_np_new - 1)).Name = Trim(Text1.Text)
              my_np(xu_num - (fp_np_new - 1)).path = Trim(Text2.Text)
          Else
              np_new(fp_np_new - xu_num).Name = Trim(Text1.Text)
              np_new(fp_np_new - xu_num).path = Trim(Text2.Text)
          End If
          List1.RemoveItem xu_num - 1
          List1.AddItem Trim(Text1.Text), xu_num - 1
          xu_press = True
          Exit Sub
          
          List1.Clear
          'For i = 1 To fp_np_new
          For i = fp_np_new - 1 To 1 Step -1
              MystrTmp = np_new(i).Name
              List1.AddItem Trim(MystrTmp)
          Next
          For i = 1 To kk
              MystrTmp = my_np(i).Name
              List1.AddItem Trim(MystrTmp)
          Next
       End If
       List1.ListIndex = 0
    End If
    If List1.ListIndex = -1 Then
       Exit Sub
    End If
    xu_press = True
    xu_num = List1.ListIndex + 1
    Text1.SetFocus

End Sub

Private Sub Command4_Click()   'Exit
    Dim i As Integer
    
    On Error Resume Next
    If ad_fr = True Then
       ad_fr = False
       If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" Then
          np_new(fp_np_new).Name = Trim(Text1.Text)
          np_new(fp_np_new).path = Trim(Text2.Text)
          fp_np_new = fp_np_new + 1
          List1.AddItem Trim(Text1.Text), 0
          
       End If
    End If
    If xu = True Then
       xu = False
       If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" Then
          If xu_num > fp_np_new - 1 Then
              my_np(xu_num - (fp_np_new - 1)).Name = Trim(Text1.Text)
              my_np(xu_num - (fp_np_new - 1)).path = Trim(Text2.Text)
          Else
              np_new(fp_np_new - xu_num).Name = Trim(Text1.Text)
              np_new(fp_np_new - xu_num).path = Trim(Text2.Text)
          End If
          'List1.Clear
          'For i = 1 To fp_np_new
          '    MystrTmp = fp_np_new(i).Name
          '    List1.AddItem MystrTmp
          'Next
          'For i = 1 To kk
          '    MystrTmp = my_np(i).Name
          '    List1.AddItem MystrTmp
          'Next
       End If
       List1.ListIndex = 0
    End If
    Gsm_FileName = Gsm_Path + "\switch.dat"
    If Dir(Gsm_FileName, 0) <> "" Then
        Kill Gsm_FileName
    End If
    Open Gsm_FileName For Binary As #hFreeFile
    For i = fp_np_new - 1 To 1 Step -1
        Put #hFreeFile, , np_new(i)
    Next
    For i = 1 To kk
        Put #hFreeFile, , my_np(i)
    Next
    Close #hFreeFile
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim MystrTmp As String
    
    On Error Resume Next
    hFreeFile = FreeFile
    ad_fr = False
    xu = False
    xu_press = False
    Gsm_FileName = Gsm_Path + "\switch.dat"
    If Dir(Gsm_FileName) <> "" Then
       Open Gsm_FileName For Binary As #hFreeFile
       i = 1
       kk = 0
       If FileLen(Gsm_FileName) > 0 Then
          Do While i * 180 <= FileLen(Gsm_FileName)
             Get #hFreeFile, , my_np(i)
             MystrTmp = my_np(i).Name
             List1.AddItem Trim(MystrTmp)
             i = i + 1
             kk = kk + 1
          Loop
          List1.ListIndex = 0
       End If
       Close #hFreeFile
    End If
    fp_np_new = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Close
End Sub

Private Sub List1_Click()
    Dim i As Integer
    Dim MystrTmp As String
    
    On Error Resume Next
    If ad_fr = True Then
       ad_fr = False
       If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" Then
          np_new(fp_np_new).Name = Trim(Text1.Text)
          np_new(fp_np_new).path = Trim(Text2.Text)
          fp_np_new = fp_np_new + 1
          List1.AddItem Trim(Text1.Text), 0
       End If
    End If
    If xu = True Then
       xu = False
       If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" Then
          If xu_num > fp_np_new - 1 Then
              my_np(xu_num - (fp_np_new - 1)).Name = Trim(Text1.Text)
              my_np(xu_num - (fp_np_new - 1)).path = Trim(Text2.Text)
          Else
              np_new(fp_np_new - xu_num).Name = Trim(Text1.Text)
              np_new(fp_np_new - xu_num).path = Trim(Text2.Text)
          End If
          List1.RemoveItem xu_num - 1
          List1.AddItem Trim(Text1.Text), xu_num - 1

          'List1.Clear
          'For i = 1 To fp_np_new - 1
          'For i = fp_np_new - 1 To 1 Step -1
          '    MystrTmp = np_new(i).Name
          '    List1.AddItem Trim(MystrTmp)
          'Next
          'For i = 1 To kk
          '    MystrTmp = my_np(i).Name
          '    List1.AddItem Trim(MystrTmp)
          'Next
       End If
       List1.ListIndex = 0
    End If
    np_num = List1.ListIndex + 1
    If np_num > fp_np_new - 1 Then
        Text1.Text = Trim(my_np(np_num - (fp_np_new - 1)).Name)
        Text1.Text = Trim(Text1.Text)
        Text2.Text = Trim(my_np(np_num - (fp_np_new - 1)).path)
        Text2.Text = Trim(Text2.Text)
    Else
        Text1.Text = Trim(np_new(fp_np_new - np_num).Name)
        Text1.Text = Trim(Text1.Text)
        Text2.Text = Trim(np_new(fp_np_new - np_num).path)
        Text2.Text = Trim(Text2.Text)
    End If
    
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If xu_press = True Then
       xu = True
       xu_press = False
    End If

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If xu_press = True Then
       xu = True
       xu_press = False
    End If

End Sub

Private Sub Text2_LostFocus()
    On Error Resume Next
    'If ad_fr = True Then
    '   List1_Click
    'End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error Resume Next
    Select Case Button.Index
       Case 2
            Command1_Click
       Case 3
            Command2_Click
       Case 4
            Command3_Click
       Case 6
            Command4_Click
    End Select
End Sub
