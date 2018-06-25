VERSION 5.00
Begin VB.Form Center 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "基站选择"
   ClientHeight    =   1455
   ClientLeft      =   3450
   ClientTop       =   3225
   ClientWidth     =   3180
   Icon            =   "Sel_cent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1455
   ScaleWidth      =   3180
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      DataField       =   " "
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton SBSCancel 
      Caption         =   "&C 取消"
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
      Left            =   1620
      TabIndex        =   2
      Top             =   1020
      Width           =   1080
   End
   Begin VB.CommandButton SBSOK 
      Caption         =   "&O 确认"
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
      Left            =   405
      TabIndex        =   1
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "基站选择："
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
      Index           =   0
      Left            =   375
      TabIndex        =   3
      Top             =   405
      Width           =   900
   End
End
Attribute VB_Name = "Center"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Dim connect As String
  On Error Resume Next
  i = 0
  row = Val(mapinfo.eval("tableinfo(base,8)"))
  mapinfo.do "fetch First from base"
  While i < row
       Combo1.AddItem mapinfo.eval("base.bs_NAME")
       mapinfo.do "fetch next from base"
       i = i + 1
  Wend

   mapinfo.do "fetch First from base"
   Combo1.Text = mapinfo.eval("base.bs_NAME")   'win95
End Sub

Private Sub SBSCancel_Click()
   On Error Resume Next
  Unload Me
End Sub

Private Sub SBSOK_Click()
 Dim X, Y As Double
  
 Dim ci(4), bs_name, bs_overlay  As String
 Dim i, row, BSIC(3), ARFCN(3), Lac As Integer
 Dim Rxlev1, Rxqual1  As Integer
 Dim finds As Integer
  
 On Error Resume Next
 Screen.MousePointer = 11
 If Combo1.Text <> "" Then
    i = 0
    row = Val(mapinfo.eval("tableinfo(base,8)"))
    mapinfo.do "fetch First from base"
    msg = mapinfo.eval("base.bs_NAME")
    finds = InStr(msg, Chr(0))
    If finds > 0 Then
       msg = Trim(Left(msg, finds - 1))
    End If
    While i < row And msg <> Combo1.Text
             mapinfo.do "fetch next from base"
             msg = mapinfo.eval("base.bs_NAME")
             finds = InStr(msg, Chr(0))
             If finds > 0 Then
                msg = Trim(Left(msg, finds - 1))
             End If
             i = i + 1
    Wend
  
 If Menu_Flag = 151 Then
    i = 0
    row = Val(mapinfo.eval("tableinfo(base,8)"))
    mapinfo.do "fetch First from base"
    msg = mapinfo.eval("base.bs_NAME")
    finds = InStr(msg, Chr(0))
    If finds > 0 Then
       msg = Trim(Left(msg, finds - 1))
    End If
    While i < row And msg <> Combo1.Text
             mapinfo.do "fetch next from base"
             msg = mapinfo.eval("base.bs_NAME")
             finds = InStr(msg, Chr(0))
             If finds > 0 Then
                msg = Trim(Left(msg, finds - 1))
             End If
             i = i + 1
    Wend

   bs_name = Combo1.Text

   Center.Hide
   mapinfo.do "x1= base.lon"
   mapinfo.do "y1= base.lat"
        
   msg = "set map Center(x1,y1) Smart redraw zoom 4.5 units " + Chr(34) + "km" + Chr(34)
   mapinfo.do msg
End If

 If Menu_Flag = 4600 Then
    mapinfo.do "open table " + Chr(34) + Gsm_Path + "\map\base_add.tab" + Chr(34)
    i = 0
    row = Val(mapinfo.eval("tableinfo(base_add,8)"))
    mapinfo.do "fetch First from base_add"
    msg = mapinfo.eval("base_add.bs_NAME")
    finds = InStr(msg, Chr(0))
    If finds > 0 Then
       msg = Trim(Left(msg, finds - 1))
    End If
    While i < row And msg <> Combo1.Text
             mapinfo.do "fetch next from base_add"
             msg = mapinfo.eval("base_add.bs_NAME")
             finds = InStr(msg, Chr(0))
             If finds > 0 Then
                msg = Trim(Left(msg, finds - 1))
             End If
             i = i + 1
    Wend
    If i < row Then
'       msg1 = Mid$(mapinfo.eval("base_add.address"), 1, 5)
       msg1 = mapinfo.eval("base_add.address")
    Else
       msg1 = "无此站站址!"
    End If
    msg = msg + "站址"
    Unload Me
    If Trim(msg1) = "" Then
       msg1 = "无此站站址!"
    End If
    MsgBox msg1, 64, msg
    mapinfo.do "close table  base_add "
End If

End If
 Screen.MousePointer = 0
 Unload Me
End Sub


