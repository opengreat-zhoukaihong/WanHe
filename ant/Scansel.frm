VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form ScanSel 
   BackColor       =   &H00C0C0C0&
   Caption         =   "扫频数据选择 "
   ClientHeight    =   3705
   ClientLeft      =   1515
   ClientTop       =   2595
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3705
   ScaleWidth      =   8655
   Begin VB.CommandButton ScanCancel 
      Caption         =   "&C 取消"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   18
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton ScanOk 
      Caption         =   "&O 确定"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   855
      Left            =   5760
      TabIndex        =   22
      Top             =   1800
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "场强门限 "
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.TextBox Rxlev 
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   23
         Text            =   "93"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   " (- dbm)"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RxLev:"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1455
      Left            =   5760
      TabIndex        =   19
      Top             =   240
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   2566
      _StockProps     =   14
      Caption         =   "快速选择"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin Threed.SSCheck SSCheck2 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "四个BSIC全选99"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck SSCheck1 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "四个ARFCN全选"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   4260
      _StockProps     =   14
      Caption         =   "ARFCN/BSIC选择"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.ComboBox BSIC_VALUE 
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   3480
         TabIndex        =   16
         Text            =   " "
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox BSIC_VALUE 
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3480
         TabIndex        =   15
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox BSIC_VALUE 
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3480
         TabIndex        =   14
         Text            =   " "
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox ARF_VALUE 
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1320
         TabIndex        =   10
         Text            =   " "
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox ARF_VALUE 
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1320
         TabIndex        =   9
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox ARF_VALUE 
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Text            =   " "
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox BSIC_VALUE 
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3480
         TabIndex        =   4
         Text            =   " "
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox ARF_VALUE 
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Text            =   " "
         Top             =   360
         Width           =   1455
      End
      Begin Threed.SSCheck ARFSEL 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "ARFCN："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck ARFSEL 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "ARFCN："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck ARFSEL 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "ARFCN："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck ARFSEL 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "ARFCN："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BSIC："
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   13
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BSIC："
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BSIC："
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BSIC："
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "ScanSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ARF_VALUE_Click(Index As Integer)
    Dim row_num, i, BSIC, col_no As Integer
    Dim str As String
    
    On Error Resume Next
    str = ""
    msg = "TableInfo(""" & tblname & """, 8)"
    row_num = Val(MapInfo.eval(msg))
    
    msg = "Columninfo(""" & tblname & """,""" & ARF_VALUE(Index).Text & """, 2)"
    col_no = Val(MapInfo.eval(msg)) + 1
    
    MapInfo.Do "fetch  first from " & tblname
    
    Dim j, M, k As Integer
    Dim bs(10)
    j = 0
    M = 1
    While i < row_num
        msg = tblname + "." + "COL" & col_no
        msg = MapInfo.eval(msg)
        If str <> msg Then
           For k = 1 To 10
               If msg = bs(k) Then
                  GoTo AA
               End If
           Next k
           bs(M) = msg
           M = M + 1
           str = msg
        End If
AA:
        i = i + 1
        On Error Resume Next
         MapInfo.Do "fetch  next from " & tblname
     Wend
    
    For k = 1 To 10
       If bs(k) <> "" Then
          BSIC_VALUE(Index).AddItem bs(k)
       End If
    Next k
    
    BSIC_VALUE(Index).Text = str
End Sub

Private Sub ARFSEL_Click(Index As Integer, Value As Integer)
    On Error Resume Next
   If Value = -1 Then
    ARF_VALUE(Index).Enabled = 1
    BSIC_VALUE(Index).Enabled = 1
  Else
    ARF_VALUE(Index).Enabled = 0
    BSIC_VALUE(Index).Enabled = 0
  End If
End Sub

Private Sub Form_Load()
    Dim col_num, i As Integer
    Dim str As String
    On Error Resume Next
    
    msg = "TableInfo(""" & tblname & """, 4)"
    col_num = Val(MapInfo.eval(msg))
    For i = 4 To col_num Step 2
        msg = "Columninfo(""" & tblname & """,""COL" & i & """, 1)"
        On Error Resume Next
        str = MapInfo.eval(msg)
        
        ARF_VALUE(0).AddItem str
        ARF_VALUE(1).AddItem str
        ARF_VALUE(2).AddItem str
        ARF_VALUE(3).AddItem str
    Next i
    ARF_VALUE(0).Text = str
    ARF_VALUE(1).Text = str
    ARF_VALUE(2).Text = str
    ARF_VALUE(3).Text = str
    
    ARF_VALUE(1).Enabled = 0
    ARF_VALUE(2).Enabled = 0
    ARF_VALUE(3).Enabled = 0
    
    BSIC_VALUE(1).Enabled = 0
    BSIC_VALUE(2).Enabled = 0
    BSIC_VALUE(3).Enabled = 0
    
End Sub

Private Sub ScanCancel_Click()
    Unload Me
End Sub

Private Sub SCANHELP_Click()

End Sub

Private Sub SCANOK_Click()
  Dim col_no, Rxlev_value, BSIC As Integer
  Dim Name As String
   
    On Error Resume Next
  ScanSel.Hide
  Select Case Menu_Flag
   Case 912, 913
      If ARF_VALUE(0).Enabled = True And BSIC_VALUE(0).Enabled = True Then
         msg = "Columninfo(""" & tblname & """,""" & ARF_VALUE(0).Text & """, 2)"
         col_no = Val(MapInfo.eval(msg))
         Rxlev_value = Val(Rxlev.Text)
         BSIC = Val(BSIC_VALUE(0).Text)
         
         If Menu_Flag = 912 Then
            Name = "SCAN_GOOD1"
            msg = "select  *  from " & tblname & "  where  COL" & col_no & " <= " & Rxlev_value & " "
         Else
            Name = "SCAN_POOR1"
            msg = "select  *  from " & tblname & "  where  COL" & col_no & " > " & Rxlev_value & " "
         End If
         msg = msg + "AND COL" & col_no + 1 & " = " & BSIC & "  "
         msg = msg + "  into " & Name
         MapInfo.Do msg
         
         msg = "Add Map Auto Layer " + Chr(34) + Name + Chr(34)
         MapInfo.Do msg

         msg = "shade window   Frontwindow()  " & Name & " with  COL" & col_no & "   ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  20: 35 Symbol (39,1052688,8,""MapInfo Cartographic"",0,0) ,35: 50 Symbol (39,16754768,8,""MapInfo Cartographic"",0,0) ,50: 65 Symbol (39,128,8,""MapInfo Cartographic"",0,0) ,65: 70 Symbol (39,32768,8,""MapInfo Cartographic"",0,0) ,70: 75 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,75: 80 Symbol (39,8388608,8,""MapInfo Cartographic"",0,0) ,80: 85 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,85: 90 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,90: 95 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,95: 100 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,100: 105 Symbol (39,16776960,8,""MapInfo Cartographic"",0,0) ,105: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
         MapInfo.Do msg

         If legendid = 0 Then
                MapInfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                MapInfo.Do "Create Legend From Window  Frontwindow()"
                legendid = MapInfo.eval("windowinfo(1009,12)")
         End If
         If Menu_Flag = 912 Then
            msg = " Title " + Chr(34) + "频率 " + ARF_VALUE(0).Text + " 优化图 (-dBm)  " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34)
         Else
            msg = " Title " + Chr(34) + "频率 " + ARF_VALUE(0).Text + " 弱区图 (-dBm) " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34)
         End If
         MapInfo.Do "set legend window FrontWindow()  Layer prev " & msg
      End If
      
      If ARF_VALUE(1).Enabled = True And BSIC_VALUE(1).Enabled = True Then
        If Map_No > 0 And Map_No < 4 Then
         SUB_24.Enabled = 0
         ReDim ViceMap(Map_No)
         ViceMap(Map_No).Caption = "副本视图：" + MapForm.Caption
         MapInfo.Do "Set Next Document Parent " & ViceMap(Map_No).hwnd & " Style 1"

         MapInfo.Do "Run Command WindowInfo(" & mapid & ",15)"
         Map_No = Map_No + 1
         
         ViceMap(Map_No - 1).SetFocus
         
         msg = "Columninfo(""" & tblname & """,""" & ARF_VALUE(1).Text & """, 2)"
         col_no = Val(MapInfo.eval(msg))
         Rxlev_value = Val(Rxlev.Text)
         BSIC = Val(BSIC_VALUE(1).Text)
         
         If Menu_Flag = 912 Then
            Name = "SCAN_GOOD2"
            msg = "select  *  from " & tblname & "  where  COL" & col_no & " <= " & Rxlev_value & " "
         Else
            Name = "SCAN_POOR2"
            msg = "select  *  from " & tblname & "  where  COL" & col_no & " > " & Rxlev_value & " "
         End If
         msg = msg + "AND COL" & col_no + 1 & " = " & BSIC & "  "
         msg = msg + "  into " & Name
         MapInfo.Do msg
         
         msg = "Add Map Auto Layer " + Chr(34) + Name + Chr(34)
         MapInfo.Do msg

         msg = "shade window   Frontwindow()  " & Name & " with  COL" & col_no & "  ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  20: 35 Symbol (39,1052688,8,""MapInfo Cartographic"",0,0) ,35: 50 Symbol (39,16754768,8,""MapInfo Cartographic"",0,0) ,50: 65 Symbol (39,128,8,""MapInfo Cartographic"",0,0) ,65: 70 Symbol (39,32768,8,""MapInfo Cartographic"",0,0) ,70: 75 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,75: 80 Symbol (39,8388608,8,""MapInfo Cartographic"",0,0) ,80: 85 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,85: 90 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,90: 95 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,95: 100 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,100: 105 Symbol (39,16776960,8,""MapInfo Cartographic"",0,0) ,105: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
         MapInfo.Do msg

         If legendid = 0 Then
                MapInfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                MapInfo.Do "Create Legend From Window  Frontwindow()"
                legendid = MapInfo.eval("windowinfo(1009,12)")
         End If
         If Menu_Flag = 912 Then
            msg = " Title " + Chr(34) + "频率 " + ARF_VALUE(1).Text + " 优化图 (-dBm) " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34)
         Else
            msg = " Title " + Chr(34) + "频率 " + ARF_VALUE(1).Text + " 弱区图 (-dBm) " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34)
         End If
         MapInfo.Do "set legend window FrontWindow()  Layer prev " & msg
        End If
      End If
       
       If ARF_VALUE(2).Enabled = True And BSIC_VALUE(2).Enabled = True Then
        If Map_No > 0 And Map_No < 4 Then
         SUB_24.Enabled = 0
         ReDim ViceMap(Map_No)
         ViceMap(Map_No).Caption = "副本视图：" + MapForm.Caption
         MapInfo.Do "Set Next Document Parent " & ViceMap(Map_No).hwnd & " Style 1"

         MapInfo.Do "Run Command WindowInfo(" & mapid & ",15)"
         Map_No = Map_No + 1
         
         ViceMap(Map_No - 1).SetFocus
         
         msg = "Columninfo(""" & tblname & """,""" & ARF_VALUE(2).Text & """, 2)"
         col_no = Val(MapInfo.eval(msg))
         Rxlev_value = Val(Rxlev.Text)
         BSIC = Val(BSIC_VALUE(2).Text)
         
         If Menu_Flag = 912 Then
            Name = "SCAN_GOOD3"
            msg = "select  *  from " & tblname & "  where  COL" & col_no & " <= " & Rxlev_value & " "
         Else
            Name = "SCAN_POOR3"
            msg = "select  *  from " & tblname & "  where  COL" & col_no & " > " & Rxlev_value & " "
         End If
         msg = msg + "AND COL" & col_no + 1 & " = " & BSIC & "  "
         msg = msg + "  into " & Name
         MapInfo.Do msg
         
         msg = "Add Map Auto Layer " + Chr(34) + Name + Chr(34)
         MapInfo.Do msg

         msg = "shade window   Frontwindow()  " & Name & " with  COL" & col_no & " ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  20: 35 Symbol (39,1052688,8,""MapInfo Cartographic"",0,0) ,35: 50 Symbol (39,16754768,8,""MapInfo Cartographic"",0,0) ,50: 65 Symbol (39,128,8,""MapInfo Cartographic"",0,0) ,65: 70 Symbol (39,32768,8,""MapInfo Cartographic"",0,0) ,70: 75 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,75: 80 Symbol (39,8388608,8,""MapInfo Cartographic"",0,0) ,80: 85 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,85: 90 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,90: 95 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,95: 100 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,100: 105 Symbol (39,16776960,8,""MapInfo Cartographic"",0,0) ,105: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
         MapInfo.Do msg

         If legendid = 0 Then
                MapInfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                MapInfo.Do "Create Legend From Window  Frontwindow()"
                legendid = MapInfo.eval("windowinfo(1009,12)")
         End If
         If Menu_Flag = 912 Then
            msg = " Title " + Chr(34) + "频率 " + ARF_VALUE(1).Text + " 优化图 (-dBm) " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34)
         Else
            msg = " Title " + Chr(34) + "频率 " + ARF_VALUE(1).Text + " 弱区图 (-dBm) " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34)
         End If
         MapInfo.Do "set legend window FrontWindow()  Layer prev " & msg
        End If
     Else
        MDIMain.Arrange 2   'TITLE
      End If

       If ARF_VALUE(3).Enabled = True And BSIC_VALUE(3).Enabled = True Then
        If Map_No > 0 And Map_No < 4 Then
         SUB_24.Enabled = 0
         ReDim ViceMap(Map_No)
         ViceMap(Map_No).Caption = "副本视图：" + MapForm.Caption
         MapInfo.Do "Set Next Document Parent " & ViceMap(Map_No).hwnd & " Style 1"

         MapInfo.Do "Run Command WindowInfo(" & mapid & ",15)"
         Map_No = Map_No + 1
         
         ViceMap(Map_No - 1).SetFocus
         msg = "Columninfo(""" & tblname & """,""" & ARF_VALUE(3).Text & """, 2)"
         col_no = Val(MapInfo.eval(msg))
         Rxlev_value = Val(Rxlev.Text)
         BSIC = Val(BSIC_VALUE(3).Text)
         
         If Menu_Flag = 912 Then
            Name = "SCAN_GOOD4"
            msg = "select  *  from " & tblname & "  where  COL" & col_no & " <= " & Rxlev_value & " "
         Else
            Name = "SCAN_POOR4"
            msg = "select  *  from " & tblname & "  where  COL" & col_no & " > " & Rxlev_value & " "
         End If
         msg = msg + "AND COL" & col_no + 1 & " = " & BSIC & "  "
         msg = msg + "  into " & Name
         MapInfo.Do msg
         
         msg = "Add Map Auto Layer " + Chr(34) + Name + Chr(34)
         MapInfo.Do msg

         msg = "shade window   Frontwindow()  " & Name & " with  COL" & col_no & "  ignore 0 ranges apply all use all Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)  20: 35 Symbol (39,1052688,8,""MapInfo Cartographic"",0,0) ,35: 50 Symbol (39,16754768,8,""MapInfo Cartographic"",0,0) ,50: 65 Symbol (39,128,8,""MapInfo Cartographic"",0,0) ,65: 70 Symbol (39,32768,8,""MapInfo Cartographic"",0,0) ,70: 75 Symbol (39,8421376,8,""MapInfo Cartographic"",0,0) ,75: 80 Symbol (39,8388608,8,""MapInfo Cartographic"",0,0) ,80: 85 Symbol (39,16711935,8,""MapInfo Cartographic"",0,0) ,85: 90 Symbol (39,255,8,""MapInfo Cartographic"",0,0) ,90: 95 Symbol (39,65535,8,""MapInfo Cartographic"",0,0) ,95: 100 Symbol (39,65280,8,""MapInfo Cartographic"",0,0) ,100: 105 Symbol (39,16776960,8,""MapInfo Cartographic"",0,0) ,105: 110 Symbol (39,16711680,8,""MapInfo Cartographic"",0,0)"
         MapInfo.Do msg

         If legendid = 0 Then
                MapInfo.Do "Set Next Document Parent " & MDIMain.hwnd & " Style 0"
                MapInfo.Do "Create Legend From Window  Frontwindow()"
                legendid = MapInfo.eval("windowinfo(1009,12)")
         End If
         If Menu_Flag = 912 Then
            msg = " Title " + Chr(34) + "频率 " + ARF_VALUE(1).Text + " 优化图 (-dBm) " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34)
         Else
            msg = " Title " + Chr(34) + "频率 " + ARF_VALUE(1).Text + " 弱区图 (-dBm) " + tblname + Chr(34) + " Subtitle" + Chr(34) + USERNAME + Chr(34)
         End If
         MapInfo.Do "set legend window FrontWindow()  Layer prev " & msg
        End If
        MDIMain.Arrange 2   'TITLE
      Else
        MDIMain.Arrange 2   'TITLE
      End If
      
   Case 914
   
   Case 915
  
  End Select

    Unload Me
End Sub

Private Sub SSCheck1_Click(Value As Integer)
On Error Resume Next
If Value = -1 Then
    ARFSEL(0).Value = -1
    ARFSEL(1).Value = -1
    ARFSEL(2).Value = -1
    ARFSEL(3).Value = -1
    
    ARF_VALUE(0).Enabled = 1
    ARF_VALUE(1).Enabled = 1
    ARF_VALUE(2).Enabled = 1
    ARF_VALUE(3).Enabled = 1
    
    BSIC_VALUE(0).Enabled = 1
    BSIC_VALUE(1).Enabled = 1
    BSIC_VALUE(2).Enabled = 1
    BSIC_VALUE(3).Enabled = 1
  Else
    ARFSEL(0).Value = 0
    ARFSEL(1).Value = 0
    ARFSEL(2).Value = 0
    ARFSEL(3).Value = 0
    
    ARF_VALUE(0).Enabled = 0
    ARF_VALUE(1).Enabled = 0
    ARF_VALUE(2).Enabled = 0
    ARF_VALUE(3).Enabled = 0
    
    BSIC_VALUE(0).Enabled = 0
    BSIC_VALUE(1).Enabled = 0
    BSIC_VALUE(2).Enabled = 0
    BSIC_VALUE(3).Enabled = 0
  End If
End Sub

Private Sub SSCheck2_Click(Value As Integer)
On Error Resume Next
If Value = -1 Then
    BSIC_VALUE(0).Text = 99
    BSIC_VALUE(1).Text = 99
    BSIC_VALUE(2).Text = 99
    BSIC_VALUE(3).Text = 99
 Else
   BSIC_VALUE(0).Text = ""
   BSIC_VALUE(1).Text = ""
   BSIC_VALUE(2).Text = ""
   BSIC_VALUE(3).Text = ""
 End If
End Sub

