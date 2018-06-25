VERSION 5.00
Begin VB.Form MessageReplay 
   BackColor       =   &H8000000B&
   Caption         =   "信令回放"
   ClientHeight    =   6795
   ClientLeft      =   5580
   ClientTop       =   1440
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MessageReplay.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   5865
   Begin VB.PictureBox My_Picture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      Height          =   4875
      Left            =   195
      ScaleHeight     =   4815
      ScaleMode       =   0  'User
      ScaleWidth      =   5402.986
      TabIndex        =   24
      Top             =   1440
      Width           =   5490
      Begin VB.Line Line2 
         BorderColor     =   &H80000006&
         X1              =   2373.135
         X2              =   2373.135
         Y1              =   0
         Y2              =   5340
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   165
      ScaleHeight     =   345
      ScaleWidth      =   2265
      TabIndex        =   15
      Top             =   1110
      Width           =   2265
      Begin VB.Line Line1 
         Index           =   7
         X1              =   2160
         X2              =   2160
         Y1              =   270
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   1860
         X2              =   1860
         Y1              =   270
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   1560
         X2              =   1560
         Y1              =   270
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   1260
         X2              =   1260
         Y1              =   270
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   960
         X2              =   960
         Y1              =   270
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   660
         X2              =   660
         Y1              =   270
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   360
         X2              =   360
         Y1              =   270
         Y2              =   360
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   60
         X2              =   60
         Y1              =   270
         Y2              =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "7"
         Height          =   180
         Index           =   0
         Left            =   15
         TabIndex        =   23
         Top             =   75
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   180
         Index           =   1
         Left            =   330
         TabIndex        =   22
         Top             =   75
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   2
         Left            =   630
         TabIndex        =   21
         Top             =   75
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   3
         Left            =   930
         TabIndex        =   20
         Top             =   75
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   4
         Left            =   1230
         TabIndex        =   19
         Top             =   75
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   5
         Left            =   1530
         TabIndex        =   18
         Top             =   75
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   6
         Left            =   1845
         TabIndex        =   17
         Top             =   75
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   180
         Index           =   7
         Left            =   2130
         TabIndex        =   16
         Top             =   75
         Width           =   90
      End
   End
   Begin VB.CommandButton Play 
      Caption         =   "Play"
      Height          =   300
      Left            =   225
      TabIndex        =   14
      Top             =   6435
      Width           =   1080
   End
   Begin VB.CommandButton Pause 
      Caption         =   "Pause"
      Height          =   300
      Left            =   1305
      TabIndex        =   13
      Top             =   6435
      Width           =   1080
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      Height          =   300
      Left            =   2385
      TabIndex        =   12
      Top             =   6435
      Width           =   1080
   End
   Begin VB.CommandButton Step 
      Caption         =   "Step"
      Height          =   300
      Left            =   3465
      TabIndex        =   11
      Top             =   6435
      Width           =   1080
   End
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   300
      Left            =   4545
      TabIndex        =   10
      Top             =   6435
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "图例说明"
      Height          =   1035
      Left            =   2655
      TabIndex        =   3
      Top             =   30
      Width           =   2865
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Both"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   22
         Left            =   1860
         TabIndex        =   9
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "DownLink"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   21
         Left            =   1860
         TabIndex        =   8
         Top             =   510
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "UpLink"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   20
         Left            =   1845
         TabIndex        =   7
         Top             =   255
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CM"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   19
         Left            =   765
         TabIndex        =   6
         Top             =   750
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "MM"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   18
         Left            =   765
         TabIndex        =   5
         Top             =   510
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RR"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   17
         Left            =   750
         TabIndex        =   4
         Top             =   270
         Width           =   180
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         Index           =   5
         X1              =   1470
         X2              =   1710
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   2
         Index           =   4
         X1              =   1470
         X2              =   1710
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Index           =   3
         X1              =   1470
         X2              =   1710
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         Index           =   2
         X1              =   375
         X2              =   615
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Index           =   1
         X1              =   375
         X2              =   615
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   375
         X2              =   615
         Y1              =   375
         Y2              =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   420
      TabIndex        =   0
      Top             =   405
      Width           =   2040
      Begin VB.OptionButton Option1 
         Caption         =   "Full"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   285
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sub"
         Height          =   240
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   285
         Width           =   570
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   -150
      Top             =   1590
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "UnKown"
      Height          =   180
      Left            =   1290
      TabIndex        =   28
      Top             =   165
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "当前小区:"
      Height          =   180
      Left            =   405
      TabIndex        =   27
      Top             =   165
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "RxQual"
      Height          =   180
      Left            =   2535
      TabIndex        =   26
      Top             =   1185
      Width           =   540
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "信令信息"
      Height          =   180
      Index           =   10
      Left            =   3720
      TabIndex        =   25
      Top             =   1185
      Width           =   720
   End
End
Attribute VB_Name = "MessageReplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentRow As Integer, TableRows As Integer, StartRow As Integer
Dim MyTableName As String
Dim PlayFlag As Boolean
Dim MessageString(1 To 20) As String
Dim MessageColor(1 To 20) As Long, LinkColor(1 To 20) As Long
Dim RxQualBuffer(1 To 20) As Integer
Dim NumCount As Integer
Dim LabelColor As Long, LineColor As Long
Dim CiString As String, OldCiString As String, MyCellName As String
Dim StartCi As String, StartName As String

Private Sub Close_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim WinId As Variant
    
    On Error Resume Next
    
    MyTableName = mapinfo.eval("selectionInfo(1)")
    mapinfo.do "set map redraw off"
    mapinfo.do "Set Map Layer 0 Editable On  "
    mapinfo.do "set map redraw on"
    mapinfo.do " reload Custom Symbols From " + Chr(34) + Gsm_Path + "\mysymb" + Chr(34)
    
    CurrentRow = mapinfo.eval("selectioninfo(3)")
    For i = 1 To mapinfo.eval("NumWindows()")
        If mapinfo.eval("windowinfo(" & mapinfo.eval("windowid(" & i & ")") & ",3)") = 1 Then
           WinId = mapinfo.eval("windowid(" & i & ")")
           If WinId = mapinfo.eval("frontwindow()") Then
              Exit For
           End If
        End If
    Next
    
    If CurrentRow <> 0 Then
       CurrentRow = Val(mapinfo.eval("searchpoint(" & WinId & ",selection.lon,selection.lat)"))
       CurrentRow = Val(mapinfo.eval("SearchInfo(1, 2)"))
       mapinfo.do "fetch rec " & CurrentRow & " from " & MyTableName
    End If
    TableRows = mapinfo.eval("tableinfo(" & MyTableName & ",8)")
    CiString = mapinfo.eval(MyTableName & ".ci_serv")
    GetCellName
    Label5.Caption = MyCellName
    OldCiString = CiString
    StartCi = CiString
    StartName = MyCellName
    StartRow = CurrentRow
    My_Picture.Cls
    PlayFlag = False
    NumCount = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mapinfo.do "set map redraw off"
    mapinfo.do "delete  from cosmetic1 "
    mapinfo.do "Set Map Layer 0 Editable Off  "
    mapinfo.do "set map redraw on"

End Sub

Private Sub Pause_Click()
    On Error Resume Next
    PlayFlag = False
End Sub

Private Sub Play_Click()
    On Error Resume Next
    
    PlayFlag = True
    
End Sub

Sub PlayProcess()
    Dim i As Integer
    
    On Error Resume Next
    
    If CurrentRow >= TableRows Then
       Exit Sub
    End If
    If NumCount >= 21 Then
       My_Picture.Cls
       For i = 2 To 20
           My_Picture.CurrentX = 2600
           My_Picture.CurrentY = 30 + (i - 2) * 240
           My_Picture.ForeColor = MessageColor(i)
           My_Picture.Print MessageString(i)
           My_Picture.Line ((7 - RxQualBuffer(i)) * 300, (i - 2) * 240)-((7 - RxQualBuffer(i)) * 300, (i - 1) * 240), &HFF&
           If LinkColor(i) <> 0 Then
              My_Picture.Line (0, (i - 1) * 240 - 120)-(2373, (i - 1) * 240 - 120), LinkColor(i)
           End If
           If i > 2 Then
              My_Picture.Line ((7 - RxQualBuffer(i - 1)) * 300, (i - 2) * 240)-((7 - RxQualBuffer(i)) * 300, (i - 2) * 240), &HFF&
           End If
           MessageString(i - 1) = MessageString(i)
           MessageColor(i - 1) = MessageColor(i)
           RxQualBuffer(i - 1) = RxQualBuffer(i)
           LinkColor(i - 1) = LinkColor(i)
       Next
       MessageString(20) = mapinfo.eval(MyTableName & ".message")
       MessageString(20) = Trim(UCase(MessageString(20)))
       Call GetColor(MessageString(20))
       MessageColor(20) = LabelColor
       LinkColor(20) = LineColor
       If Option1(0).Value = True Then
          RxQualBuffer(20) = Val(mapinfo.eval(MyTableName & ".rxqual_f"))
       Else
          RxQualBuffer(20) = Val(mapinfo.eval(MyTableName & ".rxqual_s"))
       End If
       My_Picture.CurrentX = 2600
       My_Picture.CurrentY = 30 + (20 - 1) * 240
       My_Picture.ForeColor = LabelColor
       My_Picture.Print MessageString(20)
       My_Picture.Line ((7 - RxQualBuffer(20)) * 300, 19 * 240)-((7 - RxQualBuffer(20)) * 300, 20 * 240), &HFF&
       My_Picture.Line ((7 - RxQualBuffer(19)) * 300, 19 * 240)-((7 - RxQualBuffer(20)) * 300, 19 * 240), &HFF&
       If LinkColor(20) <> 0 Then
          My_Picture.Line (0, 20 * 240 - 120)-(2373, 20 * 240 - 120), LinkColor(20)
       End If
       My_Picture.Line (2373, 0)-(2373, 4815), &H0&
       
       
       If MessageString(20) = "HANDOVER COMPLETE" Then
          mapinfo.do "Set Style Symbol MakeSymbol(37,255,24)"
       Else
          If MessageString(20) = "SETUP" Then
             mapinfo.do "Set Style Symbol MakeSymbol(47,65535,24)"
          Else
             If MessageString(20) = "RELEASE COMPLETE" Then
                mapinfo.do "Set Style Symbol MakeSymbol(48,65280,24)"
             Else
                mapinfo.do "Set Style Symbol MakeSymbol(33,255,4)"
             End If
          End If
       End If
       mapinfo.do "Create Point(" & MyTableName & ".lon ," & MyTableName & ".lat)"
    Else
       MessageString(NumCount) = mapinfo.eval(MyTableName & ".message")
       MessageString(NumCount) = Trim(UCase(MessageString(NumCount)))
       Call GetColor(MessageString(NumCount))
       MessageColor(NumCount) = LabelColor
       LinkColor(NumCount) = LineColor
       My_Picture.CurrentX = 2600
       My_Picture.CurrentY = 30 + (NumCount - 1) * 240
       My_Picture.ForeColor = LabelColor
       My_Picture.Print MessageString(NumCount)
       If Option1(0).Value = True Then
          RxQualBuffer(NumCount) = Val(mapinfo.eval(MyTableName & ".rxqual_f"))
       Else
          RxQualBuffer(NumCount) = Val(mapinfo.eval(MyTableName & ".rxqual_s"))
       End If
       My_Picture.Line ((7 - RxQualBuffer(NumCount)) * 300, NumCount * 240 - 240)-((7 - RxQualBuffer(NumCount)) * 300, NumCount * 240), &HFF&
       If LinkColor(NumCount) <> 0 Then
          'My_Picture.Line (0, NumCount * 240)-(2373, NumCount * 240), LinkColor(NumCount)
          My_Picture.Line (0, NumCount * 240 - 120)-(2373, NumCount * 240 - 120), LinkColor(NumCount)
       End If
       If NumCount > 1 Then
          My_Picture.Line ((7 - RxQualBuffer(NumCount - 1)) * 300, NumCount * 240 - 240)-((7 - RxQualBuffer(NumCount)) * 300, NumCount * 240 - 240), &HFF&
       End If
       My_Picture.Line (2373, 0)-(2373, 4815), &H0&
       
       
       If MessageString(NumCount) = "HANDOVER COMPLETE" Then
          mapinfo.do "Set Style Symbol MakeSymbol(37,255,24)"
       Else
          If MessageString(NumCount) = "SETUP" Then
             mapinfo.do "Set Style Symbol MakeSymbol(47,65535,24)"
          Else
             If MessageString(NumCount) = "RELEASE COMPLETE" Then
                mapinfo.do "Set Style Symbol MakeSymbol(48,65280,24)"
             Else
                mapinfo.do "Set Style Symbol MakeSymbol(33,255,4)"
             End If
          End If
       End If
       mapinfo.do "Create Point(" & MyTableName & ".lon ," & MyTableName & ".lat)"
       NumCount = NumCount + 1
    End If
    CiString = mapinfo.eval(MyTableName & ".ci_serv")
    If Trim(CiString) <> Trim(OldCiString) Then
       GetCellName
       Label5.Caption = MyCellName
       OldCiString = CiString
    End If
    mapinfo.do "fetch next from " & MyTableName
    CurrentRow = CurrentRow + 1

End Sub

Sub GetColor(MessageStr As String)
    On Error Resume Next
    Select Case MessageStr
       Case "ADDITIONAL ASSIGNMENT", "IMMEDIATE ASSIGNMENT", "IMMEDIATE ASSIGNMENT EXTENDED", "IMMEDIAT ASSIGNMENT REJECT", "CIPHERING MODE COMMAND", "CIPHERING MODE COMPLETE", "ASSIGNMENT COMMAND", "ASSIGNMENT COMPLETE", "ASSIGNMENT FAILURE", "HANDOVER COMMAND", "HANDOVER COMPLETE", "HANDOVER FAILURE", "CHANNEL RELEASE", "PARTIAL RELEASE", "PARTIAL RELEASE COMPLETE", "PAGING REQUEST TYPE 1", "PAGING REQUEST TYPE 2", "PAGING REQUEST TYPE 3", "PAGING RESPONSE", "SYSTEM INFORMATION TYPE 1", "SYSTEM INFORMATION TYPE 2", "SYSTEM INFORMATION TYPE 3", "SYSTEM INFORMATION TYPE 4", "SYSTEM INFORMATION TYPE 5", "SYSTEM INFORMATION TYPE 6", "CHANNEL MODE MODIFY", "CHANNEL MODE MODIFY ACK", "CLASSMARK CHANGE", "FREQUENCY REDEFINITION", "MEAUREMENT REPORT", "RR STATUS"
            'LabelColor = &HFF&
            LabelColor = &HFF0000
       Case "IMSI DETACH INDICATION", "LOCATION UPDATING ACCEPT", "LOCATION UPDATING REJECT", "LOCATION UPDATING REQUEST", "AUTHENTICATION REJECT", "AUTHENTICATION REQUEST", "AUTHENTICATION RESPONSE", "IDENTITY REQUEST", "IDENTITY RESPONSE", "TMSI REALLOCATION COMMAND", "TMSI REALLOCATION COMPLETE", "CM SERVICE ACCEPT", "CM SERVICE REJECT", "CM SERVICE REQUEST", "CM REESTABLISHMENT REQUEST", "MM STATUS"
            'LabelColor = &HFF0000
            LabelColor = &H8000&
       Case "ALERTING", "CALL CONFIRMED", "CALL PROCEEDING", "CONNECT", "CONNECT ACKNOWLEDGE", "EMERGENCY SETUP", "PROGRESS", "SETUP", "MODIFY", "MODIFY COMPLETE", "MODIFY REJECT", "USER INFORMATION", "DISCONNECT", "RELEASE", "RELEASE COMPLETE", "CONGESTION CONTROL", "NOTIFY", "START DTMF", "START DTMF ACKNOWLEGDE", "START DTMF REJECT", "STATUS", "STATUS ENQUIRY", "STOP DTMF", "STOP DTMF ACKNOWLEDGE"
            LabelColor = &HFFFF&
       Case Else
            LineColor = &H0&
    End Select
    Select Case MessageStr
       Case "STOP DTMF", "START DTMF", "EMERGENCY SETUP", "CALL CONFIRMED", "CM REESTABLISHMENT REQUEST", "CM SERVICE REQUEST", "TMSI REALLOCATION COMPLETE", "IDENTITY RESPONSE", "AUTHENTICATION RESPONSE", "LOCATION UPDATING REQUEST", "IMSI DETACH INDICATION", "MEAUREMENT REPORT", "CLASSMARK CHANGE", "CHANNEL MODE MODIFY ACK", "PAGING RESPONSE", "PARTIAL RELEASE COMPLETE", "HANDOVER FAILURE", "HANDOVER COMPLETE", "ASSIGNMENT FAILURE", "ASSIGNMENT COMPLETE", "CIPHERING MODE COMPLETE"
            LineColor = &HFF00&
       Case "ADDITIONAL ASSIGNMENT", "IMMEDIATE ASSIGNMENT", "IMMEDIATE ASSIGNMENT EXTENDED", "IMMEDIAT ASSIGNMENT REJECT", "CIPHERING MODE COMMAND", "ASSIGNMENT COMMAND", "HANDOVER COMMAND", "CHANNEL RELEASE", "PARTIAL RELEASE", "PAGING REQUEST TYPE 1", "PAGING REQUEST TYPE 2", "PAGING REQUEST TYPE 3", "SYSTEM INFORMATION TYPE 1", "SYSTEM INFORMATION TYPE 2", "SYSTEM INFORMATION TYPE 3", "SYSTEM INFORMATION TYPE 4", "SYSTEM INFORMATION TYPE 5", "SYSTEM INFORMATION TYPE 6", "CHANNEL MODE MODIFY", "FREQUENCY REDEFINITION", "LOCATION UPDATING ACCEPT", "LOCATION UPDATING REJECT", "AUTHENTICATION REJECT", "AUTHENTICATION REQUEST", "IDENTITY REQUEST", "TMSI REALLOCATION COMMAND", "CM SERVICE ACCEPT", "CM SERVICE REJECT", "CALL PROCEEDING", "PROGRESS", "START DTMF ACKNOWLEGDE", "START DTMF REJECT", "STOP DTMF ACKNOWLEDGE"
            'LineColor = &HFF00FF   紫色
            LineColor = &HC000C0    '深紫色
       Case "STATUS ENQUIRY", "STATUS", "NOTIFY", "CONGESTION CONTROL", "RELEASE COMPLETE", "RELEASE", "DISCONNECT", "USER INFORMATION", "MODIFY REJECT", "MODIFY COMPLETE", "MODIFY", "SETUP", "CONNECT ACKNOWLEDGE", "CONNECT", "ALERTING", "MM STATUS", "RR STATUS"
            LineColor = &HFFFF00
       Case Else
            LineColor = 0
    End Select
End Sub

Private Sub step_Click()
    On Error Resume Next
    PlayFlag = False
    PlayProcess
End Sub

Private Sub Stop_Click()
    On Error Resume Next
    PlayFlag = False
    My_Picture.Cls
    CurrentRow = StartRow
    NumCount = 1
    Label5.Caption = StartName
    OldCiString = StartCi
    CiString = StartCi
    mapinfo.do "fetch rec " & CurrentRow & " from " & MyTableName
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If PlayFlag = True Then
       PlayProcess
    End If

End Sub

Sub GetCellName()
    Dim TempRow As Integer
    Dim leefind As Integer
    
    On Error Resume Next
    mapinfo.do "select * from cell where ci = " & Chr(34) & CiString & Chr(34) & " into mytemp"
    TempRow = mapinfo.eval("tableinfo(mytemp,8)")
    If TempRow = 0 Then
       MyCellName = "UnKnow"
    Else
       MyCellName = mapinfo.eval("mytemp.cell_name")
       leefind = InStr(MyCellName, Chr(0))
       If leefind > 0 Then
          MyCellName = Trim(Left(MyCellName, leefind - 1))
       End If
    End If
    mapinfo.do "close table mytemp"
End Sub
