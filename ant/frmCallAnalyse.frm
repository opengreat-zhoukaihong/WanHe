VERSION 5.00
Begin VB.Form frmCallAnalyse 
   Caption         =   "通话过程统计"
   ClientHeight    =   5055
   ClientLeft      =   4470
   ClientTop       =   405
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCallAnalyse.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6780
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4335
      Index           =   0
      Left            =   165
      TabIndex        =   2
      Top             =   105
      Width           =   6435
      Begin VB.CheckBox Check5 
         Caption         =   "信道释放(CHANNEL RELEASE)"
         Height          =   300
         Index           =   3
         Left            =   3525
         TabIndex        =   27
         Top             =   3855
         Width           =   2595
      End
      Begin VB.CheckBox Check5 
         Caption         =   "释放(RELEASE)"
         Height          =   300
         Index           =   2
         Left            =   345
         TabIndex        =   26
         Top             =   3855
         Width           =   1590
      End
      Begin VB.CheckBox Check5 
         Caption         =   "断开连接(DISCONNECT)"
         Height          =   300
         Index           =   1
         Left            =   3525
         TabIndex        =   25
         Top             =   3435
         Width           =   2205
      End
      Begin VB.CheckBox Check5 
         Caption         =   "切换失败(HANDOVER FAILUER)"
         Height          =   300
         Index           =   0
         Left            =   345
         TabIndex        =   24
         Top             =   3435
         Width           =   2715
      End
      Begin VB.CheckBox Check1 
         Caption         =   "小区重选"
         Height          =   300
         Index           =   12
         Left            =   4845
         TabIndex        =   22
         Top             =   720
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "位置更新失败"
         Height          =   300
         Index           =   15
         Left            =   4845
         TabIndex        =   20
         Top             =   2010
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.CheckBox Check1 
         Caption         =   "位置更新尝试"
         Height          =   300
         Index           =   13
         Left            =   4845
         TabIndex        =   19
         Top             =   1155
         Width           =   1410
      End
      Begin VB.CheckBox Check1 
         Caption         =   "位置更新接受"
         Height          =   300
         Index           =   14
         Left            =   4845
         TabIndex        =   18
         Top             =   1575
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "切换尝试"
         Height          =   300
         Index           =   5
         Left            =   1980
         TabIndex        =   17
         Top             =   720
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "掉话"
         Height          =   300
         Index           =   11
         Left            =   3525
         TabIndex        =   16
         Top             =   1155
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "正常释放"
         Height          =   300
         Index           =   10
         Left            =   3525
         TabIndex        =   15
         Top             =   720
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "非服务区"
         Height          =   300
         Index           =   9
         Left            =   1980
         TabIndex        =   14
         Top             =   2445
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "切换失败"
         Height          =   300
         Index           =   7
         Left            =   1980
         TabIndex        =   13
         Top             =   1590
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "切换成功"
         Height          =   300
         Index           =   6
         Left            =   1980
         TabIndex        =   12
         Top             =   1155
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "噪音通话"
         Height          =   300
         Index           =   8
         Left            =   1980
         TabIndex        =   11
         Top             =   2025
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "非服务区"
         Height          =   300
         Index           =   4
         Left            =   4845
         TabIndex        =   7
         Top             =   2445
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "建立拥塞"
         Height          =   300
         Index           =   2
         Left            =   330
         TabIndex        =   6
         Top             =   1590
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "呼叫建立失败"
         Height          =   330
         Index           =   3
         Left            =   330
         TabIndex        =   5
         Top             =   2010
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.CheckBox Check1 
         Caption         =   "建立通话"
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   4
         Top             =   1170
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "建立尝试"
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   3
         Top             =   735
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "查看信令原因值："
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   4
         Left            =   345
         TabIndex        =   23
         Top             =   3045
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "其他："
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   3
         Left            =   4830
         TabIndex        =   21
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "释放过程："
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   2
         Left            =   3495
         TabIndex        =   10
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "通话过程："
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   1
         Left            =   1980
         TabIndex        =   9
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "建立过程："
         ForeColor       =   &H8000000D&
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   8
         Top             =   345
         Width           =   900
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   330
      Left            =   3435
      TabIndex        =   1
      Top             =   4620
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   345
      Left            =   2265
      TabIndex        =   0
      Top             =   4620
      Width           =   1065
   End
End
Attribute VB_Name = "frmCallAnalyse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Integer, j As Integer
    Dim strSelect(15) As String
    Dim strLegend(15) As String
    Dim MystrSelect As String
    Dim MystrLegend As String
    Dim OpenTableNum As Integer
    Dim MyRows As Integer
    Dim MyMark1 As String, MyMark As String
    Dim Mytempstr As String, MyMsgs As String
    Dim CauseValue() As Integer
    Dim CVString() As String
    
    On Error Resume Next
    Me.Hide
    For i = 0 To 15
        If Check1(i).Value = 1 Then
            Select Case i
                Case 0
                    strSelect(i) = "Left$(mark1,2)=""CA"""
                    strLegend(i) = "建立尝试"
                Case 1
                    strSelect(i) = "Left$(mark1,2)=""CS"""
                    strLegend(i) = "建立通话"
                Case 2
                    'strSelect(i) = "Left$(mark1,7)=""CF 网络拥塞"""
                    strSelect(i) = "mark=""Blocked Call"""
                    strLegend(i) = "建立拥塞"
                Case 3
                    strSelect(i) = "Left$(mark1,2)=""CF"""
                    strLegend(i) = "呼叫建立失败"
                Case 4
                    strSelect(i) = "mark=""No Service"""
                    strLegend(i) = "非服务区"
                Case 5
                    strSelect(i) = "Left$(mark1,3)=""HOA"""
                    strLegend(i) = "切换尝试"
                Case 6
                    strSelect(i) = "Left$(mark1,3)=""HOS"""
                    strLegend(i) = "切换成功"
                Case 7
                    strSelect(i) = "Left$(mark1,3)=""HOF"""
                    strLegend(i) = "切换失败"
                Case 8
                    strSelect(i) = "mark=""Noisy Call"""
                    strLegend(i) = "噪音通话"
                Case 9
                Case 10
                    strSelect(i) = "Left$(mark1,5)=""CD 正常"""
                    strLegend(i) = "正常释放"
                Case 11
                    strSelect(i) = "Left$(mark1,5)=""CD 掉话"""
                    strLegend(i) = "掉话"
                Case 12
                    strSelect(i) = "Left$(mark1,3)=""CRL"""
                    strLegend(i) = "小区重选"
                Case 13
                    strSelect(i) = "Left$(mark1,3)=""LUR"""
                    strLegend(i) = "位置更新尝试"
                Case 14
                    strSelect(i) = "Left$(mark1,3)=""LUA"""
                    strLegend(i) = "位置更新接受"
                Case 15
                    strSelect(i) = "Left$(mark1,3)=""LUF"""
                    strLegend(i) = "位置更新失败"
            End Select
        End If
    Next
    For i = 0 To 15
        If Check1(i).Value = 1 Then
            MystrSelect = MystrSelect & strSelect(i) & " or "
        End If
    Next
    If MystrSelect <> "" Then
        MystrSelect = Left(MystrSelect, Len(MystrSelect) - 3)
       OpenTableNum = mapinfo.eval("NumTables()")
       For i = 1 To OpenTableNum
           If UCase(mapinfo.eval("tableinfo(" & i & ",1)")) = "CALLEVENT" Then
              mapinfo.do "close table CALLEVENT"
              Exit For
           End If
       Next
        
        mapinfo.do "select * from " & tblname & " where " & MystrSelect & " into callevent"
        If mapinfo.eval("tableinfo(CALLEVENT,8)") = 0 Then
            MsgBox "不存在所要分析的采集事件", 64, "提示"
            mapinfo.do "close table callevent"
        Else
            'mapinfo.do "commit table CALLEVENT as " + Chr(34) + Gsm_Path + "\User\CALLEVENT.tab" + Chr(34)
            mapinfo.do "commit table CALLEVENT as " + Chr(34) + "CALLEVENT.tab" + Chr(34)
            mapinfo.do "close table CALLEVENT"
            mapinfo.do "open table " + Chr(34) + "CALLEVENT.tab" + Chr(34)
            MyRows = mapinfo.eval("tableinfo(CALLEVENT,8)")
            mapinfo.do "fetch first from callevent"
            For i = 1 To MyRows
                MyMark = mapinfo.eval("callevent.mark")
                MyMark1 = mapinfo.eval("callevent.mark1")
                If MyMark1 <> "" Then
                    If Left(MyMark1, 2) = "CA" Or Left(MyMark1, 2) = "CS" Or Left(MyMark1, 2) = "CF" Then
                        Select Case Left(MyMark1, 2)
                            Case "CA"
                                Mytempstr = "建立尝试"
                            Case "CS"
                                Mytempstr = "建立通话"
                            Case "CF"
                                Mytempstr = "呼叫建立失败"
                        End Select
                    ElseIf Left(MyMark1, 3) = "HOA" Or Left(MyMark1, 3) = "HOF" Or Left(MyMark1, 3) = "HOS" Or Left(MyMark1, 3) = "CRL" Or Left(MyMark1, 3) = "LUA" Or Left(MyMark1, 3) = "LUR" Or Left(MyMark1, 3) = "LUF" Then
                        Select Case Left(MyMark1, 3)
                            Case "HOA"
                                Mytempstr = "切换尝试"
                            Case "HOF"
                                Mytempstr = "切换失败"
                            Case "HOS"
                                Mytempstr = "切换成功"
                            Case "CRL"
                                Mytempstr = "小区重选"
                            Case "LUR"
                                Mytempstr = "位置更新尝试"
                            Case "LUA"
                                Mytempstr = "位置更新接受"
                            Case "LUF"
                                Mytempstr = "位置更新失败"
                        End Select
                    
                    ElseIf Left(MyMark1, 5) = "CD 正常" Then
                        Mytempstr = "正常释放"
                    ElseIf Left(MyMark1, 5) = "CD 掉话" Then
                        Mytempstr = "掉话"
                    ElseIf Left(MyMark1, 7) = "CD 切换掉话" Then
                        Mytempstr = "掉话"
                    ElseIf Left(MyMark1, 8) = "CD 无服务掉话" Then
                        Mytempstr = "掉话"
                    Else
                        If MyMark = "Blocked Call" Then
                            Mytempstr = "建立拥塞"
                        ElseIf MyMark = "No Service" Then
                            Mytempstr = "非服务区"
                        ElseIf MyMark = "Noisy Call" Then
                            Mytempstr = "噪音通话"
                        End If
                        
                    End If
                Else
                    If MyMark = "Blocked Call" Then
                        Mytempstr = "建立拥塞"
                    ElseIf MyMark = "No Service" Then
                        Mytempstr = "非服务区"
                    ElseIf MyMark = "Noisy Call" Then
                        Mytempstr = "噪音通话"
                    End If
                End If
                mapinfo.do "UPDATE callevent set mark2 = """ & Mytempstr & """ where rowid = " & Format(i)
                mapinfo.do "fetch next from callevent"
            Next
            mapinfo.do "commit table callevent"
            mapinfo.do "Add Map window FrontWindow() Layer callevent"
            MyMsgs = "shade window FrontWindow() callevent with mark2 values ""建立尝试"" Symbol (""Start.bmp"",16776960,24,0),"
            MyMsgs = MyMsgs + """建立通话"" Symbol (""good.bmp"",255,22,0),"     '
            MyMsgs = MyMsgs + """呼叫建立失败"" Symbol (""conn_f.bmp"",16776960,24,0),"
            MyMsgs = MyMsgs + """切换尝试"" Symbol (""hand_com.bmp"",255,24,0),"     '
            MyMsgs = MyMsgs + """切换失败"" Symbol (""hand_f.bmp"",16776960,24,0),"   '
            MyMsgs = MyMsgs + """切换成功"" Symbol (""hand_c.bmp"",16776960,24,0),"      '
            MyMsgs = MyMsgs + """小区重选"" Symbol (""Watchpoi.bmp"",16776960,24,0),"
            MyMsgs = MyMsgs + """位置更新尝试"" Symbol (""New1.bmp"",16776960,24,0),"
            MyMsgs = MyMsgs + """位置更新接受"" Symbol (""loc_acc.bmp"",16776960,24,0),"
            MyMsgs = MyMsgs + """位置更新失败"" Symbol (""LOC_F.bmp"",16776960,24,0),"
            MyMsgs = MyMsgs + """正常释放"" Symbol (""release.bmp"",19711765,24,0),"    '
            MyMsgs = MyMsgs + """掉话"" Symbol (""rele_f.bmp"",19711765,24,0),"        '
            MyMsgs = MyMsgs + """建立拥塞"" Symbol (""Blocked.bmp"",16776960,24,0),"
            MyMsgs = MyMsgs + """非服务区"" Symbol (""NoService.bmp"",16776960,24,0),"
            MyMsgs = MyMsgs + """噪音通话"" Symbol (""Noisy.bmp"",16776960,24,0)"
            mapinfo.do MyMsgs
            mapinfo.do "set legend window FrontWindow() Layer prev Title " + Chr(34) + "通话过程统计 " + tblname + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) ""其余全部"" display off"
            
        End If
    End If
    For j = 0 To 3
        If Check5(j).Value = 1 Then
            If j = 0 Then
                'mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "HANDOVER FAILUER" + Chr(34) + " and Rxle_same1>0 group by RXLE_SAME1 order by Rxle_same1 into Selection1"
                mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "HANDOVER FAILURE" + Chr(34) + " group by RXLE_SAME1 order by Rxle_same1 into Selection1"
            ElseIf j = 1 Then
                'mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "DISCONNECT" + Chr(34) + " and Rxle_same1>0 group by RXLE_SAME1 order by Rxle_same1 into Selection1"
                mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "DISCONNECT" + Chr(34) + " group by RXLE_SAME1 order by Rxle_same1 into Selection1"
            ElseIf j = 2 Then
                'mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "RELEASE" + Chr(34) + " and Rxle_same1>0 group by RXLE_SAME1 order by Rxle_same1 into Selection1"
                mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "RELEASE" + Chr(34) + " group by RXLE_SAME1 order by Rxle_same1 into Selection1"
            Else
                'mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "CHANNEL RELEASE" + Chr(34) + " and Rxle_same1>0 group by RXLE_SAME1 order by Rxle_same1 into Selection1"
                mapinfo.do "Select Rxle_same1 from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "CHANNEL RELEASE" + Chr(34) + " group by RXLE_SAME1 order by Rxle_same1 into Selection1"
            End If
            MyRows = Val(mapinfo.eval("tableinfo(Selection1,8)"))
            If MyRows > 0 Then
                ReDim CauseValue(1 To MyRows) As Integer
                ReDim CVString(1 To MyRows) As String
                For i = 1 To MyRows
                    CauseValue(i) = mapinfo.eval("Selection1.Rxle_same1")
                    mapinfo.do "fetch next from Selection1"
                Next
                mapinfo.do "close table Selection1"
                If j = 0 Then
                   mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "HANDOVER FAILURE" + Chr(34) + " into HOF_Result"
                ElseIf j = 1 Then
                   mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "DISCONNECT" + Chr(34) + " into DC_Result"
                ElseIf j = 2 Then
                   mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "RELEASE" + Chr(34) + " into RL_Result"
                Else
                   mapinfo.do "select * from " & tblname & " where ucase$(MESSAGE)= " + Chr(34) + "CHANNEL RELEASE" + Chr(34) + " into CRL_Result"
                End If
                If j = 0 Or j = 3 Then
                   For i = 1 To MyRows
                       Select Case CauseValue(i)
                           Case 0
                               CVString(i) = "Normal event"
                           Case 1
                               CVString(i) = "Abnormal release,unspecified"
                           Case 2
                               CVString(i) = "Abnormal release,channel unacceptable"
                           Case 3
                               CVString(i) = "Abnormal release,timer expired"
                           Case 4
                               CVString(i) = "Abnormal release,no activity on the radio path"
                           Case 5
                               CVString(i) = "preemptive release"
                           Case 8
                               CVString(i) = "handover impossible,timing advance out of range"
                           Case 9
                               CVString(i) = "Channel mode unacceptable"
                           Case 10
                               CVString(i) = "Frequency not implemented"
                           Case 65
                               CVString(i) = "call already cleared"
                           Case 95
                               CVString(i) = "semantically incorrect message"
                           Case 96
                               CVString(i) = "invalid mandatory information"
                           Case 97
                               CVString(i) = "Message type non-existent or not implenmented"
                           Case 98
                               CVString(i) = "Message type not compatible with protocol state"
                           Case 100
                               CVString(i) = "Conditional IE error"
                           Case 101
                               CVString(i) = "No cell allocation available"
                           Case 111
                               CVString(i) = "protocol error unspecified"
                       End Select
                   Next
                Else
                   For i = 1 To MyRows
                       Select Case CauseValue(i)
                           Case 1
                               CVString(i) = "Unassiagned number"
                           Case 3
                               CVString(i) = "No route to destination"
                           Case 6
                               CVString(i) = "Channel unacceptable"
                           Case 16
                               CVString(i) = "Normal clearing"
                           Case 17
                               CVString(i) = "User busy"
                           Case 18
                               CVString(i) = "No user responding"
                           Case 19
                               CVString(i) = "User alerting,no answer"
                           Case 21
                               CVString(i) = "Call rejected"
                           Case 22
                               CVString(i) = "Number changed"
                           Case 26
                               CVString(i) = "Non selected user clearing"
                           Case 27
                               CVString(i) = "Destination out of order "
                           Case 28
                               CVString(i) = "Incomplete number"
                           Case 29
                               CVString(i) = "Facility rejected"
                           Case 30
                               CVString(i) = "Response to status enquiry"
                           Case 31
                               CVString(i) = "Normal,unspecified"
                           Case 34
                               CVString(i) = "No circuit/channel available"
                           Case 38
                               CVString(i) = "Network out of order"
                           Case 41
                               CVString(i) = "Temporary failure"
                           Case 42
                               CVString(i) = "Switching equipment congestion"
                           Case 43
                               CVString(i) = "Access information discarded"
                           Case 44
                               CVString(i) = "Requested circuit/channel not available"
                           Case 47
                               CVString(i) = "Resources unavailable,unspecified"
                           Case 49
                               CVString(i) = "Quality of service unavailable"
                           Case 50
                               CVString(i) = "Requested facility not subscribed"
                           Case 55
                               CVString(i) = "Incoming calls barred within the CUG"
                           Case 57
                               CVString(i) = "Bearer capability not authorized"
                           Case 58
                               CVString(i) = "Bearer capability not presently available"
                           Case 63
                               CVString(i) = "Service or option not available,unspecified"
                           Case 65
                               CVString(i) = "Bearer service not implemented"
                           Case 68
                               CVString(i) = "ACM equal to or greater than ACMmax"
                           Case 69
                               CVString(i) = "Requested facility not implemented"
                           Case 70
                               CVString(i) = "Only restricted digital information bearer"
                           Case 79
                               CVString(i) = "Service or option not implemented"
                           Case 81
                               CVString(i) = "Invalid transaction identrfier value"
                           Case 87
                               CVString(i) = "User not member of CUG"
                           Case 88
                               CVString(i) = "Incompatible destination"
                           Case 91
                               CVString(i) = "Invalid mandatory information"
                           Case 95
                               CVString(i) = "Semantically incorrect message"
                           Case 96
                               CVString(i) = "Invalid mandatory information"
                           Case 97
                               CVString(i) = "Message type non-existent or not implemented"
                           Case 98
                               CVString(i) = "Message type not compatible with protocol state"
                           Case 99
                               CVString(i) = "Information element non-existent or not implemented"
                           Case 100
                               CVString(i) = "Conditional IE error "
                           Case 101
                               CVString(i) = "Message not compatible with protocol state"
                           Case 102
                               CVString(i) = "Recovery on timer expiry"
                           Case 111
                               CVString(i) = "Protocol error,unspecified"
                           Case 127
                               CVString(i) = "Interworking,unspecified"
                       End Select
                   Next
                End If
                'mapinfo.do "Add Map window FrontWindow() Layer  Result"
                If j = 0 Then
                    mapinfo.do "Add Map window FrontWindow() Layer HOF_Result"
                    MyMsgs = "shade window FrontWindow() HOF_Result with RXLE_SAME1 values "
                ElseIf j = 1 Then
                    mapinfo.do "Add Map window FrontWindow() Layer DC_Result"
                    MyMsgs = "shade window FrontWindow() DC_Result with RXLE_SAME1 values "
                ElseIf j = 2 Then
                    mapinfo.do "Add Map window FrontWindow() Layer RL_Result"
                    MyMsgs = "shade window FrontWindow() RL_Result with RXLE_SAME1 values "
                Else
                    mapinfo.do "Add Map window FrontWindow() Layer CRL_Result"
                    MyMsgs = "shade window FrontWindow() CRL_Result with RXLE_SAME1 values "
                End If
                For i = 1 To MyRows
                    'MyMsgs = MyMsgs & Format(CauseValue(i)) & " Symbol (41," & Format(MyRndColor(i)) & " ,8,""MapInfo Cartographic"",0,0),"
                    MyMsgs = MyMsgs & Format(CauseValue(i)) & " Symbol (41," & Format(MyRndColor((i + j * 100) Mod 375)) & " ,8,""MapInfo Cartographic"",0,0),"
                Next
                MyMsgs = Left(MyMsgs, Len(MyMsgs) - 1)
                mapinfo.do MyMsgs
                
                If legendid = 0 Then
                   mapinfo.do "Set Next Document Parent " & MDIMain.hWnd & " Style 0"
                   mapinfo.do "Create Legend From Window  Frontwindow()"
                   legendid = mapinfo.eval("windowinfo(1009,12)")
                End If
                
                If j = 0 Then
                   MyMsgs = " Title " + Chr(34) + "HANDOVER FAILURE 事件原因" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off ,"
                ElseIf j = 1 Then
                   MyMsgs = " Title " + Chr(34) + "DISCONNECT 事件原因" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off ,"
                ElseIf j = 2 Then
                   MyMsgs = " Title " + Chr(34) + "RELEASE 事件原因" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off ,"
                Else
                   MyMsgs = " Title " + Chr(34) + "CHANNEL RELEASE 事件原因" + Chr(34) + " Font(""宋体"",0,9,0) ascending off ranges Font(""宋体"",0,9,0) """" display off ,"
                End If
                For i = 1 To MyRows
                    MyMsgs = MyMsgs + Chr(34) + Format(CauseValue(i)) & ": " & CVString(i) + Chr(34) + " display on,"
                Next
                MyMsgs = Left(MyMsgs, Len(MyMsgs) - 1)
                mapinfo.do "set legend window FrontWindow()  Layer prev " & MyMsgs
            End If
        End If
    Next
    Unload Me
End Sub

Private Sub Command2_Click()

    On Error Resume Next
    Unload Me
    
End Sub

