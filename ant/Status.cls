VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Micallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub SetStatusText(ByVal CmdStr As String)
    Dim dd As Long, length As Integer
    Dim cmd As String, lon As String, lat As String
    Dim st1 As String, st2 As String, st3 As String
    Dim i As Integer
    
    On Error Resume Next
    cmd = ""

    length = Len(CmdStr)
    For i = 1 To length
        dd = Asc(Mid(CmdStr, i, 1))
        If dd <> 9 Then
           cmd = cmd + Mid(CmdStr, i, 1)
        Else
           cmd = cmd + " "
        End If
    Next i

If dis_flag = 1 Then
    dd = InStr(cmd, Chr(-24093))
    If dd = 0 Or InStr(cmd, ".") = 0 Then
       GoTo s1
    End If
    lon = Left(cmd, dd - 1)
    cmd = Right(cmd, Len(cmd) - dd)
    st1 = Int(Val(lon))
    st2 = Int((Val(lon) - Val(st1)) * 60)
    st3 = Int((Val(lon) - Val(st1) - Val(st2) / 60) * 3600)
    Call getst(st2)
    Call getst(st3)
    lon = st1 + Chr(-24093) + st2 + "'" + st3 + Chr(34) + Chr(44)
    dd = InStr(cmd, Chr(44))
    cmd = LTrim$(Right(cmd, Len(cmd) - dd))
    dd = InStr(cmd, Chr(-24093))
    lat = Left(cmd, dd - 1)
    cmd = Right(cmd, Len(cmd) - dd)
    st1 = Int(Val(lat))
    st2 = Int((Val(lat) - Val(st1)) * 60)
    st3 = Int((Val(lat) - Val(st1) - Val(st2) / 60) * 3600)
    Call getst(st2)
    Call getst(st3)
    lat = st1 + Chr(-24093) + st2 + "'" + st3 + Chr(34)
    cmd = lon + " " + lat + " " + cmd
s1:
    MDIMain.StatusBar.Panels(3).Text = cmd
Else
       MDIMain.StatusBar.Panels(3).Text = cmd
  End If
End Sub

Sub getst(st)

    On Error Resume Next
    If Len(st) > 2 Then
       st = Mid(st, 1, 2)
    Else
       If Len(st) < 2 Then
          st = "0" + st
       End If
    End If
End Sub

