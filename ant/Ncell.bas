Attribute VB_Name = "ncell_bas"
Public ncell_file As String

Sub getnb(linetxt, a1, a2, a3)
    Dim finds As Integer
    Dim FindChar As String * 1
    
    On Error Resume Next
    finds = InStr(linetxt, Chr(9))
    If finds > 0 Then
       FindChar = Chr(9)
    Else
       FindChar = " "
    End If
    finds = InStr(linetxt, FindChar)
    nbci = Left(linetxt, finds - 1)
    a2 = Right(nbci, 4)
    a2 = (a2)
    linetxt = Trim(Right(linetxt, Len(linetxt) - finds))
    finds = InStr(linetxt, FindChar)
    a3 = Left(linetxt, finds - 1)
    linetxt = Trim(Right(linetxt, Len(linetxt) - finds))
    For i = 1 To 4
       finds = InStr(linetxt, FindChar)
       linetxt = Trim(Right(linetxt, Len(linetxt) - finds))
    Next
    finds = InStr(linetxt, FindChar)
    a1 = Left(linetxt, finds)
End Sub

