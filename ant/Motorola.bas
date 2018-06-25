Attribute VB_Name = "Motorola"
'*******************************************
'Input:   sinput1---------source data
'         sinput2---------source dbf
'output:  soutput
'******************************************

Sub getfield(lines, ss)  'Get field data from the source data
    On Error Resume Next
    finds = InStr(lines, " ")
    If finds = 0 Then
       ss = lines
       lines = ""
    Else
       ss = Left(lines, finds - 1)
       lines = LTrim$(Right(lines, Len(lines) - finds))
    End If
End Sub

