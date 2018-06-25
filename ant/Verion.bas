Attribute VB_Name = "Verion"
Public My_Ver As Integer
Public Convert_Stop As Boolean
Public Is_Done As Boolean


Function Get_date() As Integer
   Dim dd As String
   Dim year, month, day As Integer

   On Error Resume Next
   dd = DATE
   year = Val(Left(dd, 2))
   month = Val(Mid(dd, 4, 2))
   If (year <= 98 And month < 5) Or (year < 98) Then
      Get_date = 1
   Else
      Get_date = 0
   End If
End Function

