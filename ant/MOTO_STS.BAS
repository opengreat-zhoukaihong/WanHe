Attribute VB_Name = "mot_sts"
Option Explicit
Type sdcch
     bb As String * 1
     Name As String * 10
     ci As String * 5
     scol1 As String * 6
     scol2 As String * 6
     scol3 As String * 6
     scol4 As String * 6
     scol5 As String * 6
     scol6 As String * 6
     scol(1 To 10) As String * 6
     lon As String * 12
     lat As String * 12
     bearing As String * 3
End Type

Type tch
     bb As String * 1
     Name As String * 10
     ci As String * 5
     tcol1 As String * 6
     tcol2 As String * 6
     tcol3 As String * 6
     tcol(1 To 18) As String * 6
     lon As String * 12
     lat As String * 12
     bearing As String * 3
End Type

Type cell
     bb As String * 1
     Name As String * 10
     bs_no As String * 10
     ci As String * 5
End Type

