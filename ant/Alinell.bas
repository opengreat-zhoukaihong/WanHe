Attribute VB_Name = "ALINELL"
Type coltype
     arfcn_c As String * 3
     ci_c As String * 5
     bsic_c As String * 3
     bs_no_c As String * 10
End Type

Type aell
     b As String * 1
     bs_name As String * 10
     bs_no As String * 10
     ci As String * 5
     Lac As String * 5      'lac  = arfcn
     col(1 To 16) As coltype
End Type
Public ncell_file As String
