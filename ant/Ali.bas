Attribute VB_Name = "ALI"
Type e_cch
     b As String * 1
     CCH As String * 10
     ID As String * 10
     CSR As String * 6
     SA As String * 7
     ss As String * 7
     cch_col(1 To 6) As String * 6
     lon As String * 12
     lat As String * 12
     bearing As String * 3
End Type

Type e_tch
     b As String * 1
     tch As String * 10
     ID As String * 10
     CA As String * 7
     CS As String * 7
     tch_col(1 To 11) As String * 6
     lon As String * 12
     lat As String * 12
     bearing As String * 3
End Type

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

