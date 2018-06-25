Attribute VB_Name = "errison"
Type cell_stru
     b As String * 1
     Name As String * 10
     bs_no As String * 10
     ci As String * 5
     ARFCN As String * 3
     BSIC As String * 3
     bearing As String * 3
     Lac As String * 5
     bcch(1 To 6) As String * 3
     ANT_TYPE As String * 12
     ant_angle As String * 3
     downtilt As String * 3
     MAX_BTS As String * 2
     MAX_MS As String * 2
     pref As String * 2
     tch_num As String * 2
     bsc_stsge As String * 7
     bsc_type As String * 5
     bts_type As String * 9
     power_type As String * 3
     photo As String * 12
     time As String * 8
     lon As String * 12
     lat As String * 12
     microcell As String * 1
End Type
Type Oldcellstru
     b As String * 1
     Name As String * 10
     bs_no As String * 10
     ci As String * 5
     ARFCN As String * 3
     BSIC As String * 3
     bearing As String * 3
     Lac As String * 5
     bcch(1 To 6) As String * 3
     ANT_TYPE As String * 12
     ant_angle As String * 3
     downtilt As String * 3
     MAX_BTS As String * 2
     MAX_MS As String * 2
     pref As String * 2
     tch_num As String * 2
     bsc_stsge As String * 7
     bsc_type As String * 5
     bts_type As String * 9
     power_type As String * 3
     photo As String * 12
     time As String * 8
     lon As String * 12
     lat As String * 12
End Type

Type Oldcellstru1
     b As String * 1
     Name As String * 15
     bs_no As String * 10
     ci As String * 5
     ARFCN As String * 3
     BSIC As String * 3
     bearing As String * 3
     Lac As String * 5
     NONBCCH As String * 32
     downtilt As String * 3
     MAX_BTS As String * 2
     MAX_MS As String * 2
     time As String * 8
     lon As String * 12
     lat As String * 12
     microcell As String * 1
     NCELL(1 To 16) As String * 10
End Type

Type NewCellStru
     b As String * 1
     Name As String * 15
     bs_no As String * 10
     ci As String * 5
     ARFCN As String * 3
     BSIC As String * 3
     bearing As String * 3
     Lac As String * 5
     NONBCCH As String * 64
     downtilt As String * 3
     MAX_BTS As String * 2
     MAX_MS As String * 2
     time As String * 8
     lon As String * 12
     lat As String * 12
     microcell As String * 1
     NCELL(1 To 16) As String * 10
End Type

Type NewCell1800
     b As String * 1
     Name As String * 21
     bs_no As String * 10
     ci As String * 5
     ARFCN As String * 3
     BSIC As String * 3
     bearing As String * 3
     Lac As String * 5
     NONBCCH As String * 64
     downtilt As String * 3
     MAX_BTS As String * 2
     ANT_HEIGH As String * 3
     MAX_MS As String * 2
     ANT_GAIN As String * 3
     ANT_TYPE As String * 15
     time As String * 8
     lon As String * 12
     lat As String * 12
     BASETYPE As String * 1
     LENGTH As String * 5
     NCELL(1 To 16) As String * 10
End Type

Public load_sam As Integer
Public load_new As Integer
Public FileNoMatch As String
Public CellNoUpdate As String

