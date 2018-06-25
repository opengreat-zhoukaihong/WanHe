Attribute VB_Name = "Start_Module"
Option Explicit
Public MaxBcch As Integer, MinBcch As Integer

Sub main()
    On Error Resume Next
    MDIMain.Show
End Sub

Sub GetBcchMaxMin()
    Dim i As Integer, j As Integer
    Dim CellRow As Integer
    Dim NonBcchtemp As String, MyNonBcch As String
    
    On Error Resume Next
       
    If MaxBcch = 0 Or MinBcch = 0 Then
       mapinfo.do "select max(arfcn) from cell into mytemp"
       MaxBcch = mapinfo.eval("mytemp.col1")
       mapinfo.do "select * from cell where (arfcn <> 0) into temp"
       mapinfo.do "select min(arfcn) from temp into mytemp"
       MinBcch = mapinfo.eval("mytemp.col1")
       CellRow = mapinfo.eval("tableinfo(cell,8)")
       mapinfo.do "fetch first from cell"
       For i = 1 To CellRow
           NonBcchtemp = Trim(mapinfo.eval("cell.non_bcch"))
           For j = 1 To 16
               If InStr(NonBcchtemp, ",") > 0 Then
                  MyNonBcch = Left(NonBcchtemp, InStr(NonBcchtemp, ",") - 1)
                  NonBcchtemp = Trim(Right(NonBcchtemp, Len(NonBcchtemp) - InStr(NonBcchtemp, ",")))
               Else
                  MyNonBcch = NonBcchtemp
                  NonBcchtemp = ""
               End If
               If Val(MyNonBcch) > 0 Then
                  If Val(MyNonBcch) > MaxBcch Then
                     MaxBcch = Val(MyNonBcch)
                  Else
                     If Val(MyNonBcch) < MinBcch Then
                        MinBcch = Val(MyNonBcch)
                     End If
                  End If
               Else
                  Exit For
               End If
           Next
           mapinfo.do "fetch next from cell"
       Next

       mapinfo.do "close table mytemp"
       mapinfo.do "close table temp"
    End If
End Sub
