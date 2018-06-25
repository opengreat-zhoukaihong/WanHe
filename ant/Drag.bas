Attribute VB_Name = "DRAG_bas"
Type dbf_field
     space As String * 1
     time As String * 12
     num_frame As String * 10
     lon As String * 12
     lat As String * 12
     message As String * 30
     hex As String * 90
     col(1 To 52) As String * 5
     ncell_num  As String * 1
End Type

'************************
'Input  :sinput
'Output :soutput
'************************

Sub DRAG(sinput, soutput)
    On Error Resume Next
    'Dim Data As typeNormal, new_data As typeNormal, old_data As typeNormal
    Dim Data As typeMarkNormal, new_data As typeMarkNormal, old_data As typeMarkNormal
    Dim mi As String * 1
    Dim recno1 As Long, recno2 As Long
    Dim lon As String * 12, lat As String * 12

    On Error Resume Next
    same = 0
    recno2 = 0
    n1 = 0
    
    Per_Show.Label1.Caption = "正在进行平滑处理"
    Per_Show.Label1.Refresh
    FileCopy sinput, soutput
    Open sinput For Binary As #1
    Open soutput For Binary As #2
    Seek #1, 5
    Get #1, , recno1
    'Seek #1, 2466 + 384
    Seek #1, 151 * 32 + 1 + 1
    If recno1 = 0 Then
       GoTo no_record
    End If
    bline = Fix(recno1 / 100)
    percent_step = 1
    If bline = 0 Then
       bline = 1
       percent_step = 100 / recno1
       End If
    bs = 1
    scnline = 0
    Per_Show.ProgressBar1.Value = 0

    Get #1, , old_data
    lon = old_data.lon
    lat = old_data.lat
    Do While Not EOF(1)
    
       loncopy = lon
       latcopy = lat

       scnline = scnline + 1
       If scnline = bs * bline And Per_Show.ProgressBar1.Value < 99 Then
          Per_Show.ProgressBar1.Value = Per_Show.ProgressBar1.Value + percent_step
          bs = bs + 1
       End If
       
       n1 = n1 + 1
       If n1 >= recno1 Then
          Exit Do
       Else
          Get #1, , Data
       End If
       
       d_loncopy = Data.lon
       d_latcopy = Data.lat
       
       If Data.lon = lon And Data.lat = lat Then
          same = same + 1
       Else
          If same = 0 Then
             Seek #2, 151 * 32 + 1 + 1 + recno2 * 694
             If Val(old_data.FieldCol2(3)) = 0 Then
                old_data.FieldCol2(3) = old_data.FieldCol2(1)
                old_data.FieldCol2(4) = old_data.FieldCol2(2)
             End If
             Put #2, , old_data
             recno2 = recno2 + 1
          Else
'             londx = (lon - data.lon) / (same + 1)
'             latdx = (lat - data.lat) / (same + 1)
               londx = (Val(loncopy) - Val(d_loncopy)) / (same + 1)
               latdx = (Val(latcopy) - Val(d_latcopy)) / (same + 1)
               j = 0
             For i = 1 To same + 1
                 Seek #2, 151 * 32 + 1 + 1 + (recno2) * 694
                 Get #2, , new_data
                 new_data.lon = Val(new_data.lon) - j * londx
                 new_data.lat = Val(new_data.lat) - j * latdx
                       
                 Call gett(new_data.lon)
                 Call gett(new_data.lat)
               
                 Seek #2, 151 * 32 + 1 + 1 + (recno2) * 694
                 If Val(new_data.FieldCol2(3)) = 0 Then
                    new_data.FieldCol2(3) = new_data.FieldCol2(1)
                    new_data.FieldCol2(4) = new_data.FieldCol2(2)
                 End If
                 Put #2, , new_data
                 recno2 = recno2 + 1
                 j = j + 1
             Next
             same = 0
          End If
       End If
       old_data = Data
       lon = old_data.lon
       lat = old_data.lat
          
'If recno2 > 100 Then
'   End
'End If
       DoEvents
       If Convert_Stop = True Then
          Close
          Exit Sub
       End If
    Loop

   If Per_Show.ProgressBar1.Value < 100 Then
      Per_Show.ProgressBar1.Value = 100
   End If
no_record:
  Close #1
  Close #2
End Sub

Sub gett(a)
    On Error Resume Next
    findp = InStr(a, ".")
    lenth = Len(LTrim(RTrim(Right(a, Len(a) - findp))))
    If lenth > 5 Then
       a = Left(a, findp + 5)
    Else
       If lenth < 5 Then
          a = LTrim(RTrim(a))
          For i = 1 To (5 - lenth)
              a = a + "0"
          Next
       End If
    End If
End Sub

