Attribute VB_Name = "restore"
Function restore_street() As Integer
    Dim ch  As Byte
    Dim ch1  As Byte
    
    On Error Resume Next
    Gsm_FileName = Gsm_Path + "\map\street.map"
    Gsm_File2 = Gsm_Path + "\map\gsm.tag"
    If dir(Gsm_File2, 0) = "" Then
       Exit Function
    End If
    If dir(Gsm_FileName, 0) <> "" Then
       Kill Gsm_FileName
    End If
    FileCopy Gsm_File2, Gsm_FileName
    Open Gsm_File2 For Binary As #1
    Open Gsm_FileName For Binary As #2
    i = 1
    If FileLen(Gsm_File2) < 1024 Then
       GoTo Close_File
    End If
    While i < &H400
          Get #1, i, ch
          ch1 = &HFF - ch
          ch1 = &HCD Xor ch1
          Put #2, , ch1
          i = i + 1
    Wend
    
Close_File:
    Close #1
    Close #2
End Function

