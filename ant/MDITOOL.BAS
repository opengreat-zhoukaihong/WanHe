Attribute VB_Name = "MDITolo"
Declare Sub GetCursorPOS Lib "User" (lpPoint As pointapi)

Type pointapi
     X As Integer
     Y As Integer
End Type

Public show_mark1(0 To 24) As Boolean
Public LI_POS As pointapi
Public comm_le As Integer
Public comm_wi As Integer
Public comm_he As Integer
Public tip_num As Integer
Public left_1 As Integer
Public left_2 As Integer
Public top_1 As Integer
Public top_2 As Integer
Public ncell_tip As Integer

Sub comd_pro(sh_X, sh_Y, comm)
    Dim prompt_x As Integer
    Dim prompt_y As Integer

    On Error Resume Next
    If sh_X > comm.Width / 2 Then
       'prompt_x = New_Base.Left + New_Base.SSFrame3.Left + comm.Left + sh_X * 2 / 3
    Else
      ' prompt_x = New_Base.Left + comm.Left + New_Base.SSFrame3.Left + sh_X
    End If
    If sh_Y > comm.Height / 2 Then
       'prompt_y = New_Base.Top + comm.Top + New_Base.SSFrame3.Top + 300 + comm.Height / 2 + sh_Y
    Else
       'prompt_y = New_Base.Top + comm.Top + 300 + New_Base.SSFrame3.Top + comm.Height + sh_Y
    End If
    'left_1 = (MDIMain.Left + New_Base.Left + comm.Left + New_Base.SSFrame3.Left) / Screen.TwipsPerPixelX
    'left_2 = (MDIMain.Left + New_Base.Left + comm.Left + New_Base.SSFrame3.Left + comm.Width) / Screen.TwipsPerPixelX
    'top_1 = 63 + (MDIMain.Top + New_Base.Top + comm.Top + New_Base.SSFrame3.Top + 300) / Screen.TwipsPerPixelY
    'top_2 = 63 + (MDIMain.Top + New_Base.Top + comm.Top + 300 + New_Base.SSFrame3.Top + comm.Height) / Screen.TwipsPerPixelY

   prompt.Height = prompt.Label1.Height
   prompt.Width = prompt.Label1.Width
   prompt.Left = prompt_x
   prompt.Top = prompt_y
   prompt.ZOrder 0
End Sub


Sub show_pro(comm, X, Y, tt)
   On Error Resume Next
   For i = 0 To 24
       If i = tt Then
          GoTo ss
       End If
       If show_mark1(i) = True Then
          show_mark1(i) = False
          MDIMain.Timer1.Enabled = False
          prompt.Left = -4000
       End If
ss:
   Next
   If show_mark1(tt) = False Then
      show_mark1(tt) = True
      comm_le = comm.Left
      comm_wi = comm.Width
      comm_he = comm.Height
      prompt.Height = prompt.Label1.Height
      prompt.Width = prompt.Label1.Width
      MDIMain.Timer1.Enabled = True
      If tt = 3 Or tt = 4 Or tt = 5 Or tt = 6 Then
         tt = tt - 0.5
      End If
      If tt > 6 Then
         tt = tt - 1.8
      End If
      If X >= 2 * comm.Width / 3 Then
         prompt.Left = 2 * X / 3 + tt * comm.Width
      Else
         prompt.Left = X + tt * comm.Width
      End If
      If Y >= 2 * comm.Height / 3 Then
         prompt.Top = Y * 2 / 3
      Else
         prompt.Top = Y      '1/4
      End If
      prompt.ZOrder 0
   End If

End Sub

Sub Ncell_Hint(sh_X, sh_Y, comm)
    Dim prompt_x As Integer
    Dim prompt_y As Integer
    On Error Resume Next
    If sh_X > comm.Width / 2 Then
       prompt_x = NcellFrm.Left + comm.Left + sh_X * 2 / 3
    Else
       prompt_x = NcellFrm.Left + comm.Left + sh_X
    End If
    If sh_Y > comm.Height / 2 Then
       prompt_y = NcellFrm.Top + comm.Top + 300 + comm.Height / 2 + sh_Y
    Else
       prompt_y = NcellFrm.Top + comm.Top + 300 + comm.Height + sh_Y
    End If
    left_1 = (MDIMain.Left + NcellFrm.Left + comm.Left) / Screen.TwipsPerPixelX
    left_2 = (MDIMain.Left + NcellFrm.Left + comm.Left + comm.Width) / Screen.TwipsPerPixelX
    top_1 = 63 + (MDIMain.Top + NcellFrm.Top + comm.Top + 300) / Screen.TwipsPerPixelY
    top_2 = 63 + (MDIMain.Top + NcellFrm.Top + comm.Top + 300 + comm.Height) / Screen.TwipsPerPixelY

    prompt.Height = prompt.Label1.Height
    prompt.Width = prompt.Label1.Width
    NcellFrm.Timer1.Enabled = True
    prompt.Left = prompt_x
    prompt.Top = prompt_y
    prompt.ZOrder 0
End Sub

