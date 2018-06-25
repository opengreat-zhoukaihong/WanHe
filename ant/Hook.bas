Attribute VB_Name = "Hook"
Option Explicit

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Const WH_MOUSE = 7
Public Const GWL_WNDPROC = -4
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203

Public IsSetMyHook As Boolean
Public lpMyMapWndProc As Long
Public lngMyMapHWnd As Long
 
Public Sub SetMyHook(hWnd As Long)

    On Error Resume Next
    lngMyMapHWnd = hWnd
    lpMyMapWndProc = SetWindowLong(lngMyMapHWnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Public Sub UnMyHook()
    Dim lngReturnValue As Long
    
    On Error Resume Next
    lngReturnValue = SetWindowLong(lngMyMapHWnd, GWL_WNDPROC, lpMyMapWndProc)

End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Select Case uMsg
        Case WM_LBUTTONDBLCLK
             MsgBox "Hello_LBUTTON"
        Case WM_RBUTTONUP
             MsgBox "Hello_RBUTTON"
        Case Else
            WindowProc = CallWindowProc(lpMyMapWndProc, hw, uMsg, wParam, lParam)
    End Select
End Function
