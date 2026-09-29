Attribute VB_Name = "Module1"
Option Explicit

Public OldWindowProc As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As WINDOWPOS) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)

Type WINDOWPOS
    hwnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    flags As Long
End Type
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47
' *********************************************
' Process messages.
' *********************************************
Public Function NewWindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As WINDOWPOS) As Long
Static done_before As Boolean
Static aspect As Single
Dim new_aspect As Single

    ' Keep the aspect ratio.
    If msg = WM_WINDOWPOSCHANGING Then
        If lParam.cy > 0 Then
            ' Save the aspect ratio the first.
            If Not done_before Then
                aspect = lParam.cx / lParam.cy
                done_before = True
            End If
            
            new_aspect = lParam.cx / lParam.cy
            If new_aspect > aspect Then
                ' Too short/wide. Make it taller.
                lParam.cy = lParam.cx / aspect
            Else
                ' Too tall/thin. Make it wider.
                lParam.cx = aspect * lParam.cy
            End If
        End If
    End If

    ' Continue normal processing. VERY IMPORTANT!
    NewWindowProc = CallWindowProc( _
        OldWindowProc, hwnd, msg, wParam, _
        lParam)
End Function


