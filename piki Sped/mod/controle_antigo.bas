Attribute VB_Name = "controle"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40

Sub TaskBar(Stat As Boolean)
    Dim Thwnd As Long
    Thwnd = FindWindow("Shell_traywnd", "")
    If Stat Then
        Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    Else
        Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)

    End If
End Sub


