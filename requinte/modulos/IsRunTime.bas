Attribute VB_Name = "IsRunTime"
Private Declare Function FindWindow _
    Lib "user32" Alias _
    "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
Public Function TestWindow() As Boolean
Dim ReturnVal As Long
If FindWindow("IDEOwner", vbNullString) > 0 Then
      TestWindow = True
Else
      TestWindow = False
End If
End Function

Function isDesignTime() As Boolean
    On Error Resume Next
    isDesignTime = TestWindow
End Function



