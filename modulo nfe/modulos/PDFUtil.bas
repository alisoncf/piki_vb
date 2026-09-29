Attribute VB_Name = "PDFUtil"
Global Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
   "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
   As String, ByVal lpFile As String, ByVal lpParameters _
   As String, ByVal lpDirectory As String, ByVal nShowCmd _
   As Long) As Long

Public Declare Function apiFindWindow Lib "user32" Alias "FindWindowA" _
   (ByVal lpClassName As Any, ByVal lpCaption As Any) As Long


Sub ImprimirPDF(caminho As String)

    Dim c As String
    c = caminho
    Dim hwnd
    Dim StartDoc
    hwnd = apiFindWindow("OPUSAPP", "0")

    StartDoc = ShellExecute(hwnd, "print", c, "", _
      "C:\", SW_SHOWNORMAL)


End Sub
Sub AbrirPDF(caminho As String)

    Dim c As String
    c = caminho
    Dim hwnd
    Dim StartDoc
    hwnd = apiFindWindow("OPUSAPP", "0")

    StartDoc = ShellExecute(hwnd, "print", c, "", _
      "C:\", SW_SHOWNORMAL)


End Sub

Function GS_On() As Boolean
    

End Function
