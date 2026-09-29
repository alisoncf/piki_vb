Attribute VB_Name = "controle"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

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


Function F_SeparaTexto(Texto As String, Optional sSepara As String = "-;| _.:", Optional bArmazenaSeparador As Boolean = True) As Collection
    Dim cTexto As New Collection
    Dim iCount As Integer
    Set cTexto = Nothing
    Set cTexto = New Collection
    Dim sPalavra As String
    Dim sBuf As String
    For iCount = 1 To Len(Texto)
        sBuf = Mid(Texto, iCount, 1)
        If sBuf = "-" Then
'            MsgBox ""
        End If
        If InStr(sSepara, sBuf) > 0 Then
            If bArmazenaSeparador Then
                If Trim(sPalavra) <> "" Then
                
                
                 
                        cTexto.Add sPalavra & sBuf
                
                ElseIf (sBuf) <> "" Then
                       cTexto.Add sPalavra & sBuf
                End If
                sPalavra = ""
            Else
                cTexto.Add sPalavra
                sPalavra = ""
            End If
        Else
            sPalavra = sPalavra & sBuf
            
        End If
    Next
    If sPalavra <> "" Then
        cTexto.Add sPalavra
    End If
    Set F_SeparaTexto = cTexto
End Function

Function IsInCollection(c As Collection, Valor As String) As Integer
    
     Dim i As Integer
     For iCount = 1 To c.count
        If Trim(c(iCount)) = Trim(Valor) Then
            IsInCollection = iCount
            Exit For
        End If
     Next

End Function

