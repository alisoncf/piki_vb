Attribute VB_Name = "controle"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40


Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" _
(ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DrawIcon Lib "user32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" _
(ByVal hIcon As Long) As Long

Public Type fonte
    nome As String
    tamanho As Integer
    cor_face As Single
    cor_fundo As Single
    negrito As Boolean
    Italico As Boolean
    Sublinhado As Boolean
    
    
End Type
Sub TaskBar(Stat As Boolean)
    Dim Thwnd As Long
    Thwnd = FindWindow("Shell_traywnd", "")
    If Stat Then
        Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    Else
        Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)

    End If
End Sub


Function F_SeparaTexto(Texto As String, Optional sSepara As String = "-;| /\_.:", _
    Optional bArmazenaSeparador As Boolean = True, Optional formato As String = "", _
    Optional aspas As String = "") As Collection
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
            
            If formato <> "" Then
                 sPalavra = Format(sPalavra, formato)
            End If
            If aspas <> "" Then
                sPalavra = aspas & sPalavra & aspas
            End If
            If bArmazenaSeparador Then
                If Trim(sPalavra) <> "" Then
                    
                
                 
                        cTexto.Add sPalavra & sBuf
                
                ElseIf (sBuf) <> "" Then
                       cTexto.Add sPalavra & sBuf
                End If
                sPalavra = ""
            Else
                If InStr(1, sSepara, sPalavra) = 0 Then
                    cTexto.Add sPalavra
                End If
                
                sPalavra = ""
            End If
        Else
            
            sPalavra = sPalavra & sBuf
            
        End If
    Next
    If sPalavra <> "" Then
    
        If formato <> "" Then
                 sPalavra = Format(sPalavra, formato)
            End If
            If aspas <> "" Then
                sPalavra = aspas & sPalavra & aspas
            End If
    
        cTexto.Add sPalavra
    End If
    Set F_SeparaTexto = cTexto
End Function

Function IsInCollection(c As Collection, valor As String) As Integer
    
     Dim i As Integer
     For iCount = 1 To c.Count
        If Trim(c(iCount)) = Trim(valor) Then
            IsInCollection = iCount
            Exit For
        End If
     Next

End Function

Sub TotalCeqBdivA(a As Object, b As Object, c As Object, Optional percent As Integer = 100)
    On Error GoTo trata_Erro
    
    c.Text = FormatVal(CCur(b.Text) / CCur(a.Text) * percent)
    
Exit Sub
trata_Erro:
    c.Text = "0"
End Sub

Sub TotalAeqBmultC(a As Object, b As Object, c As Object)
    On Error GoTo trata_Erro
    
    a.Text = FormatVal(CCur(b.Text) * CCur(c.Text) / 100)
    
    Exit Sub
trata_Erro:
    a.Text = "0"
    
End Sub
Function F_confMsflxCH(msflx As MSFlexGrid, colven As Integer, colch As Integer, _
    colage As Integer, colcco As Integer, colbco As Integer, colcpf As Integer, _
    colnome As Integer, Optional coltipo_doc As Integer = -1) As Boolean
     Dim iCount As Integer
     Dim sTipo As String
     Dim Retorno As Boolean
     Dim strSQL As String
     Dim iPassa As Boolean
     iPassa = True
     Retorno = True
     strSQL = "os seguintes dados estão faltando: " & Chr(13) & Chr(13)
     With msflx
        For iCount = 1 To .rows - 1
            iPassa = True
            .row = iCount
                If coltipo_doc > -1 Then
                    .col = coltipo_doc
                    If Left(.Text, 1) <> "2" And Left(.Text, 1) <> "C" Then
                        iPassa = False
                    End If
                End If
                
                If iPassa = True Then
                    
                    .col = colch
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & "  número do cheque" & Chr(13)
                        Retorno = False
                    End If
                    .col = colbco
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " banco" & Chr(13)
                        Retorno = False
                    End If
                    .col = colage
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " agência" & Chr(13)
                        Retorno = False
                    End If
                    .col = colcco
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " conta" & Chr(13)
                        Retorno = False
                    End If
                    .col = colnome
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " nome do titular" & Chr(13)
                        Retorno = False
                    End If
                    .col = colcpf
                    If .Text = "" Then
                        strSQL = strSQL & "linha " & iCount & " CPF do titular" & Chr(13)
                        Retorno = False
                    End If
                End If
                
                
            
        Next
     End With
     F_confMsflxCH = Retorno
     If Retorno = False Then
        MsgBox strSQL, vbExclamation
        
     End If
End Function


Sub P_SelecTexto(objTexto As Object)
    ganhaFoco objTexto
End Sub
Function CollectionToNotIn(c As Collection) As String
    Dim i As Integer
    Dim s As String
    For i = 1 To c.Count
        s = s & c(i)
        If i <> c.Count Then
            s = s & ","
        End If
        
    Next
    CollectionToNotIn = s
End Function

Sub P_AjustaHora(obj As TextBox)
    Dim strHora As String
    Dim sAux As String
    Dim iCount As Integer
    Dim sLef As String
    Dim sRig As String
    
    On Error GoTo trata_Erro
    If Len(obj.Text) = 5 Then
        obj.Text = Left(obj.Text, 2) & ":" & Right(obj.Text, 2)
    End If
    If Len(obj.Text) = 4 Then
        obj.Text = Left(obj.Text, 2) & ":" & Right(obj.Text, 2)
    End If
    If Len(obj.Text) = 3 Then
        obj.Text = Left(obj.Text, 1) & ":" & Right(obj.Text, 2)
    End If
    If Len(obj.Text) = 2 Then
        obj.Text = Left(obj.Text, 2) & ":00"
    End If
    If Len(obj.Text) = 1 Then
        obj.Text = "0" & obj.Text & ":00"
    End If
    obj.Text = Format(Trim(obj.Text), "hh:mm")
    strHora = obj.Text
    For iCount = 1 To Len(strHora)
        sAux = Mid(strHora, iCount, 1)
        If InStr(1, ".;-:,/\| ", sAux) > 0 Then
            sRig = Format(F_SoNumero(Right(strHora, Len(strHora) - iCount)), "0#")
            sLef = Format(F_SoNumero(sLef), "0#")
            Exit For
        Else
            sLef = sLef & sAux
        End If
        
    Next
    If sLef = "" Then
        sLef = "00"
    End If
    If sRig = "" Then
        sRig = "00"
    End If
    If Int(sLef) > 23 Then
        MsgBox "Hora Inválida!", vbCritical
        sLef = "00"
    End If
    If Int(sRig) > 59 Then
        MsgBox "Minuto Inválida!", vbCritical
        sRig = "00"
    End If
    obj.Text = sLef & ":" & sRig
    Exit Sub
trata_Erro:
End Sub

Public Sub Sendkeys(Text$, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys Text, wait
   Set WshShell = Nothing
End Sub
