Attribute VB_Name = "Arquivo"

Option Explicit

Private Const CP_UTF8 As Long = 65001

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long

Private Function OpenAppendUTF8(ByVal FileName As String) As Integer
    OpenAppendUTF8 = FreeFile(0)
    Open FileName For Binary Access Write As #OpenAppendUTF8
    Seek #OpenAppendUTF8, LOF(OpenAppendUTF8) + 1
End Function

Private Sub WriteUTF8( _
    ByVal FNum As Integer, _
    ByVal Text As String, _
    Optional ByVal NL As Boolean)

    Dim lngResult As Long
    Dim UTF8() As Byte

    If NL Then Text = Text & vbNewLine
    lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), Len(Text), _
                                    0, 0, 0, 0)
    If lngResult > 0 Then
        ReDim UTF8(lngResult - 1)
        WideCharToMultiByte CP_UTF8, 0, StrPtr(Text), Len(Text), _
                            VarPtr(UTF8(0)), lngResult, 0, 0
        Put #FNum, , UTF8
    End If
End Sub



Function GravaArquivoExternoUTF8(caminho As String, conteudo As String) As Boolean
    On Error GoTo trata_Erro
    Dim ado As New adodb.Stream
    
    ado.Open
    ado.Charset = "UTF-8"
    ado.WriteText Trim(conteudo)
    ado.SaveToFile caminho, adSaveCreateOverWrite
    ado.Close
    GravaArquivoExternoUTF8 = True
    
trata_Erro:
    
    
End Function

Function GravaArquivoExterno(caminho As String, conteudo) As Boolean
    
    On Error Resume Next
    Close #1
    
    Open caminho For Output As #1
    Print #1, conteudo
    
    Close #1
    GravaArquivoExterno = True
End Function
Function ImprimeEmArquivoExterno(conteudo, buf As Integer) As Boolean
    On Error GoTo trata_Erro
    Print #buf, conteudo
    ImprimeEmArquivoExterno = True
    
    Exit Function
trata_Erro:
    ImprimeEmArquivoExterno = False
End Function
Function AbreArquivoExternoLeitura(caminho As String, buf As Integer) As Boolean
    On Error GoTo trata_Erro
    If FechaArquivoExterno(buf) Then
    End If
    Open caminho For Input As #buf
    AbreArquivoExternoLeitura = True
    Exit Function
trata_Erro:
    AbreArquivoExternoLeitura = False
End Function



Function AbreArquivoExterno(caminho As String, buf As Integer) As Boolean
    On Error GoTo trata_Erro
    If FechaArquivoExterno(buf) Then
    End If
    Open caminho For Output As #buf
    AbreArquivoExterno = True
    Exit Function
trata_Erro:
    AbreArquivoExterno = False
End Function
Function FechaArquivoExterno(buf As Integer) As Boolean
    On Error GoTo trata_Erro
    Close #buf
    FechaArquivoExterno = True
    Exit Function
trata_Erro:
    FechaArquivoExterno = False
End Function


Function LeArquivoExterno(caminho As String, Optional buf As Integer = 1) As String
    Dim Linha As String
    On Error GoTo trata_Erro
    
    Open caminho For Input As #buf
    Do While Not EOF(buf)
        Line Input #buf, Linha
        LeArquivoExterno = LeArquivoExterno & Linha
    Loop
    Close #buf
    Exit Function
trata_Erro:
    MsgBox Err.Description, vbExclamation

End Function

Function LeArquivoExternoUTF8(caminho As String) As String
    On Error GoTo trata_Erro
    Dim ado As New adodb.Stream
    
    ado.Open
    ado.Charset = "UTF-8"
    ado.LoadFromFile caminho
    LeArquivoExternoUTF8 = ado.ReadText
    
      
    ado.Close
    
    
trata_Erro:
    
    
End Function

Sub ExportaGridToCSV(dbgrd As Object, dtVinculado As Object, Optional quebra As String = ";", Optional caminhoArquivo As String = "")
    
    Dim c As Collection
    Set c = New Collection
    Dim conteudo As String
    
    If caminhoArquivo = "" Then
        caminhoArquivo = BuscaNomeArquivoExterno("*.csv")
    
    End If
    
    On Error GoTo trata_Erro
    dtVinculado.Visible = False
    
    
    Dim sPath As String
    Dim strSQL As String
    Dim iQtdColsSelec As Integer
    Dim iCount As Integer
    Dim jCount As Integer
    Dim i As Integer
    
'
    P_Status "Gerando arquivo externo"
    conteudo = ""
    
    
    
    
    
    With dbgrd
        
        
            For iCount = 0 To dbgrd.Columns.Count - 1
                If dbgrd.Columns(iCount).Visible = True Then
                    conteudo = conteudo & dbgrd.Columns(iCount).caption & quebra
                
                End If
            Next
        
            conteudo = conteudo & CrLf

        
        
    End With
    'dtVinculado.Refresh
    With dtVinculado
        
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        Do While Not .EOF
            
            
            
            For iCount = 0 To dbgrd.Columns.Count - 1
                        
                If dbgrd.Columns(iCount).Visible = True Then
                    conteudo = conteudo & dbgrd.Columns(iCount).Text & quebra
                
                End If
                
            Next
            conteudo = conteudo & CrLf
            .MoveNext
        Loop
    End With
    
    GravaArquivoExternoUTF8 caminhoArquivo, conteudo
    
    
    If MsgBox("Deseja abrir o arquivo?", vbYesNo + vbQuestion) = vbYes Then
        
        Shell "explorer " & caminhoArquivo
    End If
    
    
    P_Status ""
    
    dtVinculado.Visible = True
    
    
    
    
    Exit Sub
trata_Erro:
    MostraErro
    
    On Error Resume Next
    dtVinculado.Visible = True
    
    
End Sub


Function BuscaNomeArquivoExterno(Optional filter As String = "*.txt") As String
    Dim terminacao As String
    mdiFrm.cmndlg.filter = filter
    mdiFrm.cmndlg.ShowSave

    terminacao = Right(mdiFrm.cmndlg.FileName, 4)
    If terminacao <> ".txt" And terminacao <> ".csv" Then
        terminacao = ".csv"
    Else
        terminacao = ""
    End If
    BuscaNomeArquivoExterno = mdiFrm.cmndlg.FileName & terminacao
    
    
End Function

