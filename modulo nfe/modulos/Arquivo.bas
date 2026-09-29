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



Function GravaArquivoExternoUTF8(caminho As String, conteudo, Optional abrir As Boolean = False, Optional tipoArquivo As String = "json") As Boolean
    On Error GoTo trata_erro
    Dim ado As New ADODB.Stream
    If caminho = "" Then
        caminho = BuscaNomeArquivoExterno("*." & tipoArquivo)
    End If
    
    ado.Open
    ado.Charset = "UTF-8"
    ado.WriteText Trim(conteudo)
    ado.SaveToFile caminho, adSaveCreateOverWrite
    ado.Close
    GravaArquivoExternoUTF8 = True
    If abrir Then
        shell "explorer " & caminho
    End If
    
trata_erro:
    
    
End Function

Function GravaArquivoExterno(caminho As String, conteudo, Optional abreArquivo As Boolean = False, Optional tipoArquivo As String = "xml") As Boolean
    
    On Error Resume Next
    If caminho = "" Then
        caminho = BuscaNomeArquivoExterno("*." & tipoArquivo)
    End If
    
    
    
    Close #1
    
    Open caminho For Output As #1
    Print #1, conteudo
    
    Close #1
    GravaArquivoExterno = True
    If abreArquivo Then
        shell "explorer " & caminho
    End If
End Function
Function ImprimeEmArquivoExterno(conteudo, buf As Integer) As Boolean
    On Error GoTo trata_erro
    Print #buf, conteudo
    ImprimeEmArquivoExterno = True
    
    Exit Function
trata_erro:
    ImprimeEmArquivoExterno = False
End Function
Function AbreArquivoExternoLeitura(caminho As String, buf As Integer) As Boolean
    On Error GoTo trata_erro
    If FechaArquivoExterno(buf) Then
    End If
    Open caminho For Input As #buf
    AbreArquivoExternoLeitura = True
    Exit Function
trata_erro:
    AbreArquivoExternoLeitura = False
End Function



Function AbreArquivoExterno(caminho As String, buf As Integer) As Boolean
    On Error GoTo trata_erro
    If FechaArquivoExterno(buf) Then
    End If
    Open caminho For Output As #buf
    AbreArquivoExterno = True
    Exit Function
trata_erro:
    AbreArquivoExterno = False
End Function
Function FechaArquivoExterno(buf As Integer) As Boolean
    On Error GoTo trata_erro
    Close #buf
    FechaArquivoExterno = True
    Exit Function
trata_erro:
    FechaArquivoExterno = False
End Function


Function LeArquivoExterno(caminho As String, Optional buf As Integer = 1) As String
    Dim Linha As String
    On Error GoTo trata_erro
    
    Open caminho For Input As #buf
    Do While Not EOF(buf)
        Line Input #buf, Linha
        LeArquivoExterno = LeArquivoExterno & Linha
    Loop
    Close #buf
    Exit Function
trata_erro:
'    MsgBox Err.Description, vbExclamation

End Function
Function ExtraiNumeroPaginasPDF(caminho As String, Optional buf As Integer = 5) As String
    Dim Linha As String
    Dim bLinearized As Boolean
    Dim bKids As Boolean
    On Error GoTo trata_erro
    
    Open caminho For Input As #buf
    Do While Not EOF(buf)
        Line Input #buf, Linha
        'Debug.Print Linha
        If InStr(1, Linha, "/Kids") Then
            bKids = True
        End If
        
        If InStr(1, Linha, "/Linearized") Then
            bLinearized = True
        End If
        
        If InStr(1, Linha, "/Count") Then
            
            ExtraiNumeroPaginasPDF = RetornaValorExpressaoArquivo(Linha, "/Count", "/><")
            FechaArquivoExterno buf
            Exit Function
        
        End If
        If InStr(1, Linha, "/N ") Then
            bLinearized = True
            ExtraiNumeroPaginasPDF = RetornaValorExpressaoArquivo(Linha, "/N", "/><")
            FechaArquivoExterno buf
            Exit Function
            
        End If
        
        
        
    Loop
    ExtraiNumeroPaginasPDF = "999"
    Close #buf
    Exit Function
trata_erro:
    If Err.Number = 53 Then
        MsgBox "Não foi possível abrir o arquivo. Verifique se há caracteres não legíveis no nome ou se o arquivo de fato existe", vbExclamation
    Else
        MsgBox Err.Description, vbExclamation
    End If

End Function




Function LeArquivoExternoUTF8(caminho As String)
    On Error GoTo trata_erro
    Dim ado As New ADODB.Stream
    If caminho = "" Then
        
        caminho = BuscaNomeArquivoExternoAbrir
    End If
      
P_Status "abrindo arquivo... aguarde."
    FechaObjeto ado
    ado.Open
    ado.Charset = "UTF-8"
    ado.LoadFromFile caminho
    LeArquivoExternoUTF8 = ado.ReadText
    
      
    ado.Close
    P_Status ""
    Exit Function
trata_erro:
    P_MostraErro
    
End Function

Sub ExportaGridToCSV(dbgrd As Object, dtVinculado As Object, Optional quebra As String = ";", Optional caminhoArquivo As String = "")
    
    Dim c As Collection
    Set c = New Collection
    Dim conteudo As String
    
    If caminhoArquivo = "" Then
        caminhoArquivo = BuscaNomeArquivoExterno("*.csv")
    
    End If
    
    On Error GoTo trata_erro
    dbgrd.Visible = False
    
    
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
            dbgrd.Visible = False
        
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
        
        shell "explorer " & caminhoArquivo
    End If
    
    
    P_Status ""
    
    dbgrd.Visible = True
    
    
    
    
    Exit Sub
trata_erro:
    P_MostraErro
    
    On Error Resume Next
    dbgrd.Visible = True
    
    
End Sub

Sub ExportaMSFLXGridToCSV(grid As Object, Optional quebra As String = ";", Optional caminhoArquivo As String = "")
    
    Dim c As Collection
    Set c = New Collection
    Dim conteudo As String
    
    If caminhoArquivo = "" Then
        caminhoArquivo = BuscaNomeArquivoExterno("*.csv")
    
    End If
    
    On Error GoTo trata_erro
    grid.Visible = False
    
    
    Dim sPath As String
    Dim strSQL As String
    Dim iQtdColsSelec As Integer
    Dim iCount As Integer
    Dim jCount As Integer
    Dim i As Integer
    
'
    P_Status "Gerando arquivo externo"
    conteudo = ""
    
    
    
    
    
    With grid
        
        
            For iCount = 0 To .Cols - 1
                .row = 0
                .col = iCount
                
                    conteudo = conteudo & .Text & quebra
                
                
            Next
        
            conteudo = conteudo & CrLf

        
        
    End With
    'dtVinculado.Refresh
    With grid
        
        If .rows > 1 Then
            .row = 1
        End If
        For i = 1 To .rows - 1
            
            .row = i
            
            For iCount = 0 To .Cols - 1
                .col = iCount
                
                    conteudo = conteudo & .Text & quebra
                
                
                
            Next
            conteudo = conteudo & CrLf
            
        Next
    End With
    
    GravaArquivoExternoUTF8 caminhoArquivo, conteudo
    
    
    If MsgBox("Deseja abrir o arquivo?", vbYesNo + vbQuestion) = vbYes Then
        
        shell "explorer " & caminhoArquivo
    End If
    
    
    P_Status ""
    
    grid.Visible = True
    
    
    
    
    Exit Sub
trata_erro:
    P_MostraErro
    
    On Error Resume Next
    grid.Visible = True
    
    
End Sub


Function BuscaNomeArquivoExterno(Optional filter As String = "*.txt", Optional nome As String = "") As String
    Dim terminacao As String
    mdiFrm.cmndlg.filter = filter
    terminacao = Replace(filter, "*", "")
    terminacao = Replace(terminacao, ".", "")
    mdiFrm.cmndlg.FileName = nome
    mdiFrm.cmndlg.DefaultExt = terminacao
    mdiFrm.cmndlg.ShowSave
    
    BuscaNomeArquivoExterno = mdiFrm.cmndlg.FileName
    
End Function

Function BuscaNomeArquivoExternoAbrir(Optional filter As String = "*.*") As String
    Dim terminacao As String
    mdiFrm.cmndlg.filter = filter
    'mdiFrm.Refresh
    mdiFrm.cmndlg.ShowOpen

    terminacao = Right(mdiFrm.cmndlg.FileName, 4)
    
    BuscaNomeArquivoExternoAbrir = mdiFrm.cmndlg.FileName
    
    
End Function
Function BuscaNomeArquivoExternoAbrirPasta() As String
    Dim terminacao As String
    
    'mdiFrm.Refresh
    mdiFrm.cmndlg.ShowOpen

    terminacao = Right(mdiFrm.cmndlg.FileName, 4)
    
    BuscaNomeArquivoExternoAbrirPasta = RetornaNomePastaArquivo(mdiFrm.cmndlg.FileName)
    
    
End Function

Function RetornaNomeArquivoResumido(caminhoArquivo As String) As String
    Dim i As Integer
    On Error Resume Next
    i = InStrRev(caminhoArquivo, "\")
    RetornaNomeArquivoResumido = Mid(caminhoArquivo, i + 1)
    
End Function
Function RetornaNomePastaArquivo(caminhoArquivo As String) As String
    Dim i As Integer
    On Error Resume Next
    i = InStrRev(caminhoArquivo, "\")
    RetornaNomePastaArquivo = Left(caminhoArquivo, i)
    
End Function

Function AjustaNomeArquivo(caminho As String) As String
    Dim s As String
    Dim Texto As String
    Dim intCont As Integer
    
    For intCont = 1 To Len(caminho)
        s = Mid(caminho, intCont, 1)
        If s = "\" Then
           Texto = Texto & "\\"
        Else
            Texto = Texto & s
        End If
    Next
    AjustaNomeArquivo = Texto
End Function


Function BuscaNomeArquivoExternoSalvar(Optional filter As String = "*.txt", Optional nome As String) As String
    Dim terminacao As String
    mdiFrm.cmndlg.filter = filter
    mdiFrm.cmndlg.FileName = nome
    mdiFrm.cmndlg.ShowSave

    terminacao = Right(mdiFrm.cmndlg.FileName, 4)
    If terminacao <> ".txt" And terminacao <> ".csv" And terminacao <> ".jpg" Then
        terminacao = Right(filter, 4)
    Else
        terminacao = ""
    End If
    BuscaNomeArquivoExternoSalvar = mdiFrm.cmndlg.FileName & terminacao
    
    
End Function

Function Esperar_Arquivo(path As String, tempo As Long) As String
    'espera até o arquivo existir e ter conteúdo ou dar timeout em milisegundos

    
    Dim Retorno As String
    Dim fileSystem As Object
    Dim timeout As Double
    Dim fileTemp As String
    Dim ts As Object
    Dim conteudo_arquivo As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    
    timeout = Timer + tempo
    Do While Not fileSystem.FileExists(path)
        DoEvents
        If Timer > timeout Then
            Esperar_Arquivo = "Timeout: o comando não terminou dentro do tempo esperado."
            Exit Do
        End If
    Loop
    
    ' Verifica se o conteúdo já foi escrito no arquivo
    timeout = Timer + tempo
    If fileSystem.FileExists(path) Then
        Do While conteudo_arquivo = "" Or Timer > timeout
            DoEvents
            Set ts = fileSystem.opentextfile(path)
            On Error Resume Next
            conteudo_arquivo = ts.ReadAll
        Loop

        Esperar_Arquivo = conteudo_arquivo 'LeArquivoExterno(fileTemp)
        
    End If
    
    
    


End Function

Sub AbrirCSV(grid As Object, Optional caminhoArquivo As String = "")
    Dim conteudo As String
    Dim CSVData As Variant
    Dim l As Integer
    Dim c As Integer
    Dim i As Single
    Dim j As Single
    Dim t As String
    Dim palavra As String
    Dim Linha As String
    Dim k As Single
    Dim strSQL As String
    grid.Clear
    grid.Visible = False
    If caminhoArquivo = "" Then
        caminhoArquivo = BuscaNomeArquivoExternoAbrir("*.csv")
    
    End If
    conteudo = LeArquivoExternoUTF8(caminhoArquivo)
    
    
    l = NLinhasInCsv(conteudo, Chr(10))
    c = NColunasInCsv(conteudo, ";", Chr(10))
    
    ' Preencher o MSHFlexGrid com os dados
    grid.FixedCols = 0
    grid.FixedRows = 0
    grid.rows = l
    grid.Cols = c
    Linha = ""
    palavra = ""
    For i = 1 To Len(conteudo)
        t = Mid(conteudo, i, 1)
        If t = Chr(10) Then
            grid.AddItem Linha, k
             
            Linha = ""
            palavra = ""
            k = k + 1
        ElseIf t = ";" Then
            Linha = Linha & palavra & Chr(9)
            palavra = ""
        Else
            palavra = palavra & t
        End If
        
        
            
            
    Next
    
    grid.FixedRows = 1
    
    grid.Visible = True

End Sub
Function NColunasInCsv(conteudo As String, Optional separator As String = ";", Optional linebreak As String = vbCrLf)
    Dim i As Single
    Dim colunas As Single
    For i = 1 To Len(conteudo)
        If Mid(conteudo, i, 1) = separator Then
            colunas = colunas + 1
            
        End If
        If Mid(conteudo, i, 1) = linebreak Then
            Exit For
        End If
    Next

    NColunasInCsv = colunas

End Function
Function NLinhasInCsv(conteudo As String, Optional linebreak As String = vbCrLf)
    Dim i As Single
    Dim linhas As Single
    Dim x As Single
    For x = 1 To Len(conteudo)
        If Mid(conteudo, x, 1) = linebreak Then
            linhas = linhas + 1
            
        End If
        
    Next
    NLinhasInCsv = linhas


End Function






Function ParseCSV(FileContent As String) As Variant
    Dim Lines() As String
    Dim i As Integer, j As Integer
    Dim Columns() As String
    Dim result() As Variant
    Dim numRows As Integer
    Dim numCols As Integer
    
    
    'for i = 1 to len
    
    
    ' Dividir o conteúdo em linhas
    Lines = Split(FileContent, Chr(10))
    numRows = UBound(Lines) - LBound(Lines) + 1
    
    ' Dividir a primeira linha para determinar o número de colunas
    Columns = Split(Lines(0), ";")
    numCols = UBound(Columns) - LBound(Columns) + 1
    
    ' Redimensionar a matriz de resultados para o número de linhas e colunas
    ReDim result(1 To numRows, 1 To numCols)
    
    ' Dividir cada linha em colunas
    For i = 1 To numRows
        Columns = Split(Lines(i - 1), ";")
        For j = 1 To numCols
            result(i, j) = Columns(j - 1)
        Next j
    Next i
    
    ParseCSV = result
End Function


Sub FillMSHFlexGrid(grid As MSHFlexGrid, data As Variant)
    Dim i As Integer, j As Integer
    Dim numRows As Integer, numCols As Integer
    
    ' Determine o número de linhas e colunas
    numRows = UBound(data, 1)
    numCols = UBound(data, 2)
    
    ' Configurar o MSHFlexGrid
    With grid
        .rows = numRows + 1 ' Adicionar uma linha para os cabeçalhos
        .Cols = numCols
        
        ' Preencher o cabeçalho da primeira linha
        For j = 1 To numCols
            .TextMatrix(0, j - 1) = "Coluna " & j
        Next j
        
        ' Preencher as células do MSHFlexGrid
        For i = 1 To numRows
            For j = 1 To numCols
                .TextMatrix(i, j - 1) = data(i, j)
            Next j
        Next i
        
        ' Ajustar automaticamente o tamanho das colunas
        '.AutoSize = True
    End With
End Sub
Public Function GetFileHash(ByVal filePath As String, Optional ByVal algorithm As String = "MD5") As String
    Dim tempFile As String
    Dim shellCmd As String
    Dim result As String
    Dim fileNumber As Integer
    
    ' Define o caminho do arquivo temporário para armazenar a saída
    tempFile = Environ("TEMP") & "\hash_result.txt"
    
    ' Monta o comando certutil com redirecionamento de saída
    shellCmd = "cmd /c certutil -hashfile """ & filePath & """ " & algorithm & " > """ & tempFile & """"
    
    ' Executa o comando
    shell shellCmd, vbHide
    
    ' Aguarda a criação do arquivo temporário
    Do
        DoEvents
    Loop Until Dir(tempFile) <> ""
    
    ' Lê o arquivo temporário
    result = LeArquivoExterno(tempFile)
    
    ' Extrai apenas o código hash usando uma expressão regular
    ' Aqui está um padrão básico para MD5: [A-Fa-f0-9]{32}
    ' Você pode precisar ajustar o padrão para outros algoritmos de hash.
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "[A-Fa-f0-9]{32}"
    regex.IgnoreCase = True
    
    If regex.Test(result) Then
        ' Obtém a primeira correspondência
        GetFileHash = regex.Execute(result)(0)
    Else
        ' Se nenhum hash for encontrado, retorna uma string vazia ou um valor padrão
        GetFileHash = ""
    End If
    
    ' Apaga o arquivo temporário
    On Error Resume Next
    Kill tempFile
End Function


Sub wait(arquivo As String, milissegundos As Double)
    Dim timeout As Double
    timeout = Timer + milissegundos
    Do While Dir(arquivo) = ""
        DoEvents
        If Timer > timeout Then
            
            Exit Do
        End If
    Loop
End Sub
Function extrai_nome_arquivo(caminho As String, Optional separador = "\") As String
    Dim i As Integer
    i = InStrRev(caminho, separador)
    extrai_nome_arquivo = Mid(caminho, i + 1)
End Function
Function extrai_pasta_arquivo(caminho As String, Optional separador = "\") As String
    Dim i As Integer
    i = InStrRev(caminho, separador)
    extrai_pasta_arquivo = Left(caminho, i)
End Function

Public Function GetFileSize(file_path As String) As Long
    Dim fileNumber As Integer
    On Error GoTo ErrorHandler
    
    ' Open the file for binary input
    fileNumber = FreeFile ' Get a file number that is not currently in use
    Open file_path For Binary Access Read As #fileNumber
    
    ' Get the file size
    GetFileSize = LOF(fileNumber) ' LOF returns the size of the file in bytes
    
    ' Close the file
    Close #fileNumber
    
    Exit Function ' Exit the function normally if no error occurred
    
ErrorHandler:
    ' Handle errors here
    MsgBox "Error opening or reading file: " & Err.Description
    GetFileSize = -1 ' Return -1 or another error code to indicate an error
End Function

