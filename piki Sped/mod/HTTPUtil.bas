Attribute VB_Name = "HTTPUtil"
Public Function sendGetCurl(url As String, Optional bearerToken As String = "", _
    Optional ContentType As String = "") As String
    Dim s As String
    Dim retorno As String
    Dim fileSystem As Object
    Dim timeout As Double
    Dim fileTemp As String
    Dim ts As Object
    Dim conteudo_arquivo As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Randomize
    
    
    fileTemp = App.path & "\" & F_SoNumero(Now) + Int((100 * Rnd) + 1) & ".tmp"
    
    s = " curl -X GET " & aspas & url & aspas
    s = s & " -H " & aspas & "User-Agent: Piki App." & aspas
    s = s & " -H " & aspas & "Accept: */*" & aspas
    If bearerToken <> "" Then
        s = s & " -H " & aspas & "Authorization: Bearer " & bearerToken & aspas
    End If
    If ContentType <> "" Then
        s = s & " -H Content-Type: " & ContentType
    End If
    
    
    
    
   shell "cmd.exe /c  " & s & " > " & fileTemp, vbHide
    
    
    
    timeout = Timer + 5000
    Do While Not fileSystem.FileExists(fileTemp)
        DoEvents
        If Timer > timeout Then
            sendGetCurl = "Timeout: o comando não terminou dentro do tempo esperado."
            Exit Do
        End If
    Loop
    
    ' Verifica se o conteúdo já foi escrito no arquivo
    timeout = Timer + 5000
    If fileSystem.FileExists(fileTemp) Then
        Do While conteudo_arquivo = "" Or Timer > timeout
            DoEvents
            'Set ts = fileSystem.opentextfile(fileTemp)
            On Error Resume Next
            conteudo_arquivo = LeArquivoExterno(fileTemp)
        Loop

        retorno = conteudo_arquivo 'LeArquivoExterno(fileTemp)
        If retorno = "" Then
            MsgBox "arquivo vazio"
        End If
        'ts.Close
        FechaArquivoExterno 1
        fileSystem.DeleteFile fileTemp
        Kill fileTemp
    End If
    
    sendGetCurl = retorno
    






End Function
Public Function sendPost2(url As String, DataRaw As String, Optional bearerToken As String = "", _
    Optional ContentType As String = "application/x-www-form-urlencoded") As String
    Dim s
    Dim retorno As String
    Dim fileSystem As Object
    Dim timeout As Double
    Dim fileTemp As String
    Dim ts As Object
    Dim conteudo_arquivo As String
    Dim tmp_folder As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Randomize
    
    tmp_folder = App.path & "\tmp"
    If Dir(tmp_folder, vbDirectory) = "" Then
        MkDir tmp_folder
    End If
    fileTemp = tmp_folder & "\" & F_SoNumero(Now) + Int((100 * Rnd) + 1) & ".tmp"
    
    s = "curl -k " & url & " -X POST "
    s = s & " -H " & aspas & "User-Agent: App Piki" & aspas
    If bearerToken <> "" Then
        s = s & " -H " & aspas & "Authorization: Bearer " & bearerToken & aspas
    End If
    If ContentType <> "" Then
        s = s & " -H " & aspas & "Content-Type: " & ContentType & aspas
    End If
    s = DuplicaAD(s)
    s = s & " --data-raw " & aspas & DataRaw & aspas
    
    s = TiraEnterTab(s)
    
    shell "cmd.exe /c  " & s & " > " & fileTemp, vbHide
    
    Dim command As String
    'command = "curl -k https://vps49416.publiccloud.com.br/api/doc/inserir -X POST -H ""User-Agent: App Piki"" -H ""Authorization: Bearer kbsjYVILGsY5HV5aYSBvMHEEKzUabKGJKqSYa6Z3jmagd9f8TEKLUuxJjE950jAykeXet1"" -H ""Content-Type: application/x-www-form-urlencoded"" --data-raw ""num_doc=0173092&valor_doc=160,65&historico_doc=ACERTO CONSIGN. FRANCISCO DAS CHAGAS VIEIRA&cod_emp=01&cod_cli=4&cod_centro_custo=1&listaParcelas=[{""seq_itd"": ""0760250"", ""seq_doc"": ""0191984"", ""valor_itd"":  160.65, ""reman_itd"":  160.65, ""vencini_itd"": ""15/06/2024"", ""venc_itd"": ""15/06/2024"", ""parcela_itd"":  1, ""juros_itd"":  0, ""desconto_itd"":  0, ""baixa_itd"": """", ""boleto_itd"": """", ""operacao_itd"": ""2"", ""cod_cob"": ""000"", ""cod_tip"": ""2"", ""impressao_itd"": ""1"", ""prejuizo_itd"": ""N"", ""marcado_itd"": """", ""cod_emp"":  1}]]"""
    
    timeout = Timer + 10
    Do While Not fileSystem.FileExists(fileTemp)
        DoEvents
        If Timer > timeout Then
            sendPost2 = "Timeout: o comando não terminou dentro do tempo esperado."
            Exit Do
        End If
    Loop
    
    ' Verifica se o conteúdo já foi escrito no arquivo
    timeout = Timer + 1
    If fileSystem.FileExists(fileTemp) Then
        Do While conteudo_arquivo = "" Or Timer < timeout
            DoEvents
            'Set ts = fileSystem.opentextfile(fileTemp)
            On Error Resume Next
            conteudo_arquivo = LeArquivoExterno(fileTemp)
        Loop

        retorno = conteudo_arquivo 'LeArquivoExterno(fileTemp)
        If retorno = "" Then
            MsgBox "arquivo de retorno vazio"
        End If
        'ts.Close
        FechaArquivoExterno 1
        fileSystem.DeleteFile fileTemp
        If Dir(fileTemp) Then
            Kill fileTemp
        End If
    End If
    
    sendPost2 = retorno
    



End Function

Public Function sendPostFile(url As String, file As String, Optional bearerToken As String = "", _
    Optional ContentType As String = "application/x-www-form-urlencoded") As String
    Dim s
    Dim retorno As String
    Dim fileSystem As Object
    Dim timeout As Double
    Dim fileTemp As String
    Dim ts As Object
    Dim conteudo_arquivo As String
    Dim tmp_folder As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Randomize
    
    tmp_folder = App.path & "\tmp"
    If Dir(tmp_folder, vbDirectory) = "" Then
        MkDir tmp_folder
    End If
    fileTemp = tmp_folder & "\" & F_SoNumero(Now) + Int((100 * Rnd) + 1) & ".tmp"
    
    s = "curl -k " & url & " -X POST "
    s = s & " -H " & aspas & "User-Agent: App Piki" & aspas
    If bearerToken <> "" Then
        s = s & " -H " & aspas & "Authorization: Bearer " & bearerToken & aspas
    End If
    If ContentType <> "" Then
        s = s & " -H " & aspas & "Content-Type: " & ContentType & aspas
    End If
    s = DuplicaAD(s)
    s = s & " --data-binary @" & file
    
    s = TiraEnterTab(s)
    
    shell "cmd.exe /c  " & s & " > " & fileTemp, vbHide
    
    Dim command As String
    'command = "curl -k https://vps49416.publiccloud.com.br/api/doc/inserir -X POST -H ""User-Agent: App Piki"" -H ""Authorization: Bearer kbsjYVILGsY5HV5aYSBvMHEEKzUabKGJKqSYa6Z3jmagd9f8TEKLUuxJjE950jAykeXet1"" -H ""Content-Type: application/x-www-form-urlencoded"" --data-raw ""num_doc=0173092&valor_doc=160,65&historico_doc=ACERTO CONSIGN. FRANCISCO DAS CHAGAS VIEIRA&cod_emp=01&cod_cli=4&cod_centro_custo=1&listaParcelas=[{""seq_itd"": ""0760250"", ""seq_doc"": ""0191984"", ""valor_itd"":  160.65, ""reman_itd"":  160.65, ""vencini_itd"": ""15/06/2024"", ""venc_itd"": ""15/06/2024"", ""parcela_itd"":  1, ""juros_itd"":  0, ""desconto_itd"":  0, ""baixa_itd"": """", ""boleto_itd"": """", ""operacao_itd"": ""2"", ""cod_cob"": ""000"", ""cod_tip"": ""2"", ""impressao_itd"": ""1"", ""prejuizo_itd"": ""N"", ""marcado_itd"": """", ""cod_emp"":  1}]]"""
    
    timeout = Timer + 10
    Do While Not fileSystem.FileExists(fileTemp)
        DoEvents
        If Timer > timeout Then
            sendPostFile = "Timeout: o comando não terminou dentro do tempo esperado."
            Exit Do
        End If
    Loop
    
    ' Verifica se o conteúdo já foi escrito no arquivo
    timeout = Timer + 1
    If fileSystem.FileExists(fileTemp) Then
        Do While conteudo_arquivo = "" Or Timer < timeout
            DoEvents
            'Set ts = fileSystem.opentextfile(fileTemp)
            On Error Resume Next
            conteudo_arquivo = LeArquivoExterno(fileTemp)
        Loop

        retorno = conteudo_arquivo 'LeArquivoExterno(fileTemp)
        If retorno = "" Then
            MsgBox "arquivo de retorno vazio"
        End If
        'ts.Close
        FechaArquivoExterno 1
        fileSystem.DeleteFile fileTemp
        If Dir(fileTemp) Then
            Kill fileTemp
        End If
    End If
    
    sendPostFile = retorno
    



End Function

Public Function sendGet(url As String, Optional bearerToken As String = "", _
    Optional ContentType As String = "") As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    On Error GoTo erroHandler

    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "App Piki"

    If bearerToken <> "" Then
        http.setRequestHeader "Authorization", "Bearer " & bearerToken
    End If

    If ContentType <> "" Then
        http.setRequestHeader "Content-Type", ContentType
    End If

    http.sEnd

    sendGet = http.responseText
    Exit Function

erroHandler:
    sendGet = "Erro: " & Err.Description
End Function


Public Function sendPost(url As String, conteudo As String, Optional bearerToken As String = "", _
    Optional ContentType As String = "application/x-www-form-urlencoded", Optional encodeContent As Boolean = False) As String
    Dim s
    Dim retorno As String
    Dim fileSystem As Object
    Dim timeout As Double
    Dim fileTemp As String
    Dim ts As Object
    Dim conteudo_arquivo
    Dim tmp_folder As String
    Dim http As Object
    Dim lista As String
    Dim encodedData
    Dim pathLog As String
    Dim fileLog As String
    Dim scaped As String
    On Error GoTo trata_Erro
    If encodeContent Then
        encodedData = URLEncode(conteudo)
    Else
        encodedData = conteudo
    End If
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "POST", url, False
    http.setRequestHeader "User-Agent", "App Piki"
    http.setRequestHeader "Authorization", "Bearer " & bearerToken
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    http.sEnd encodedData
    
    Dim responseBytes() As Byte
    Dim responseText As String

    responseBytes = http.responseBody
'    responseText = StrConv(responseBytes, vbUnicode)

    scaped = UnicodeUnescape(http.responseText)
    sendPost = UnicodeUnescape(http.responseText)
    pathLog = App.path & "\envio_post"
    If Dir(pathLog, vbDirectory) = "" Then
        MkDir (pathLog)
    End If
    fileLog = F_SoNumero(Now) & ".txt"
    
    GravaArquivoExternoUTF8 App.path & "\post.txt", sendPost
    
    If InStr(1, sendPost, "sucesso") = 0 Then
        GravaArquivoExternoUTF8 App.path & "\err_" & fileLog, sendPost & CrLf & conteudo
    End If
    
    
trata_Erro:
    
End Function
Public Function UnicodeUnescape(s As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim m As Object
    Dim result As String
    Dim i As Long
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "\\u([0-9a-fA-F]{4})"
    
    result = s
    Set matches = regex.Execute(s)
    
    For i = matches.Count - 1 To 0 Step -1
        Set m = matches.item(i)
        Dim unicodeChar As String
        unicodeChar = ChrW("&H" & m.SubMatches(0))
        result = Left(result, m.FirstIndex) & unicodeChar & Mid(result, m.FirstIndex + m.Length + 1)
    Next i

    UnicodeUnescape = result
End Function

Function Base64EncodeUTF8(text As String) As String
    Dim UTF8() As Byte
    Dim xml As Object
    Dim node As Object

    ' Converte UTF-16 para UTF-8
    UTF8 = TextToUTF8(text)

    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = UTF8

    Base64EncodeUTF8 = Replace(node.text, vbLf, "")
End Function

Private Function TextToUTF8(str As String) As Byte()
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    stream.type = 2  ' text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText str
    stream.Position = 0

    stream.type = 1  ' binary
    TextToUTF8 = stream.Read

    stream.Close
End Function




Function Base64Encode(text As String) As String
    Dim arr() As Byte
    Dim xml As Object
    Dim node As Object

    ' Texto do VB6 está em Unicode (UTF-16): convertemos para ANSI/UTF-8 conforme necessidade
    arr = StrConv(text, vbFromUnicode)

    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = arr

    Base64Encode = Replace(node.text, vbLf, "") ' remove line breaks se houver
End Function

Sub exclui_arquivo(caminho As String)
    shell "cmd.exe /c del " & caminho
    
End Sub
Public Function URLEncode(str As String) As String
    Dim i As Long
    Dim charCode As Integer
    Dim result As String
    Dim ch As String

    result = ""

    For i = 1 To Len(str)
        ch = Mid$(str, i, 1)
        charCode = Asc(ch)

        Select Case charCode
            Case 48 To 57, 65 To 90, 97 To 122 ' 0-9, A-Z, a-z
                result = result & ch
            Case 32
                result = result & "+"
            Case Else
                result = result & "%" & Right$("0" & Hex(charCode), 2)
        End Select
    Next i

    URLEncode = result
End Function


Function TiraEnterTab(originalString) As String
    Dim cleanedString
    
    cleanedString = Replace(originalString, vbCrLf, "") ' Remove <enter>
    cleanedString = Replace(cleanedString, vbTab, "")   ' Remove <tab>
    cleanedString = Replace(cleanedString, vbCr, "")    ' Remove carriage return (apenas se presente)
    cleanedString = Replace(cleanedString, vbLf, "")
    
    TiraEnterTab = cleanedString
End Function

Function DuplicaAD(originalString) As String
    Dim cleanedString
    
    cleanedString = Replace(originalString, aspas, """") ' Remove <enterz
    
    DuplicaAD = cleanedString
End Function
Function DuplicaADCB(originalString As String) As String
    Dim cleanedString As String
    
    cleanedString = Replace(originalString, aspas, "\" & """") ' Remove <enter>
    
    DuplicaADCB = cleanedString
End Function

