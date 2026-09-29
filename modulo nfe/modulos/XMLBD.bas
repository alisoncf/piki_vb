Attribute VB_Name = "XMLBD"
Public Enum typeValor
    Char = 0
    inteiro = 1
    real = 2
    data = 3
    booleano = 4
End Enum
Type par
    nome As String
    valor As String
    type As typeValor
End Type
Function CabXML() As String
    CabXML = "<?xml version=" & aspas & "1.0" & aspas & " encoding=" & aspas & "ISO-8859-1" & aspas & "?>" + CrLf
End Function
Function TabelaToXml(tabela As String, Optional sql As String = "select * from ", _
    Optional Where As String = "", Optional geraXmlCompleto As Boolean = True, Optional file As Integer = 0) As String
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim xml As String
    Dim campo As String
    Dim valor As String
    Dim tupla As String
    P_Status "Gerando esquema..."
    If geraXmlCompleto Then
        xml = "<?xml version=" & aspas & "1.0" & aspas & " encoding=" & aspas & "ISO-8859-1" & aspas & "?>" + CrLf
    End If
    xml = xml + "<tabela nome=" + aspas + tabela + aspas + ">" + CrLf
    
    strSQL = "show fields from " & tabela
    
    abreObjetoAdo rPesq, strSQL
    
    If file > 0 Then
         
        ImprimeEmArquivoExterno xml, 1
    End If
    
    
    With rPesq
        xml = xml + "<schema>" + CrLf
        Do While Not .EOF
            xml = xml + "<fieldSet name=" + aspas + .Fields(0).Value + aspas + ">" + CrLf
            For i = 0 To .Fields.Count - 1
                
                campo = .Fields(i).Name
                If (.Fields(i).type) = "205" Then
                    valor = TypeName(.Fields(i).Value)
                ElseIf .Fields(i).type = "200" Or .Fields(i).type = "129" Then
                    valor = .Fields(i).Value
                Else
                    MsgBox "tipo não encontrado"
                    valor = .Fields(i).Value
                End If
                tupla = "<" + campo + ">" + valor + "</" + campo + ">" + CrLf
                xml = xml + tupla
                
                
            Next
            xml = xml + "</fieldSet>" + CrLf
            .MoveNext
        Loop
        .Close
        xml = xml + "</schema>" + CrLf
        If file > 0 Then
            ImprimeEmArquivoExterno xml, 1
            xml = ""
        End If
    End With
    P_Status "buscando dados..."
    strSQL = sql & " " & tabela & " " & Where
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        xml = xml & "<records>" + CrLf
        Do While Not .EOF
            tupla = "<record seq=" & aspas & .AbsolutePosition & aspas & ">" + CrLf
            For i = 0 To .Fields.Count - 1
                campo = .Fields(i).Name
                
                valor = Format(.Fields(i).Value)
                tupla = tupla & "<" & campo & ">" & AjustaTextoXML(valor) & "</" & campo & ">" & CrLf
            Next
            tupla = tupla + "</record>" + CrLf
            xml = xml + tupla
            If .AbsolutePosition Mod 100 = 0 Then
                P_Status "Gerando registro " & .AbsolutePosition & "/" & .RecordCount
            End If
            If file > 0 Then
                ImprimeEmArquivoExterno xml, 1
                xml = ""
            End If
            .MoveNext
        Loop
        .Close
    End With
    xml = xml & "</records>" + CrLf
    xml = xml + "</tabela>"
    
    TabelaToXml = xml
    P_Status ""

End Function


Sub GeraXmlTabela(tabela As String, Optional sql As String = "select * from ", Optional Where As String = "")
    Dim sPath As String
    
    Dim xml As String
    
    
    xml = TabelaToXml(tabela, sql, Where)
    
    GravaArquivoExterno "", xml, True
    
    

End Sub
Function entreAspas(termo) As String
    entreAspas = aspas & termo & aspas
End Function
Function ValorJSON(termo As String, valor As String, Optional valorEntreAspas As Boolean = True) As String
    valor = TiraEnterTab(valor)
    ValorJSON = entreAspas(termo) & ": " & IIf(valorEntreAspas, entreAspas(valor), valor)
End Function

Function ConjuntoJSONRecord(record As ADODB.record) As String
    Dim s As String
    s = "{"
End Function
Function FieldToPar(fil As ADODB.Field) As par
    FieldToPar.nome = fil.Name
    FieldToPar.valor = fil.Value
    If fil.type = adBigInt Or fil.type = adCurrency Or fil.type = adDecimal Or _
        fil.type = adDouble Or fil.type = adInteger Or fil.type = adNumeric Or fil.type = adSingle _
        Or fil.type = adSmallInt Then
        FieldToPar.type = inteiro
    ElseIf fil.type = adDate Or fil.type = adDBTime Then
        FieldToPar.type = data
    Else
        FieldToPar.type = Char
    End If
        
    
End Function
Function FieldsToPar(rec As ADODB.Fields) As par()
    ReDim p(rec.Count) As par
    Dim i As Integer
    Dim t As typeValor
    For i = 1 To rec.Count
    
        p(i).nome = rec(i - 1).Name
        p(i).valor = TiraNull(rec(i - 1).Value)
        t = rec(i - 1).type
        If t = adBigInt Or t = adCurrency Or t = adDecimal Or _
            t = adDouble Or t = adInteger Or t = adNumeric Or t = adSingle _
            Or t = adSmallInt Then
            p(i).type = inteiro
        ElseIf t = adDate Or t = adDBTime Then
            p(i).type = data
        Else
            p(i).type = Char
        End If
    Next
    FieldsToPar = p
End Function

Function ConjuntoJSON(p() As par, Optional iSpace As Integer = 0, Optional objeto As String = "", _
    Optional CamposNaoGerar As String = "") As String
    Dim i As Integer
    Dim s As String
    Dim t As Integer
    Dim v As String
    Dim d As String
    t = UBound(p)
    If objeto <> "" Then
        s = getTab(iSpace) & entreAspas(objeto) & ": {" & CrLf
    Else
        s = getTab(iSpace) & "{" & CrLf
    End If
    
    For i = 1 To t
        If p(i).type = inteiro Or p(i).type = real Then
            v = aspas & Trim(F_GravaDec(p(i).valor)) & aspas
        ElseIf p(i).type = data Then
            v = substituir_aspas(Trim(F_GravaData(p(i).valor)))
        Else
            d = seForDataFormataYYYYMMDD(p(i).valor)
            If d <> "" Then
                v = d
            Else
                v = Replace(p(i).valor, """", "\""")
            End If
        End If
        s = s & getTab(iSpace + 2) & ValorJSON(p(i).nome, v, IIf(p(i).type <> inteiro And p(i).type <> real, True, False))
        If i < t Then
            s = s & ", " & CrLf
        End If
        
    Next
    s = s & CrLf & getTab(iSpace + 1) & "}"
    ConjuntoJSON = s
End Function

Function TestarConj() As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim s As String
    strSQL = "select * from tab_funcionarios"
    abreObjetoAdo rPesq, strSQL
    Dim p() As par
    Dim i As Integer
    Dim r As ADODB.Fields
    With rPesq
        Do While Not .EOF
            Set r = rPesq.Fields
        
            s = s & CrLf & ConjuntoJSON(FieldsToPar(r)) + "," + CrLf
            
            .MoveNext
        Loop
        .Close
    End With
    Debug.Print s
    
End Function
Function getTab(i As Integer) As String
    getTab = String(i, Chr(9))
    
End Function
Function TabToJson(tabela As String, Optional NomeObjeto As String = "", _
    Optional caminhoArquivo As String = "", Optional iTab As Integer = 0, _
    Optional buffer As Integer = 0, Optional inicioJSON As Boolean = True, _
    Optional listaCamposNaoGerar As String = "", Optional geraArquivo As Boolean = True) As String
    Dim strSQL As String
    strSQL = "select * from " & tabela
    
    TabToJson = RecordSetToJson(strSQL, NomeObjeto, caminhoArquivo, iTab, buffer, inicioJSON, _
    listaCamposNaoGerar, geraArquivo)
    
End Function


Function RecordSetToJson(strSQL As String, Optional NomeObjeto As String = "", _
    Optional caminhoArquivo As String = "", Optional iTab As Integer = 0, _
    Optional buffer As Integer = 0, Optional inicioJSON As Boolean = True, _
    Optional listaCamposNaoGerar As String = "", Optional geraArquivo As Boolean = True) As String
    
    Dim rPesq As New ADODB.Recordset
    Dim s As String
    Dim retorno As String
    Dim espaco As String
    Dim espaco2 As String
    Dim abertura As String
    Dim buf As Integer
    Dim fechaArquivo As Boolean
    
    If inicioJSON = True Then
        abertura = "{"
    Else
        abertura = ","
    End If
    espaco = getTab(iTab)
    espaco2 = getTab(iTab + 1)
    If geraArquivo Then
        'P_Status "Gerando arquivo... aguarde..."
    End If
    If buffer = 0 Then
        If caminhoArquivo = "" And geraArquivo Then
            caminhoArquivo = BuscaNomeArquivoExterno("*.json")
        
        End If
        
    End If
    If NomeObjeto <> "" Then
        s = espaco & abertura & CrLf & espaco2 & entreAspas(NomeObjeto) & ": ["
      
    End If
    If geraArquivo = False Then
        retorno = s
    End If
    
    abreObjetoAdo rPesq, strSQL
    
    If buffer = 0 And geraArquivo Then
        AbreArquivoExterno caminhoArquivo, 1
        buf = 1
        fechaArquivo = True
    Else
        buf = buffer
    End If
    If geraArquivo Then
        ImprimeEmArquivoExterno s, buf
    End If
    Dim p() As par
    Dim i As Integer
    Dim r As ADODB.Fields
    With rPesq
        Do While Not .EOF
            Set r = rPesq.Fields
            s = ""
            s = s & CrLf & espaco2 & ConjuntoJSON(FieldsToPar(r), iTab)
            If .AbsolutePosition < .RecordCount Then
                s = s & espaco & "," & CrLf
            End If
            If geraArquivo Then
                If .AbsolutePosition Mod 100 = 0 Then
                    'P_Status "gerando arquivo " & .AbsolutePosition & " de " & .RecordCount
                End If
                ImprimeEmArquivoExterno s, buf
            Else
                retorno = retorno & s
            End If
            .MoveNext
            
        Loop
        .Close
    End With
    s = ""
    If NomeObjeto <> "" Then
        s = s & CrLf & espaco & "]" & CrLf
    End If
    If inicioJSON Then
        s = s & espaco & "}"
    End If
    If geraArquivo Then
        ImprimeEmArquivoExterno s, buf
        If fechaArquivo Then
            FechaArquivoExterno buf
        End If
        retorno = s
    Else
       
    End If
    
'    P_Status ""
    RecordSetToJson = retorno
End Function
Function seForDataFormataYYYYMMDD(valor As String) As String
    If Len(valor) <> 10 Then
        Exit Function
    End If
    If Mid(valor, 3, 1) <> "/" Or Mid(valor, 6, 1) <> "/" Then
        Exit Function
    End If
    seForDataFormataYYYYMMDD = Format(valor, "YYYY-MM-DD")
End Function


Function substituir_aspas(Texto As String, Optional por As String = "'") As String
    substituir_aspas = Replace(Texto, """", por)
End Function

Function TiraEnterTab(originalString) As String
    Dim cleanedString
    
    cleanedString = Replace(originalString, vbCrLf, "") ' Remove <enter>
    cleanedString = Replace(cleanedString, vbTab, "")   ' Remove <tab>
    cleanedString = Replace(cleanedString, vbCr, "")    ' Remove carriage return (apenas se presente)
    cleanedString = Replace(cleanedString, vbLf, "")
    
    TiraEnterTab = cleanedString
End Function
