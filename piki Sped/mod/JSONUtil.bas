Attribute VB_Name = "JSONUtil"

Function busca_cep(cep As String) As endereco
    Dim json As String
    json = HTTPUtil.sendGet("https://viacep.com.br/ws/" & F_SoNumero(cep) & "/json/")

    busca_cep.bairro = GetValorJson("bairro", json)
    busca_cep.complemento = GetValorJson("complemento", json)
    busca_cep.cep = GetValorJson("cep", json)
    busca_cep.localidade = GetValorJson("localidade", json)
    busca_cep.Logradouro = GetValorJson("logradouro", json)
    busca_cep.ibge = GetValorJson("ibge", json)

End Function





Function GetValorJson(Chave As String, json As String) As String
    Dim i As Integer
    Dim iPA As Integer
    Dim iUA As Integer
    Dim i2P As Integer
    i = InStr(1, json, aspas & Chave & aspas)
    
    If i > 0 Then
        i2P = InStr(i + Len(Chave) + 2, json, ":")
        If Mid(json, i2P + 1, 4) = "null" Then
            GetValorJson = ""
        Else
            iPA = InStr(i2P, json, aspas)
            iUA = InStr(iPA + 1, json, aspas)
            GetValorJson = Mid(json, iPA + 1, iUA - iPA - 1)
        End If
        
        
    End If
    
End Function

Function AspasSimplesParaAspasDuplas(json As String) As String
    AspasSimplesParaAspasDuplas = Replace(json, "'", """")
End Function

Function AspasDuplasParaAspasSimples(json As String) As String
    AspasDuplasParaAspasSimples = Replace(json, """", "'")
End Function




Function ParseJsonIntoList(jsonStr) As Collection
    Dim tempStr As String
    Dim listaDeItens() As String
    Dim item As Variant
    Dim result As New Collection
    Dim colcheteI As Integer
    
    tempStr = Mid(jsonStr, 2, Len(jsonStr) - 2)
    
    
    tempStr = Replace(tempStr, "}, {", "},{")
    colcheteI = InStr(1, tempStr, "[")
    colcheteF = InStr(1, tempStr, "]")
    If colcheteI > 1 Then
        tempStr = Mid(tempStr, colcheteI + 1)
    End If
    If colcheteF > 1 Then
        tempStr = Mid(tempStr, 1, colcheteF)
    End If
    listaDeItens = Split(tempStr, "},{")

    
    For Each item In listaDeItens
    
        item = "" & item & ""
        result.Add item
    Next item

    
    Set ParseJsonIntoList = result
End Function

Function RetornaNodeJson(node As String, identificador, json)

    Dim colcheteI As Long
    Dim colcheteF As Long
    Dim i As Long
    i = InStr(1, json, node)
    If i > 0 Then
        colcheteI = InStr(i, json, "[")
        If colcheteI > i Then
            colcheteF = InStr(colcheteI + 1, json, "]")
        End If
        
        If colcheteI > 0 And colcheteF > 0 Then
            RetornaNodeJson = Mid(json, colcheteI + 1, (colcheteF - colcheteI - 1))
        End If
        
    End If


End Function

Function RetornaLinhasJson(json) As Collection
    'espera sempre {"objeto":"valor"},{"objeto":"valor"},
    Dim partes() As String
    Dim i As Integer
    Dim c As New Collection
    partes = Split(Trim(json), "{")
    For i = LBound(partes) To UBound(partes)
        partes(i) = Replace(partes(i), "},", "")
        If Len(partes(i)) > 3 Then
            c.Add partes(i)
        End If
    Next
    Set RetornaLinhasJson = c
End Function
Function ExtrairValorObjetoJSON(objeto, Chave As String, _
    Optional retira_aspas As Boolean = True, Optional transforma_double As Boolean = False) As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim valor As String
    Dim o As Integer
    objeto = Replace(objeto, ":False", ":" & entreAspas("f"))
    objeto = Replace(objeto, ":True", ":" & entreAspas("t"))
    objeto = Replace(objeto, ":false", ":" & entreAspas("f"))
    objeto = Replace(objeto, ":true", ":" & entreAspas("t"))
    objeto = Replace(objeto, "\", "")
    
    i = InStr(1, objeto, entreAspas(Trim(Chave)))
    If i = 0 Then
        ExtrairValorObjetoJSON = ""
        Exit Function
    End If
    
    If i > 0 Then
        j = InStr(i + Len(Chave), objeto, ":")
    End If
    If LCase(Mid(objeto, j + 1, 4)) = "null" Then
        ExtrairValorObjetoJSON = "null"
        Exit Function
    End If
    
    If j > 0 Then
        k = InStr(j + 1, objeto, """")
    End If
    If k > 0 Then
        l = InStr(k + 1, objeto, """")
    End If
    valor = Mid(objeto, j + 1, l - k + 2)
    
    If Right(valor, 1) = "," Then
        valor = Mid(valor, 1, Len(valor) - 1)
    End If
    If retira_aspas Then
        
        valor = Trim(Replace(valor, """", ""))
        
    End If
    If transforma_double Then
        valor = Trim(Replace(valor, ".", ","))
    End If
    ExtrairValorObjetoJSON = valor
End Function

Public Function CollectionToRecordset(col As Collection) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim obj As Object
    Dim prop As Variant
    Dim tipo As VbVarType
    Dim i As Long

    ' Cria o recordset desconectado
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic

    ' Pega o primeiro objeto da coleção para descobrir os campos
    If col.Count = 0 Then
        Set CollectionToRecordset = rs
        Exit Function
    End If

    Set obj = col(1)

    ' Cria os campos do recordset dinamicamente
    For Each prop In obj
        tipo = VarType(prop)
        Select Case tipo
            Case vbString
                rs.Fields.Append propName(prop), adVarChar, 255
            Case vbInteger, vbLong
                rs.Fields.Append propName(prop), adInteger
            Case vbBoolean
                rs.Fields.Append propName(prop), adBoolean
            Case vbDate
                rs.Fields.Append propName(prop), adDate
            Case vbDouble, vbSingle
                rs.Fields.Append propName(prop), adDouble
            Case Else
                rs.Fields.Append propName(prop), adVariant
        End Select
    Next

    rs.Open

    ' Preenche o recordset com os dados da collection
    For i = 1 To col.Count
        Set obj = col(i)
        rs.AddNew
        For Each prop In obj
            rs.Fields(propName(prop)).Value = prop
        Next
        rs.Update
    Next

    rs.MoveFirst
    Set CollectionToRecordset = rs
End Function

Private Function propName(p As Variant) As String
    ' Dummy function para nomes de propriedades
    ' Substitua por código que saiba os nomes corretos
    propName = "Campo" & Rnd() * 10000
End Function



Function ParseJsonIntoListaDeColunas(jsonStr) As Collection
    Dim tempStr As String
    Dim listaDeItens() As String
    Dim listaDeColunas() As String
    Dim item As Variant
    Dim result As New Collection
    Dim colcheteI As Integer
    Dim objeto As String
    tempStr = Mid(jsonStr, 2, Len(jsonStr) - 2)
    
    
    tempStr = Replace(tempStr, "}, {", "},{")
    colcheteI = InStr(1, tempStr, "[")
    colcheteF = InStr(1, tempStr, "]")
    If colcheteI > 1 Then
        tempStr = Mid(tempStr, colcheteI + 1)
    End If
    If colcheteF > 1 Then
        tempStr = Mid(tempStr, 1, colcheteF)
    End If
    listaDeItens = Split(tempStr, "},{")
    
    item = listaDeItens(1)
    
    listaDeColunas = Split(tempStr, """:")
    

    
    Set ParseJsonIntoListaDeColunas = result
End Function

Public Function JsonParaListaProdutos(jsonString As String) As Collection
    Dim jsonArray As Collection
    Dim item As Variant
    Dim prop As Variant
    Dim produto As ProdutoClass
    Dim lista As New Collection

    ' Parse JSON
    Set jsonArray = JsonConverter.ParseJson(jsonString)

    ' Para cada item no JSON
    For Each item In jsonArray
        Set produto = New ProdutoClass
        
        ' Tenta preencher propriedades do objeto
        For Each prop In item.Keys
            On Error Resume Next ' Ignora propriedades que não existem
            CallByName produto, prop, VbLet, item(prop)
            On Error GoTo 0
        Next prop

        lista.Add produto
    Next item

    Set JsonParaListaProdutos = lista
End Function

