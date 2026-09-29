Attribute VB_Name = "API"
Public Const bearer_token = "kbsjYVILGsY5HV5aYSBvMHEEKzUabKGJKqSYa6Z3jVmagd9f8TEKLUuxJjE950jAykeXet1"
Public Enum tipo_fila
    cliente = 1
    documento = 2
    produtos = 3
End Enum


Function APIInserirDoc(seq_doc As String, Optional mostra_status As Boolean = False, Optional atuliza_cliente As Boolean = True) As String
    Dim url As String
    Dim parc As String
    Dim d As doc.doc
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    If mostra_status Then
    On Error GoTo trata_Erro
    
        P_Status "atualizando cliente"
    End If
    url = "https://piki.app.br/api/desktop_doc/alterar"
    
    d = doc.F_buscaDoc(seq_doc)
    
    If atuliza_cliente Then
    
        APIInserirCliente d.cod_ter
    End If
    If mostra_status Then
        P_Status
    End If
    parc = doc.BuscaItdocJson(seq_doc)
    

    data_raw = "num_doc=" & d.num_doc
    data_raw = data_raw & "&valor_doc=" & d.valor_doc
    data_raw = data_raw & "&historico_doc=" & d.historico_doc
    data_raw = data_raw & "&data_doc=" & Format(d.data_doc, "yyyy-mm-dd")
    data_raw = data_raw & "&cod_emp=" & strCod_Emp
    data_raw = data_raw & "&seq_doc=" & d.seq_doc
    data_raw = data_raw & "&cpf=" & F_SoNumero(d.cpf_ter)
    data_raw = data_raw & "&cod_centro_custo=1"
    data_raw = DuplicaAD(data_raw)
    parcelas = "&listaParcelas=[" & parc & "]"
    'parcelas = DuplicaADCB(parcelas)
    
    
    ret = sendPost(url, data_raw & parcelas, bearer_token)
    
    APIInserirDoc = ret
    If mostra_status Then
        P_Status
    End If
Exit Function
trata_Erro:

    
End Function

Function APIExcluirDocPeloNumero(num_doc As String, cod_emp As String) As String
    Dim url As String
    Dim parc As String
    Dim d As doc.doc
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    
    
    url = "https://piki.app.br/api/desktop_doc/deletar"

    data_raw = "num_doc=" & d.num_doc
    data_raw = data_raw & "&cod_emp=" & strCod_Emp
    data_raw = data_raw & "&data_doc=" & d.data_doc
    data_raw = DuplicaAD(data_raw)
    parcelas = "&listaParcelas=[" & "]"
    'parcelas = DuplicaADCB(parcelas)
    
    
    ret = sendPost(url, data_raw & parcelas, bearer_token)
    
    APIExcluirDocPeloNumero = ret
    If mostra_status Then
        P_Status
    End If
End Function

Function APIInserirCliente(cod_cli As String) As String
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    'P_Status "enviando cliente"
    
    url = "https://piki.app.br/api/desktop_clientes/alterar"
    
    obj = Clientes.F_BuscaClienteCodigo(cod_cli)
    
    
    

    data_raw = "cpf=" & F_SoNumero(obj.cnpj_Cli)
    data_raw = data_raw & "&nome=" & obj.nome_cli
    data_raw = data_raw & "&registro_geral=" & obj.ie_Cli
    data_raw = data_raw & "&sexo=" & obj.sexo_cli
    data_raw = data_raw & "&data_nasc=" & Format(obj.datanasc_cli, "yyyy-mm-dd")
    data_raw = data_raw & "&email_cli=" & obj.email_cli
    data_raw = data_raw & "&endereco=" & obj.Logradouro
    data_raw = data_raw & "&complemento=" & obj.complemento
    data_raw = data_raw & "&bairro=" & obj.bairro_cli
    data_raw = data_raw & "&mae=" & obj.mae_cli
    data_raw = data_raw & "&pai=" & obj.pai_cli
    data_raw = data_raw & "&dataalt_cli=" & Format(obj.dataalt_cli, "yyyy-mm-dd")
    data_raw = DuplicaAD(data_raw)
    
    
    ret = sendPost(url, data_raw, bearer_token)
    
    APIInserirCliente = ret
    'P_Status
End Function
Function APIAtualizarCliente(cod_cli As String, id_cli As String) As String
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    'P_Status "enviando cliente"
    
    url = "https://piki.app.br/api/desktop_clientes/alterar"
    
    obj = Clientes.F_BuscaClienteCodigo(cod_cli)
    
    
    

    data_raw = "cpf=" & F_SoNumero(obj.cnpj_Cli)
    data_raw = data_raw & "&cod_cli=" & id_cli
    data_raw = data_raw & "&nome=" & obj.nome_cli
    data_raw = data_raw & "&registro_geral=" & obj.ie_Cli
    data_raw = data_raw & "&sexo=" & obj.sexo_cli
    data_raw = data_raw & "&data_nasc=" & obj.datanasc_cli
    data_raw = data_raw & "&email_cli=" & obj.email_cli
    data_raw = data_raw & "&endereco=" & obj.Logradouro
    data_raw = data_raw & "&complemento=" & obj.complemento
    data_raw = data_raw & "&bairro=" & obj.bairro_cli
    data_raw = data_raw & "&mae=" & obj.mae_cli
    data_raw = data_raw & "&pai=" & obj.pai_cli
    data_raw = data_raw & "&dataal_cli=" & obj.dataalt_cli
    data_raw = DuplicaAD(data_raw)
    
    
    ret = sendPost(url, data_raw, bearer_token)
    
    APIAtualizarCliente = ret
    'P_Status
End Function

Function APIListarClientes(nome As String) As String
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    P_Status "enviando documento"
    url = "https://piki.app.br/api/desktop_clientes/"
    
    
    
    
    

    data_raw = "nome=" & nome
    data_raw = DuplicaAD(data_raw)
    
    
    ret = sendPost(url, data_raw, bearer_token)
    
    APIListarClientes = ret
    P_Status
End Function
Function APIPesquisarCliente(cpf_cli As String) As String
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    P_Status "enviando documento"
    url = "https://piki.app.br/api/desktop_clientes/"
    
    
    
    
    

    data_raw = "cpf=" & F_SoNumero(cpf_cli)
    data_raw = DuplicaAD(data_raw)
    
    
    ret = sendPost(url, data_raw, bearer_token)
    
    APIPesquisarCliente = ret
    P_Status
End Function
Function APIBuscarDoc(cod_emp As String, seq_doc As String, Optional id_doc As String = "") As String
    Dim url As String
    Dim parc As String
    Dim d As doc.doc
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    P_Status "buscando documento"
    url = "https://piki.app.br/api/desktop_doc?seq_doc=" & seq_doc & "&cod_emp=" & cod_emp & "&doc.id=" & id_doc
    
    
    ret = sendGet(url, bearer_token)
    
    APIBuscarDoc = substituir_aspas(ret)
    P_Status
End Function
Function APIBuscarItDoc(id_doc As String, Optional it_doc As String = "") As String
    Dim url As String
    Dim parc As String
    Dim d As doc.doc
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    P_Status "buscando documento"
    url = "https://piki.app.br/api/desktop_doc/lista_parcelas?id_doc=" & id_doc & "&it_doc=" & it_doc
    
    
    ret = sendGet(url, bearer_token)
    
    'APIBuscarItDoc = substituir_aspas(ret)
    APIBuscarItDoc = (ret)
    P_Status
End Function
Function APIListarFormasPagamento() As String
    Dim url As String
    Dim parc As String
    Dim d As doc.doc
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    P_Status "buscando formas de pagamento"
    url = "https://piki.app.br/api/desktop_baixa/lista_forma_pagamento"
    
    
    ret = sendGet(url, bearer_token)
    
    'APIBuscarItDoc = substituir_aspas(ret)
    APIListarFormasPagamento = (ret)
    P_Status
End Function
Function APIListarParcelasBaixa(Optional num_doc As String, _
    Optional id_doc As String, Optional em_aberto As String = "V", _
    Optional it_doc As String, Optional setor As String, _
    Optional nome_cli As String, Optional somente_abertas As String = "F", _
    Optional cpf As String = "", Optional data_ini As String, Optional data_fim As String, _
    Optional historico As String, Optional status_cli As String) As String
    Dim url As String
    Dim parc As String
    Dim d As doc.doc
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    P_Status "buscando parcelas"
    url = "https://piki.app.br/api/desktop_baixa/lista_parcelas?num_doc=" & num_doc & _
        "id_doc=" & id_doc & "emaberto=" & em_aberto & "it_doc=" & it_doc & _
        "setor=" & setor & "nome_cli=" & nome_cli & "somente_abertas=" & somente_abertas & _
        "cpf=" & cpf & "data_ini=" & data_ini & "data_fim=" & data_fim & _
        "historico=" & historico & "status_cli=" & status_cli
        
        
        
    
    
    ret = sendGet(url, bearer_token)
    
    'APIBuscarItDoc = substituir_aspas(ret)
    APIListarParcelasBaixa = (ret)
    P_Status
End Function
Function APIInserirBaixa(valor_baixa As String, historico_descricao As String, fk_forma_pagamento As String, _
    listaParcelas As String) As String
    Dim url As String
    Dim parc As String
    Dim d As doc.doc
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    P_Status "buscando parcelas"
    
    data_raw = "valor_baixa=" & valor_baixa
    data_raw = data_raw & "&historico_descricao=" & historico_descricao
    data_raw = data_raw & "&fk_forma_pagamento=" & fk_forma_pagamento
    data_raw = data_raw & "&listaParcelas=" & listaParcelas
    
    url = "https://piki.app.br/api/desktop_baixa/inserir"
    ret = sendPost(url, data_raw, bearer_token)
        
        
    
    
    ret = sendGet(url, bearer_token)
    
    'APIBuscarItDoc = substituir_aspas(ret)
    APIInserirBaixa = (ret)
    P_Status
End Function

Sub EnviarDocInternet(seq_doc As String, Optional avisar As Boolean = True)
    Dim s As String
    Dim id As String
    
    
    
    s = API.APIInserirDoc(seq_doc)
     
    id = GetValorJson("id", s)
    If s <> "" Then
        If Not doc.altera_id_doc_web(seq_doc, id) Then
            If avisar Then
               MsgBox "erro na vinculação do documento web!" & s & " - seq_doc: " & seq_doc, vbInformation
            Else
                P_Status "erro na vinculação do documento web!" & s & " - seq_doc: " & seq_doc
            End If
            
        End If
        'If avisar Then
        '    MsgBox "Dados enviados", vbInformation
        'End If
    End If
    
End Sub



Public Sub SendWhatApp(numero As String, mensagem As String)
    Dim http As Object
    Dim url As String
    Dim data As String
    Dim responseText As String

    url = "http://localhost:3000/send-message"
    
    ' String JSON com os parâmetros passados
    data = "{""number"": """ & numero & """, ""message"": """ & mensagem & """}"

    ' Cria o objeto HTTP
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Configura e envia a requisição POST
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.sEnd data
    
    ' Obtém a resposta
    responseText = http.responseText
    
    ' Exibe a resposta
    MsgBox "Resposta do servidor: " & responseText
End Sub


Sub sincronizar_docs(Optional inicio As String = "", Optional fim As String = "")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select * from " & TAB_DOC & " where 1>0 "
    If inicio <> "" Then
        strSQL = strSQL & " and data_doc>=" & F_GravaData(inicio)
        
        
    End If
    If fim <> "" Then
        strSQL = strSQL & " and data_doc<=" & F_GravaData(fim)
        
        
    End If
    
    strSQL = strSQL & " order by seq_doc desc "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            
            APIInserirCliente !cod_ter
            
            API.EnviarDocInternet !seq_doc, False
            
            
            If .AbsolutePosition Mod 10 = 0 Then
                mdiFrm.lblMsg.caption = "sincronizando " & .AbsolutePosition & "/" & .RecordCount
                mdiFrm.lblMsg.Refresh
                
            End If
            'DoEvents
            .MoveNext
        Loop
        .Close
        P_Status
    End With
    
    
End Sub

Sub sincronizar_doc(seq_doc As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    API.EnviarDocInternet seq_doc, True
            
     
    
    
    
End Sub
Sub excluir_doc(seq_doc As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    API.EnviarDocInternet seq_doc, True
            
     
    
    
    
End Sub


Sub AtualizaSecoes()
    Dim ret As String
    Dim c As New Collection
    Dim rSecoes As New ADODB.Recordset
    
    ret = APIListarSecao("")
    Set c = ParseJsonIntoList(ret)
    
    
    
    Set rSecoes = JSONUtil.CollectionToRecordset(c)
    With rSecoes
        Do While Not .EOF
            
            .MoveNext
        Loop
    End With
    
    
End Sub
Sub MigrarSecoes()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim ret As String
    strSQL = "select * from " & TAB_SECOES & " order by cod_sec "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If !cod_sec <> "" Then
                ret = API.APIMergeSecoes(!cod_sec, !nome_sec)
                
            End If
            .MoveNext
        Loop
        .Close
    End With
    
End Sub
Sub MigrarSubsec()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim ret As String
    strSQL = "select * from " & TAB_SUBSEC & " order by cod_sub "
    
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If !cod_sec <> "" Then
                ret = API.APIMergeSubSec(!cod_sub, !nome_sub, !cod_sec)
                P_Status "registro " & .AbsolutePosition & "/" & .RecordCount
            End If
            .MoveNext
        Loop
        .Close
    End With
    P_Status ""
    
End Sub

Sub sincronizar_clientes(Optional ano As String = "", Optional cod_cli As String = "")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim ret As String
    strSQL = "select * from " & TAB_CLIENTES & " where 1>0 "
    If ano <> "" Then
        strSQL = strSQL & " and data_cli>='" & ano & "-01-01' and data_cli<='" & ano & "-12-31'"
    End If
    If cod_cli <> "" Then
        strSQL = strSQL & " and cod_cli='" & cod_cli & "' "
    End If
    strSQL = strSQL & " order by cod_cli desc "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            
            'API.APIAtualizarCliente !cod_cli, ""
            ret = API.APIInserirCliente(!cod_cli)
            
            'If .AbsolutePosition Mod 10 = 0 Then
            '    mdiFrm.lblMsg.caption = .AbsolutePosition & " / " & .RecordCount
            '    mdiFrm.lblMsg.Refresh
            'End If
            mdiFrm.lblMsg.caption = "Atualizando " & .AbsolutePosition & " / " & .RecordCount & " Cliente: " & !cod_cli & " - " & ret
            mdiFrm.lblMsg.Refresh
            DoEvents
            .MoveNext
        Loop
        .Close
    End With
    
    
End Sub
Function APIListarSecao(nome As String) As String
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    P_Status "enviando documento"
    url = "https://piki.app.br/api/desktop_secao/"
    data_raw = "nome=" & nome
    data_raw = DuplicaAD(data_raw)
    
    ret = sendPost(url, data_raw, bearer_token)
    
    APIListarSecao = ret
    P_Status
End Function


Function APIListarSubSec(nome As String) As String
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    P_Status "enviando documento"
    url = "https://piki.app.br/api/desktop_sub_secao/"
    data_raw = "nome=" & nome
    data_raw = DuplicaAD(data_raw)
    
    ret = sendPost(url, data_raw, bearer_token)
    
    APIListarSubSec = ret
    P_Status
End Function
Function APIMergeSubSec(cod_sub As String, nome_sub As String, cod_sec As String) As String
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String

    url = "https://piki.app.br/api/desktop_sub_secao/merge"
    
    
    
    
    

    data_raw = "cod_sub=" & cod_sub
    data_raw = data_raw & "&cod_sub=" & cod_sub
    data_raw = data_raw & "&nome_sub=" & nome_sub
    data_raw = data_raw & "&cod_sec=" & cod_sec
    data_raw = DuplicaAD(data_raw)
    
    
    ret = sendPost(url, data_raw, bearer_token)
    
    APIMergeSubSec = ret
    
End Function
Function APIMergeSecoes(cod_sec As String, nome_sec As String) As String
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    P_Status "enviando documento"
    url = "https://piki.app.br/api/desktop_secao/merge"
    
    
    
    
    

    data_raw = "cod_sec=" & cod_sec
    
    data_raw = data_raw & "&nome_sec=" & nome_sec
    
    data_raw = DuplicaAD(data_raw)
    
    
    ret = sendPost(url, data_raw, bearer_token)
    
    APIMergeSecoes = ret
    P_Status
End Function

Function APIPesquisarProduto(Optional cod_pro As String, Optional cod_sub As String, Optional cod_sec As String, Optional nome_pro As String) As String
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    
    url = "https://piki.app.br/api/desktop_produtos/"
    
    
    
    

    data_raw = "id_pro=" & cod_pro
    If cod_sub <> "" Then
        data_raw = data_raw & "&cod_sub=" & cod_sub
    End If
    If cod_sec <> "" Then
        data_raw = data_raw & "&cod_sec=" & cod_sec
    End If
    data_raw = DuplicaAD(data_raw)
    
    
    ret = sendPost(url, data_raw, bearer_token)
    
    APIPesquisarProduto = ret
    
End Function
Sub ApiAtualizarProdutos(Optional listaProd As String = "")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim cod_sub As String
    Dim ret As String
    Dim temSubSec As Boolean
    If strCod_Emp = "01" Then
        MsgBox "Não é possível atualizar produtos sendo da empresa matriz", vbExclamation
        Exit Sub
    End If
    If InStr(1, listaProd, "subsec") > 0 Then
        cod_sub = Right(listaProd, 3)
        temSubSec = True
    End If
    
    
    If listaProd = "" Or temSubSec Then
        strSQL = "select * from " & TAB_SUBSEC & " where cod_sub <> '' "
        If cod_sub <> "" Then
            strSQL = strSQL & " and cod_sub='" & cod_sub & "'"
        End If
        
        
        
        strSQL = strSQL & " order by cod_sub "
        abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            If listaProd = "" Or temSubSec Then
                ret = APIPesquisarProduto("", !cod_sub)
                importarProdutos ret
                P_Status "atualizando produtos da sub-seção " & !nome_sub
            Else
                
                
            End If
            
            
            
            
            .MoveNext
        Loop
        P_Status
        .Close
    End With
    
    Else
        ret = APIPesquisarProduto(listaProd, "")
        importarProdutos ret
    End If
    
    
    
End Sub

Sub APIEnviarSubSec()
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    P_Status "enviando documento"
    If strCod_Emp <> "01" Then
        Exit Sub
    End If
    url = "https://piki.app.br/api/desktop_sub_secao/alterar"
    
    
    strSQL = "select * from " & TAB_SUBSEC
    abreObjetoAdo rPesq, strSQL
    With rPesq
    
        Do While Not .EOF

            data_raw = "cod_sub=" & !cod_sub
            data_raw = data_raw & "&nome_sub=" & !nome_sub
            data_raw = data_raw & "&cod_sec=" & !cod_sec
            data_raw = DuplicaAD(data_raw)
            
            
            ret = sendPost(url, data_raw, bearer_token)
            P_Status "Atualizando subseção " & .AbsolutePosition & " de " & .RecordCount
            .MoveNext
        Loop
    End With
    
End Sub

Sub APIAlterarProduto(cod_pro As String)
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    P_Status "enviando documento"
    If strCod_Emp <> "01" Then
        Exit Sub
    End If
    url = "https://piki.app.br/api/desktop_produtos/alterar"
    
    
    strSQL = "select * from " & TAB_PRODUTOS & " where cod_pro='" & cod_pro & "'"
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        
        If .RecordCount > 0 Then
            P_Status "Atualizando produto " & !nome_pro
            data_raw = "id_pro=" & !cod_pro
            data_raw = data_raw & "&cod_pro=" & !cod_pro
            data_raw = data_raw & "&cod_sub=" & !cod_sub
            data_raw = data_raw & "&nome_pro=" & !nome_pro
            data_raw = data_raw & "&localizacao_pro=" & !endereco_pro
            data_raw = data_raw & "&margem_lucro=" & !lucro_pro
            data_raw = data_raw & "&unidade_pro=" & !unidade_pro
            data_raw = data_raw & "&gtin_pro=" & !ean_pro
            data_raw = data_raw & "&ncm_pro=" & !ncm_pro
            data_raw = data_raw & "&pv_pro=" & !pv_pro
            data_raw = data_raw & "&pq_pro=" & !pq_pro
            data_raw = data_raw & "&cf_pro=" & !cf_pro
            data_raw = data_raw & "&cr_pro=" & !cr_pro
            data_raw = data_raw & "&cod_fabrica=" & !cod_fab
            data_raw = DuplicaAD(data_raw)
            
            
            ret = sendPost(url, data_raw, bearer_token)
            
            
        End If
    End With
    P_Status
End Sub

Function APIMergeProduto(cod_pro As String, Optional mostra_status As Boolean = False) As String
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim i As Integer
    Dim inicio
    Dim fim
    On erro GoTo trata_Erro
    
    If strCod_Emp <> "01" Then
        MsgBox "Só é possível enviar produtos a partir da matriz", vbExclamation
        Exit Function
    End If
    url = "https://piki.app.br/api/desktop_produtos/merge"
    If mostra_status Then
        P_Status "enviando produto " & cod_pro
    End If
    strSQL = "select * from " & TAB_PRODUTOS & " where 1 > 0 "
    If cod_pro <> "" Then
        strSQL = strSQL & " and cod_pro='" & cod_pro & "'"
    End If
    strSQL = strSQL & " order by cod_pro "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        
        If .RecordCount > 0 Then
            Do While Not .EOF
                
                data_raw = "id_pro=" & !cod_pro
                data_raw = data_raw & "&cod_pro=" & !cod_pro
                data_raw = data_raw & "&cod_sub=" & !cod_sub
                data_raw = data_raw & "&nome_pro=" & !nome_pro
                data_raw = data_raw & "&localizacao_pro=" & !endereco_pro
                data_raw = data_raw & "&margem_lucro=" & !lucro_pro
                data_raw = data_raw & "&unidade_pro=" & !unidade_pro
                data_raw = data_raw & "&gtin_pro=" & !ean_pro
                data_raw = data_raw & "&ncm_pro=" & !ncm_pro
                data_raw = data_raw & "&pv_pro=" & !pv_pro
                data_raw = data_raw & "&pq_pro=" & !pq_pro
                data_raw = data_raw & "&cf_pro=" & !cf_pro
                data_raw = data_raw & "&cr_pro=" & !cr_pro
                
                data_raw = DuplicaAD(data_raw)
                
                inicio = Now
                ret = sendPost(url, data_raw, bearer_token)
                If mostra_status Then
                    If Not InStr(1, ret, "sucesso") > 0 Then
                        MsgBox "houve um erro no envio do produto " & !cod_pro & ". Tente novamente!" & ret, vbExclamation
                    End If
                End If
                
                fim = DateDiff("s", inicio, Now)
                i = i + 1
                If mostra_status Then
                
                    P_Status " enviando " & !cod_pro
                
                End If
                .MoveNext
            Loop
        End If
        APIMergeProduto = ret
    End With
'    P_Status
trata_Erro:
    
End Function

Sub APIEnviarProdutos(Optional tamanho_lote As Integer = 50, Optional cod_sub As String = "", Optional cod_pro = "")
    Dim url As String
    Dim parc As String
    Dim obj As Clientes.cliente
    Dim ret As String
    Dim data_raw As String
    Dim parcelas As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim limite As Integer
    Dim pagina As Integer
    Dim offset As Integer
    Dim LogLote As String
    Dim i As Integer
    Dim inicio
    Dim fim
    
    
    url = "https://piki.app.br/api/desktop_produtos/merge_lote"
    
    If strCod_Emp <> "01" Then
        MsgBox "Só é possível enviar produtos a partir da matriz", vbExclamation
        P_Status
        Exit Sub
    End If
    
    P_Status "enviando documento"
    pagina = 0
    limite = tamanho_lote
    
    
    Do
        offset = pagina * limite
        strSQL = "select * from " & TAB_PRODUTOS & " where 1 > 0 "
        If cod_sub <> "" Then
            strSQL = strSQL & " and cod_sub='" & cod_sub & "'"
        End If
        If cod_pro <> "" Then
            strSQL = strSQL & " and cod_pro in(" & cod_pro & ") "
        End If
        strSQL = strSQL & " order by cod_sub,cod_pro desc "
        
        strSQL = strSQL & " LIMIT " & limite & " OFFSET " & offset
        
        abreObjetoAdo rPesq, strSQL
        
        If rPesq.EOF Then Exit Do
        'data_raw = "{" & entreAspas("lista_produtos") & ":["
        data_raw = "["
        inicio = Now
        With rPesq
            Do While Not .EOF
                
                
                data_raw = data_raw & "{" & entreAspas("id_pro") & ":" & entreAspas(!cod_pro) & ","
                data_raw = data_raw & entreAspas("cod_pro") & ":" & entreAspas(!cod_pro) & ","
                data_raw = data_raw & entreAspas("cod_sub") & ":" & entreAspas(!cod_sub) & ","
                data_raw = data_raw & entreAspas("nome_pro") & ":" & entreAspas(!nome_pro) & ","
                'data_raw = data_raw & entreAspas("localizacao_pro") & ":" & entreAspas(!endereco_pro) & ","
                data_raw = data_raw & entreAspas("margem_lucro") & ":" & entreAspas(!lucro_pro) & ","
                data_raw = data_raw & entreAspas("unidade_pro") & ":" & entreAspas(!unidade_pro) & ","
                data_raw = data_raw & entreAspas("gtin_pro") & ":" & entreAspas(!ean_pro) & ","
                data_raw = data_raw & entreAspas("ncm_pro") & ":" & entreAspas(!ncm_pro) & ","
                data_raw = data_raw & entreAspas("pv_pro") & ":" & entreAspas(!pv_pro) & ","
                data_raw = data_raw & entreAspas("pq_pro") & ":" & entreAspas(!pq_pro) & ","
                data_raw = data_raw & entreAspas("cf_pro") & ":" & entreAspas(!cf_pro) & ","
                data_raw = data_raw & entreAspas("cr_pro") & ":" & entreAspas(!cr_pro) & "}"
                
                .MoveNext
                If Not .EOF Then
                    data_raw = data_raw & ","
                End If
            Loop
            data_raw = data_raw & "]"
            data_raw = DuplicaAD(data_raw)
            data_raw = "lista_produtos=" & Base64EncodeUTF8(data_raw)
            Dim arquivo As String
            If Dir(App.path & "\envio", vbDirectory) = "" Then
                MkDir (App.path & "\envio")
            End If
            arquivo = App.path & "\envio\" & pagina & i & ".txt"
           ' GravaArquivoExternoUTF8 arquivo, data_raw
            
            'ret = sendPost(url, data_raw, bearer_token)
            'ret = sendPostFile(url, arquivo, bearer_token)
            ret = sendPost(url, data_raw, bearer_token)
            
            
            If InStr(1, ret, "sucesso") = 0 Then
                LogLote = LogLote + ret
            End If
            fim = DateDiff("s", inicio, Now)
            i = i + 1
            P_Status " enviando " & pagina * limite
            
            .Close
        
        End With
    pagina = pagina + 1
    GravaArquivoExternoUTF8 arquivo, LogLote
Loop
    
    P_Status
End Sub


Sub enviarTodos()

End Sub

Sub insere_fila(codigo As String, tipo As tipo_fila)
    '1 cliente, 2 documento, 3 - produtos
    On Error GoTo trata_Erro:
    Dim strSQL As String
    strSQL = "replace into fila_transferencia(codigo, tipo)values('" & codigo & "',"
    strSQL = strSQL & tipo & ")"
    Conn.Execute strSQL
    Exit Sub
trata_Erro:
    
    MsgBox "Não foi possível inserir na fila de transferência!" & Err.Description, vbExclamation
End Sub


Sub insere_fila_produtos(codigo As String)
    '1 cliente, 2 documento, 3 - produtos
    On Error GoTo trata_Erro:
    insere_fila codigo, 3
    Exit Sub
trata_Erro:
    
    MsgBox "Não foi possível inserir na fila de transferência!" & Err.Description, vbExclamation
End Sub


Sub executa_fila_produtos()
 '1 cliente, 2 documento, 3 - produtos
    On Error GoTo trata_Erro:
    Dim strSQL As String
    Dim ado As New ADODB.Recordset
    Dim sIn As String
    Dim sId As String
    strSQL = "select * from fila_transferencia where tipo = 3 "
    strSQL = strSQL & " order by id limit 100"
    P_Status "executando fila..."
    abreObjetoAdo ado, strSQL
    
    With ado
        Do While Not .EOF
            sIn = sIn & "'" & !codigo & "'"
            sId = sId & !id
            
            
            
            .MoveNext
            If Not .EOF Then
                sIn = sIn & ", "
                sId = sId & ", "
            End If
        Loop
        APIEnviarProdutos 100, "", sIn
        
        strSQL = "delete from fila_transferencia where id in(" & sId & ") "
        Conn.Execute strSQL
    End With
    P_Status
    Exit Sub
trata_Erro:
    P_Status ""
    MsgBox "Não foi possível inserir na fila de transferência!" & Err.Description, vbExclamation
End Sub




Function executa_fila(tipo As String) As String
 '1 cliente, 2 documento, 3 - produtos
    On Error GoTo trata_Erro:
    Dim strSQL As String
    Dim ado As New ADODB.Recordset
    strSQL = "select * from fila_transferencia where tipo = " & tipo
    strSQL = strSQL & " order by id "
    Dim ret As String
    abreObjetoAdo ado, strSQL
    With ado
        Do While Not .EOF
            If tipo = 1 Then
                ret = APIInserirCliente(!codigo)
            ElseIf tipo = 2 Then
                ret = APIInserirDoc(!codigo, False, False)
            ElseIf tipo = 3 Then
                ret = APIMergeProduto(!codigo, False)
            End If
            'DoEvents
            If InStr(1, ret, "sucesso") Then
                strSQL = "delete from fila_transferencia where id=" & !id
                Conn.Execute strSQL
            End If
             P_Status "enviando " & .AbsolutePosition & "/" & .RecordCount
            
            .MoveNext
        Loop
    End With
    executa_fila = ret
    P_Status ""
    Exit Function
trata_Erro:
    
    MsgBox "Não foi possível inserir na fila de transferência!" & Err.Description, vbExclamation
End Function


Sub gerar_fila_doc(data_ini As String, data_fim As String, _
    Optional reman As Boolean = False, Optional atualizar_cliente As Boolean = True)
    Dim strSQL As String
    
    strSQL = "INSERT INTO fila_transferencia (codigo, tipo)"
    strSQL = strSQL & " SELECT DISTINCT td.seq_doc, 2 "
    strSQL = strSQL & "FROM tab_doc td "
    strSQL = strSQL & "INNER JOIN tab_itdoc ti ON td.seq_doc = ti.seq_doc "
    strSQL = strSQL & " WHERE td.data_doc >= " & F_GravaData(data_ini)
    strSQL = strSQL & " AND td.data_doc <= " & F_GravaData(data_fim)
    If reman Then
        strSQL = strSQL & " and ti.reman_itd>0 "
    End If
    Conn.Execute strSQL
End Sub
Sub gerar_fila_prod()
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "INSERT INTO fila_transferencia (codigo, tipo) "
    strSQL = strSQL & " SELECT cod_pro, 3 "
    strSQL = strSQL & " FROM tab_produtos order by cod_pro desc "
    
    
    Conn.Execute strSQL
    Exit Sub
trata_Erro:
MsgBox Err.Description, vbExclamation
    executa_fila "3"
    
End Sub


