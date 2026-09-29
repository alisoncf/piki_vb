Attribute VB_Name = "Servico"
Type tipo_servico
    id As String
    nome As String
End Type
Type servico
    id As String
    tipo As tipo_servico
    nome As String
    valor As Currency
    comissao As Currency
    
End Type

Function GravaTipo_Servico(nome As String, comissao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tipo_servico (nome,comissao)values("
    
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & F_GravaDec(comissao) & ")"
    
    conn.Execute strSQL
    
    GravaTipo_Servico = True
    Exit Function
trata_Erro:
    
    
    
End Function
Function AlteraTipo_Servico(id As String, nome As String, comissao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tipo_servico set nome="
    
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & " comissao=" & F_GravaDec(comissao)
    strSQL = strSQL & " where id=" & id
    
    conn.Execute strSQL
    
    AlteraTipo_Servico = True
    Exit Function
trata_Erro:
    
    
    
End Function



Function GravaServico(id_tipo As String, nome As String, valor As String, comissao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into servico (id_tipo,nome,valor,comissao)values("
    strsq = strSQL & id_tipo & ","
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & F_GravaDec(comissao) & ")"
    conn.Execute strSQL
    
    GravaServico = True
    Exit Function
trata_Erro:
    
    
    
End Function

Function AlteraServico(id As String, id_tipo As String, nome As String, valor As String, comissao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update servico set  "
    strsq = strSQL & "id_tipo=" & id_tipo & ","
    strSQL = strSQL & "nome='" & nome & "',"
    strSQL = strSQL & "valor=" & F_GravaDec(valor) & ","
    strSQL = strSQL & "comissao=" & F_GravaDec(comissao) & " "
    strSQL = strSQL & " where id=" & id
    conn.Execute strSQL
    
    AlteraServico = True
    Exit Function
trata_Erro:
    
    
    
End Function
Function BuscaTipoServico(busca As String) As tipo_servico
'    On Error Resume Next
    
    frmEdiTipoServ.txtPesquisar.Text = busca
    
    frmEdiTipoServ.Show 1
    
    If frmEdiTipoServ.tag <> "" Then
        BuscaTipoServico.id = frmEdiTipoServ.ado!id
        BuscaTipoServico.nome = frmEdiTipoServ.ado!nome
    End If
    Unload frmEdiTipoServ
    
End Function
Function BuscaServico(busca As String) As servico.servico
    'On Error Resume Next
    frmEdiServico.txtPesquisar.Text = busca
    frmEdiServico.Show 1
    If frmEdiServico.tag <> "" Then
        BuscaServico.id = frmEdiServico.ado!id
        BuscaServico.nome = frmEdiServico.ado!nome
        BuscaServico.tipo.id = frmEdiServico.ado!id_tipo
        BuscaServico.tipo.nome = frmEdiServico.ado!nome_servico
    End If
    Unload frmEdiServico
End Function

