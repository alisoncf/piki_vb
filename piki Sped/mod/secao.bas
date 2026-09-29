Attribute VB_Name = "secao"
Public Type secao
    cod_sec As String
    nome_sec As String
End Type

Public Type subsec
    cod_sub As String
    nome_sub As String
    cod_sec As String
    descmax As Currency
    comiss As Currency
End Type


Private tabela As String
Private cod As String
Private nome As String
Private cod_sec As String







Function SelecionaSecao(descricao As String) As secao.secao
    Dim p As Padrao1
    p = Padrao.ListaPadrao1(descricao, "tab_secoes", "cod_sec", "nome_sec")
    SelecionaSecao.cod_sec = p.Codigo
    SelecionaSecao.nome_sec = p.nome
End Function
Function GravaSecao(nome As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_erro
    
    tabela = "tab_secoes"
    cod = "cod_sec"
    nome = "nome_sec"
    
    
    
    strSQL = "Insert into " & tabela & "(" & nome & ") values("
    strSQL = strSQL & "'" & nome & "') "
    conn.Execute strSQL
    GravaSecao = True
    
    Exit Function
trata_erro:
End Function

Function AlteraSecao(id As String, nome As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_erro
    tabela = "tab_secoes"
    cod = "cod_sec"
    nome = "nome_sec"
    
    strSQL = "update " & tabela & " set "
    strSQL = strSQL & nome & "='" & nome & "' "
    strSQL = strSQL & " where " & cod & "='" & id & "'"
    conn.Execute strSQL
    AlteraSecao = True
    
    Exit Function
trata_erro:
End Function

Function GravaSubSecao(nome As String, cod_sec As String, descmax As String, comiss As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_erro
    
    strSQL = "Insert into " & TAB_SUBSEC & "(cod_sec,nome_sub,descmax_sub,comiss_sub)values("
    strSQL = strSQL & "'" & cod_sec & "',"
    strSQL = strSQL & "'" & nome & "', "
    strSQL = strSQL & F_GravaDec(descmax) & ","
    strSQL = strSQL & F_GravaDec(comiss) & ")"
    
    
    conn.Execute strSQL
    GravaSubSecao = True
    
    Exit Function
trata_erro:
End Function

Function AlteraSubSecao(id As String, id_secao As String, nome As String, descmax As String, comiss As String) As Boolean
    Dim strSQL As String
    
    On Error GoTo trata_erro
    
    tabela = "tab_subsec"
    cod_sec = "cod_sec "
    
    cod = "cod_sub"
    'nome = "nome_sub"
    
    
    strSQL = "update " & tabela & "  set "
    If descmax <> "" Then
        strSQL = strSQL & "descmax_sub=" & F_GravaDec(descmax) & ","
    End If
    If comiss <> "" Then
        strSQL = strSQL & "comiss_sub=" & F_GravaDec(descmax) & ","
    End If
    strSQL = strSQL & cod_sec & "=" & id_secao & ","
    strSQL = strSQL & "nome_sub='" & nome & "' "
    strSQL = strSQL & " where " & cod & "='" & id & "' "
    conn.Execute strSQL
    AlteraSubSecao = True
    
    Exit Function
trata_erro:
End Function

