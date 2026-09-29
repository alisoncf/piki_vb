Attribute VB_Name = "Servicos"
Type Tipo_Servico
    codigo As String
    nome As String
End Type
Function buscaTipoServico(Optional cod As String = "", Optional nome As String = "") As Tipo_Servico
    Dim strSQL As String
    strSQL = "select * from tipo_servi"
End Function

Function SelectServico(nome As String, Optional tipo As String = "") As String

    Dim strSQL As String
    strSQL = "select s.*,t.nome from servico right join tipo_servico as t "
    strSQL = strSQL & " on s.tipo_servico=t.id "
    strSQL = strSQL & " where s.nome like '%" & nome & "%' "
    If tipo <> "" Then
        strSQL = strSQL & " and tipo=" & tipo
    End If
    strSQL = strSQL & " order by nome "
    SelectServico = strSQL
End Function



Function SelectTipoServico(nome As String) As String

    Dim strSQL As String
    strSQL = "select * from tipo_servico  "
    
    strSQL = strSQL & " where nome like '%" & nome & "%' "
    
    strSQL = strSQL & " order by nome "
    SelectTipoServico = strSQL
End Function


Public Function GravaServico(nome As String, tipo As String, valor As String, comissao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_erro
    
    strSQL = "insert into servico(nome,tipo,valor,comissao)values("
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & tipo & ","
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & F_GravaDec(comissao) & ")"
    Conn.Execute strSQL, Conn
    
    GravaServico = True
    Exit Function
trata_erro:
    GravaServico = False
    
End Function

Public Function AlteraServico(id As String, nome As String, tipo As String, valor As String, comissao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_erro
    
    strSQL = "update servico set "
    strSQL = strSQL & "nome='" & nome & "',"
    strSQL = strSQL & "tipo=" & tipo & ","
    strSQL = strSQL & "valor=" & F_GravaDec(valor) & ","
    strSQL = strSQL & "comissao=" & F_GravaDec(comissao) & " "
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL, Conn
    
    AlteraServico = True
    Exit Function
trata_erro:
    AlteraServico = False
    
End Function

