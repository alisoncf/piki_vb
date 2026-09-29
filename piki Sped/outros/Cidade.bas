Attribute VB_Name = "Cidade"
Type cidade
    codigo As String
    nome As String
    id As String
    uf As String
End Type


Sub P_GravaCidade(cod_cid As String, nome_cid As String, id_cid As String, uf_cid As String)
    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    strSQL = "insert into tab_cidades(cod_cid,nome_cid,id_cid,uf_cid)values("
    strSQL = strSQL & "'" & cod_cid & "',"
    strSQL = strSQL & "'" & F_AjustaTexto(nome_cid) & "',"
    strSQL = strSQL & "'" & id_cid & "',"
    strSQL = strSQL & "'" & uf_cid & "')"
    banco.Execute strSQL
    
trata_Erro:
    
    
End Sub
