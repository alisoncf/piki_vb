Attribute VB_Name = "Cidade"
Type Cidade
    cod_cid As String
    nome_cid As String
    uf_cid As String
    estado_cid As String
End Type


Sub P_GravaCidade(cod_cid As String, nome_cid As String, estado_cid As String, uf_cid As String)
    Dim strSQL As String
    On Error GoTo trata_erro:
    
    strSQL = "insert into tab_cidades(cod_cid,nome_cid,estado_cid,uf_cid)values("
    strSQL = strSQL & "'" & cod_cid & "',"
    strSQL = strSQL & "'" & F_AjustaTexto(nome_cid) & "',"
    strSQL = strSQL & "'" & estado_cid & "',"
    strSQL = strSQL & "'" & uf_cid & "')"
    conn.Execute strSQL
    
trata_erro:
    
    
End Sub
Function F_BuscaCidade(nome As String, Optional uf As String = "") As Cidade.Cidade
    On Error Resume Next
    frmEdiCid.tag = "x"
    frmEdiCid.txtbxNome.Text = nome
    frmEdiCid.Show 1
    If frmEdiCid.tag <> "" Then
        With frmEdiCid.adc.Recordset
            F_BuscaCidade.cod_cid = !cod_cid
            F_BuscaCidade.estado_cid = !estado_cid
            F_BuscaCidade.nome_cid = !nome_cid
            F_BuscaCidade.uf_cid = !uf_cid
        End With
    End If
    Unload frmEdiCid
End Function
Function SelectCidade(nome As String, Optional uf As String = "") As String
    Dim strSQL As String
    strSQL = "select * from tab_cidades where nome_cid like '%" & nome & "%' "
    strSQL = strSQL & " and uf_cid like '%" & uf & "%' "
    strSQL = strSQL & " order by nome_cid"
    SelectCidade = strSQL

End Function
Sub P_AlteraCidade(cod_cid As String, nome_cid As String, estado_cid As String, uf_cid As String)
    Dim strSQL As String
    On Error GoTo trata_erro:
    
    strSQL = "update tab_cidades set "
    
    strSQL = strSQL & "nome_cid='" & F_AjustaTexto(nome_cid) & "',"
    strSQL = strSQL & "estado_cid='" & estado_cid & "',"
    strSQL = strSQL & "uf_cid='" & uf_cid & "' "
    strSQL = strSQL & " where cod_cid='" & cod_cid & "' "
    conn.Execute strSQL
    
trata_erro:
    
    
End Sub
Function BuscaCidadePorNome(nome As String, uf As String) As String
    BuscaCidadePorNome = F_BuscaUmCampo("cod_cid", "tab_cidades", "nome_cid='" & nome & "' and uf_cid='" & uf & "'")
    
End Function
