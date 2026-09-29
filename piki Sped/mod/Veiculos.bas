Attribute VB_Name = "Veiculos"
Type Veiculo_Fabricante
    id As Integer
    descricao As String
End Type
Type Veiculo
    id As Integer
    fabricante As Veiculo_Fabricante
    descricao As String
End Type


Function GravaVeiculo_Fabricante(descricao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into veiculo_fabricante(descricao)values("
    strSQL = strSQL & "'" & descricao & "')"
    conn.Execute strSQL
        
        
    GravaVeiculo_Fabricante = True
    
trata_Erro:
    
    
End Function
Function AlteraVeiculo_Fabricante(id As String, descricao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "update veiculo_fabricante set  "
    strSQL = strSQL & "descricao='" & descricao & "' "
    strSQL = strSQL & " where id='" & id & "' "
    conn.Execute strSQL
        
        
    AlteraVeiculo_Fabricante = True
    
trata_Erro:
    
    
End Function

Function GravaVeiculo(id_fab As String, descricao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "insert into veiculo(id_fab,descricao)values("
    strSQL = strSQL & "'" & id_fab & "',"
    strSQL = strSQL & "'" & descricao & "')"
    conn.Execute strSQL
        
        
    GravaVeiculo = True
    
trata_Erro:
    
    
End Function
Function AlteraVeiculo(id As String, id_fab As String, descricao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "update veiculo set "
    strSQL = strSQL & "id_fab='" & id_fab & "',"
    strSQL = strSQL & "descricao='" & descricao & "' "
    strSQL = strSQL & " where id='" & id & "'"
    conn.Execute strSQL
        
        
    AlteraVeiculo = True
    
trata_Erro:
    
    
End Function

Function SelectVeiculos(Optional id As String, Optional id_fab As String, Optional ordem As String = "descricao") As String
    Dim strSQL As String
    strSQL = "select vf.descricao as fabricante, v.* from veiculo as v, veiculo_fabricante as vf "
    strSQL = strSQL & " on v.id_fab=vf.id "
    strSQL = strSQL & " where 1> 0 "
    If id = "" Then
        strSQL = strSQL & " and id='" & id & "' "
    End If
    If id = "" Then
        strSQL = strSQL & " and id_fab='" & id_fab & "' "
    End If
    
    If ordem <> "" Then
        strSQL = strSQL & " order by " & ordem
    End If
End Function
Function BuscaVeiculo_Fabricante(descricao As String) As Veiculo_Fabricante
    frmEdiVeiculoFabricante.tag = "x"
    frmEdiVeiculoFabricante.txtPesquisar.Text = descricao
    
    frmEdiVeiculoFabricante.Show 1
    If frmEdiVeiculoFabricante.tag <> "" Then
        BuscaVeiculo_Fabricante.descricao = frmEdiVeiculoFabricante.ado!descricao
        BuscaVeiculo_Fabricante.id = frmEdiVeiculoFabricante.ado!id
    End If
    Unload frmEdiVeiculoFabricante
End Function
Function BuscaVeiculo(descricao As String, fabricante As String) As Veiculo
    frmEdiVeiculo.tag = "x"
    frmEdiVeiculo.txtPesquisar.Text = descricao
    frmEdiVeiculo.txtFabricante.Text = fabricante
    
    frmEdiVeiculo.Show 1
    
    If frmEdiVeiculo.tag <> "" Then
        BuscaVeiculo.descricao = frmEdiVeiculo.ado!descricao
        BuscaVeiculo.fabricante.descricao = frmEdiVeiculo.ado!fabricante
        BuscaVeiculo.fabricante.id = frmEdiVeiculo.ado!id_fab
        BuscaVeiculo.id = frmEdiVeiculo.ado!id
    End If
    Unload frmEdiVeiculo
End Function
Function GravaProduto_Veiculo(cod_pro As String, id As String, aplicacao As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "replace into produto_veiculo(cod_pro,id_vei,aplicacao)values("
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & "'" & id & "',"
    strSQL = strSQL & "'" & aplicacao & "')"
    conn.Execute strSQL
        
        
    GravaProduto_Veiculo = True
    
trata_Erro:
    
End Function
Function DeletaProduto_Veiculo(cod_pro As String, id As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    strSQL = "delete from produto_veiculo where "
    strSQL = strSQL & "cod_pro='" & cod_pro & "' and "
    strSQL = strSQL & "id_vei='" & id & "' "
    
    conn.Execute strSQL
        
        
    DeletaProduto_Veiculo = True
    
trata_Erro:
    
End Function

Function SelectProduto_Veiculo(cod_pro As String, Optional descricaoVeiculo As String = "") As String
    Dim strSQL As String
    strSQL = "select v.descricao as veiculo,f.descricao as fabricante,p.*,"
    strSQL = strSQL & " if(v.descricao='" & descricaoVeiculo & "',0,1) as vIg  "
    strSQL = strSQL & " from produto_veiculo  "
    
    strSQL = strSQL & " as p inner join veiculo as v on v.id=p.id_Vei "
    strSQL = strSQL & " inner join veiculo_fabricante as f on v.id_fab=f.id "
    strSQL = strSQL & "  where cod_pro='" & cod_pro & "' "
    strSQL = strSQL & " order by  vig, v.descricao,f.descricao"
    SelectProduto_Veiculo = strSQL
    
End Function
