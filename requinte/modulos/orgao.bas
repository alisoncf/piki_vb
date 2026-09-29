Attribute VB_Name = "orgao"
Type orgao
    cod_org As String
    nome_org As String
    cod_gru As String
    nome_gru As String
    cod_set As String
    nome_set As String
    nome_cid  As String
    endereco_org As String
    fone_org As String
    cep_org As String
    UF As String
End Type

Function F_BuscaOrgao(expr As String, Optional sCod As String = "") As orgao.orgao
    Dim fr As New frmEdiOrgao
    
    fr.txtbxBusca.Text = expr
    fr.tag = "x"
    If sCod <> "" Then
        fr.txtbxBusca.tag = sCod
    End If
    
    fr.Show 1
    If fr.tag <> "" Then
        With fr.ado
            If fr.ado.RecordCount > 0 Then
                F_BuscaOrgao.cod_gru = !cod_gru
                F_BuscaOrgao.cod_org = !cod_org
                F_BuscaOrgao.cod_set = !cod_set
                F_BuscaOrgao.nome_cid = !nome_cid
                F_BuscaOrgao.nome_gru = !nome_gru
                F_BuscaOrgao.nome_org = !nome_org
                F_BuscaOrgao.UF = !uf_cid
                F_BuscaOrgao.endereco_org = !endereco_org
                F_BuscaOrgao.cep_org = !cep_org
                F_BuscaOrgao.nome_set = !nome_set
            End If
        End With
    End If
    Unload fr
End Function
Function F_BuscaOrgaoPorCod(Optional sCod As String = "") As orgao.orgao
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    
    strSQL = orgao.F_SelectOrgao(sCod)
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        F_BuscaOrgaoPorCod.cod_gru = !cod_gru
        F_BuscaOrgaoPorCod.cod_org = !cod_org
        F_BuscaOrgaoPorCod.cod_set = !cod_set
        F_BuscaOrgaoPorCod.nome_cid = !nome_cid
        F_BuscaOrgaoPorCod.nome_gru = !nome_gru
        F_BuscaOrgaoPorCod.nome_org = !nome_org
        F_BuscaOrgaoPorCod.UF = !uf_cid
        F_BuscaOrgaoPorCod.endereco_org = !endereco_org
        F_BuscaOrgaoPorCod.cep_org = !cep_org
        F_BuscaOrgaoPorCod.nome_set = !nome_set
        
        .Close
    End With
trata_Erro:
    
End Function

Function F_GravaOrgao(cod_gru As String, nome_org As String, _
    cod_set As String, endereco_org As String, fone_org As String, _
    cep_org As String, data_org As String, obs As String) As String
    
    Dim sCod As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    sCod = MaxKeySQL("tab_orgaos", "cod_org", "", "0000#")
    strSQL = "insert into tab_orgaos(cod_org,cod_gru,nome_org,cod_set,endereco_org,fone_org,cep_org,data_org,obs)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & cod_gru & "',"
    strSQL = strSQL & "'" & nome_org & "',"
    strSQL = strSQL & "'" & cod_set & "',"
    strSQL = strSQL & "'" & endereco_org & "',"
    strSQL = strSQL & "'" & fone_org & "',"
    strSQL = strSQL & "'" & cep_org & "',"
    strSQL = strSQL & F_GravaData(data_doc) & ","
    strSQL = strSQL & "'" & obs & "')"
    
    Conn.Execute strSQL
    F_GravaOrgao = strSQL
    Exit Function
trata_Erro:
    F_GravaOrgao = ""
End Function
Function F_SelectOrgao(Optional cod_org As String = "", Optional nome_org As String = "") As String
    strSQL = "select tog.*, tg.nome_gru,ts.nome_set,tc.nome_cid,tc.uf_cid,ts.cod_set,tb.nome_cob "
    strSQL = strSQL & " from  tab_orgaos as tog  left join tab_setores as ts on ts.cod_set=tog.cod_set "
    strSQL = strSQL & " left join tab_contacob as tb on tb.cod_cob=ts.cobrador_set  "
    strSQL = strSQL & " left join tab_cidades as tc on tc.cod_cid=ts.cod_cid  "
    strSQL = strSQL & " left join tab_grupocli as tg on tg.cod_gru=tog.cod_gru  "
    strSQL = strSQL & " where 1>0"
    
    If nome_org <> "" Then
        strSQL = strSQL & " and nome_org like '%" & nome_org & "%'"
    End If
    If cod_org <> "" Then
        strSQL = strSQL & " and cod_org ='" & cod_org & "'"
    End If
    F_SelectOrgao = strSQL
End Function
Sub P_AlteraOrgao(cod_org As String, cod_gru As String, nome_org As String, cod_set As String, _
    endereco_org As String, cep_org As String, fone_org As String, cobrador_set As String, obs As String, data_org As String)
    Dim sCod As String
    Dim strSQL As String
    'on error goto trata_Erro
    
    strSQL = "update  tab_orgaos  set "
    
    strSQL = strSQL & "cod_gru='" & cod_gru & "',"
    strSQL = strSQL & "nome_org='" & nome_org & "',"
    strSQL = strSQL & "endereco_org='" & endereco_org & "',"
    strSQL = strSQL & "cep_org='" & cep_org & "',"
    strSQL = strSQL & "Fone_org='" & fone_org & "',"
    strSQL = strSQL & "cod_set='" & cod_set & "', "
    strSQL = strSQL & "data_org=" & F_GravaData(data_org) & ", "
    strSQL = strSQL & "obs='" & obs & "', "
    strSQL = strSQL & "cobrador_org='" & cobrador_set & "' "
    strSQL = strSQL & " where cod_org='" & cod_org & "' "
    Conn.Execute strSQL
    


End Sub

Function F_SelectOrgaoCli(cod_cli As String, Optional cod_org As String = "", Optional nome_org As String = "") As String
    strSQL = "select tog.*, tg.nome_gru,ts.nome_set,tc.nome_cid,tc.uf_cid,tb.cod_cob,tb.nome_cob "
    strSQL = strSQL & " from tab_contacob as tb, tab_setores as ts, tab_cidades as tc, "
    strSQL = strSQL & " tab_grupocli as tg, tab_orgaos as tog, tab_clientes as tcli "
    strSQL = strSQL & " where tg.cod_gru=tog.cod_gru "
    strSQL = strSQL & " and ts.cod_set=tog.cod_set "
    strSQL = strSQL & " and tc.cod_cid=ts.cod_cid "
    strSQL = strSQL & " and tb.cod_cob=tog.cobrador_org "
    strSQL = strSQL & " and tcli.cod_org=tog.cod_org "
    If cod_cli <> "" Then
        strSQL = strSQL & " and cod_cli='" & cod_cli & "' "
    End If
    If nome_org <> "" Then
        strSQL = strSQL & " and nome_org like '%" & nome_org & "%'"
    End If
    If cod_org <> "" Then
        strSQL = strSQL & " and cod_org ='" & cod_org & "'"
    End If
    F_SelectOrgaoCli = strSQL
End Function

