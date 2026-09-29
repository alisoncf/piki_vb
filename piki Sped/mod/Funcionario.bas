Attribute VB_Name = "Funcionario"
Type funcionario
    cod_fun As String
    nome_fun As String
    
    guerra_fun As String
    endereco_fun As String
    bairro_fun As String
    cidade_fun As String
    uf_fun As String
    cep_fun As String
    nasc_fun As String
    natural_fun As String
    rg_fun As String
    cpf_fun As String
    sexo_fun As String
    conjunge_fun As String
    pai_fun As String
    mae_fun As String
    endpais_fun As String
    datacad_fun As String
    obs_fun As String
    cod_old As String
    cod_Car As String
    nome_car As String
    venda_car As String
    

End Type
Type usuario
    cod_usu As String
    nome_usu As String
End Type
Function AlteraCorFundoUsu(cod_usu As String, cor)
    Dim strSQL As String
    On Error GoTo trata_Erro:
    strSQL = "update tab_usuarios set corfundo_usu='" & cor & "' where cod_fun = '" & cod_usu & "'"
    conn.Execute strSQL
    Exit Function
trata_Erro:
    P_MostraErro
End Function
Function buscaCorFundoUso(cod_usu) As String
    buscaCorFundoUso = F_BuscaUmCampo("corfundo_usu", "tab_usuarios", "cod_fun='" & cod_usu & "'")
    
End Function
Function F_GravaFun(nome_fun As String, guerra_fun As String, endereco_fun As String, bairro_fun As String, _
    cidade_fun As String, uf_fun As String, cep_fun As String, nasc_fun As String, natural_fun As String, _
    rg_fun As String, cpf_fun As String, sexo_fun As String, conjuge_fun As String, pai_fun As String, _
    mae_fun As String, endpais_fun As String, obs_fun As String, cod_Car As String, telefone_fun As String) As String
    
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    sCod = MaxKeySQL(TAB_FUNCIONARIOS, "cod_fun", "", "00#")
    strSQL = "insert into tab_funcionarios(cod_fun,nome_fun,guerra_fun,endereco_fun,"
    strSQL = strSQL & "bairro_fun,cidade_fun,uf_fun,cep_fun,nasc_fun,natural_fun,rg_fun,cpf_fun,sexo_fun,"
    strSQL = strSQL & "conjuge_fun,pai_fun,mae_fun,endpais_fun,obs_fun,cod_car,telefone_fun)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_fun & "',"
    strSQL = strSQL & "'" & guerra_fun & "',"
    strSQL = strSQL & "'" & endereco_fun & "',"
    strSQL = strSQL & "'" & bairro_fun & "',"
    strSQL = strSQL & "'" & cidade_fun & "',"
    strSQL = strSQL & "'" & uf_fun & "',"
    strSQL = strSQL & "'" & cep_fun & "',"
    strSQL = strSQL & F_GravaData(nasc_fun) & ","
    strSQL = strSQL & "'" & natural_fun & "',"
    strSQL = strSQL & "'" & rg_fun & "',"
    strSQL = strSQL & "'" & cpf_fun & "',"
    strSQL = strSQL & "'" & sexo_fun & "',"
    strSQL = strSQL & "'" & conjuge_fun & "',"
    strSQL = strSQL & "'" & pai_fun & "',"
    strSQL = strSQL & "'" & mae_fun & "',"
    strSQL = strSQL & "'" & endpais_fun & "',"
    strSQL = strSQL & "'" & obs_fun & "',"
    strSQL = strSQL & "'" & cod_Car & "', "
    strSQL = strSQL & "'" & telefone_fun & "') "
    conn.Execute strSQL
    F_GravaFun = sCod
    Exit Function
trata_Erro:
    MostraErro
End Function
Function F_MigraFun(cod_fun As String, nome_fun As String, guerra_fun As String, endereco_fun As String, bairro_fun As String, _
    cidade_fun As String, uf_fun As String, cep_fun As String, nasc_fun As String, natural_fun As String, _
    rg_fun As String, cpf_fun As String, sexo_fun As String, conjuge_fun As String, pai_fun As String, _
    mae_fun As String, endpais_fun As String, obs_fun As String, cod_Car As String) As String
    
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    sCod = cod_fun
    strSQL = "insert into tab_funcionarios(cod_fun,nome_fun,guerra_fun,endereco_fun,"
    strSQL = strSQL & "bairro_fun,cidade_fun,uf_fun,cep_fun,nasc_fun,natural_fun,rg_fun,cpf_fun,sexo_fun,"
    strSQL = strSQL & "conjuge_fun,pai_fun,mae_fun,endpais_fun,obs_fun,cod_car)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_fun & "',"
    strSQL = strSQL & "'" & guerra_fun & "',"
    strSQL = strSQL & "'" & endereco_fun & "',"
    strSQL = strSQL & "'" & bairro_fun & "',"
    strSQL = strSQL & "'" & cidade_fun & "',"
    strSQL = strSQL & "'" & uf_fun & "',"
    strSQL = strSQL & "'" & cep_fun & "',"
    strSQL = strSQL & F_GravaData(nasc_fun) & ","
    strSQL = strSQL & "'" & natural_fun & "',"
    strSQL = strSQL & "'" & rg_fun & "',"
    strSQL = strSQL & "'" & cpf_fun & "',"
    strSQL = strSQL & "'" & sexo_fun & "',"
    strSQL = strSQL & "'" & conjuge_fun & "',"
    strSQL = strSQL & "'" & pai_fun & "',"
    strSQL = strSQL & "'" & mae_fun & "',"
    strSQL = strSQL & "'" & endpais_fun & "',"
    strSQL = strSQL & "'" & obs_fun & "',"
    strSQL = strSQL & "'" & cod_Car & "') "
    conn.Execute strSQL
    F_MigraFun = sCod
    Exit Function
trata_Erro:
    MostraErro
End Function

Function F_AlteraFun(cod_fun, nome_fun As String, guerra_fun As String, endereco_fun As String, bairro_fun As String, _
    cidade_fun As String, uf_fun As String, cep_fun As String, nasc_fun As String, natural_fun As String, _
    rg_fun As String, cpf_fun As String, sexo_fun As String, conjuge_fun As String, pai_fun As String, _
    mae_fun As String, endpais_fun As String, obs_fun As String, status_fun As String, cod_Car As String, telefone_fun As String) As Boolean
    
    Dim strSQL As String
    
    On Error GoTo trata_Erro
    strSQL = "update tab_funcionarios set     "
    strSQL = strSQL & "nome_fun='" & nome_fun & "',"
    strSQL = strSQL & "GUERRA_fun='" & guerra_fun & "',"
    strSQL = strSQL & "endereco_fun='" & endereco_fun & "',"
    strSQL = strSQL & "bairro_fun='" & bairro_fun & "',"
    strSQL = strSQL & "cidade_fun='" & cidade_fun & "',"
    strSQL = strSQL & "uf_fun='" & uf_fun & "',"
    strSQL = strSQL & "cep_fun='" & cep_fun & "',"
    strSQL = strSQL & "nasc_fun=" & F_GravaData(nasc_fun) & ","
    strSQL = strSQL & "natural_fun='" & natural_fun & "',"
    strSQL = strSQL & "rg_fun='" & rg_fun & "',"
    strSQL = strSQL & "cpf_fun='" & cpf_fun & "',"
    strSQL = strSQL & "sexo_fun='" & sexo_fun & "',"
    strSQL = strSQL & "conjuge_fun='" & conjuge_fun & "',"
    strSQL = strSQL & "pai_fun='" & pai_fun & "',"
    strSQL = strSQL & "mae_fun='" & mae_fun & "',"
    strSQL = strSQL & "endpais_fun='" & endpais_fun & "',"
    strSQL = strSQL & "obs_fun= '" & obs_fun & "', "
    strSQL = strSQL & "status_fun='" & status_fun & "',"
    strSQL = strSQL & "telefone_fun='" & telefone_fun & "',"
    strSQL = strSQL & "cod_car='" & cod_Car & "' "
    strSQL = strSQL & " where cod_fun='" & cod_fun & "' "
    conn.Execute strSQL
    F_AlteraFun = True
    Exit Function
trata_Erro:
    MostraErro
End Function

Function F_SelectVendedor(sExpr As String) As funcionario.funcionario
    
    frmSelect.sSelect = "select cod_fun as Cod,nome_fun as nome,guerra_fun as Chamado,nome_car as cargo, " & _
        "venda_car as Venda from tab_funcionarios as tf, tab_cargos as tc " & _
        " where tc.cod_car=tf.cod_car " & _
         " and status_fun=1"
         
    frmSelect.cCriterio.Add " and nome_fun like '%" & sExpr & "%'"
    frmSelect.cCriterio.Add " and cod_fun like '%" & sExpr & "%'"
    frmSelect.sOrd = " order by nome_fun "
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_SelectVendedor.cod_fun = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 1
        F_SelectVendedor.nome_fun = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 2
        F_SelectVendedor.guerra_fun = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 3
        F_SelectVendedor.nome_car = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 4
        F_SelectVendedor.venda_car = frmSelect.msflxgrd
        
        
    End If
    Unload frmSelect
    

End Function
Function F_SelectCobrador(sExpr As String) As funcionario.funcionario
    
    frmSelect.sSelect = "select cod_fun as Cod,nome_fun as nome,guerra_fun as Chamado,nome_car as cargo, " & _
        "venda_car as Venda from tab_funcionarios as tf, tab_cargos as tc " & _
        " where tc.cod_car=tf.cod_car and cobranca_car='S'" & _
         " and status_fun=1"
        
    frmSelect.cCriterio.Add " and nome_fun like '%" & sExpr & "%'"
    frmSelect.cCriterio.Add " and cod_fun like '%" & sExpr & "%'"
    frmSelect.sOrd = " order by nome_fun "
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_SelectCobrador.cod_fun = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 1
        F_SelectCobrador.nome_fun = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 2
        F_SelectCobrador.guerra_fun = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 3
        F_SelectCobrador.nome_car = frmSelect.msflxgrd
        frmSelect.msflxgrd.col = 4
        F_SelectCobrador.venda_car = frmSelect.msflxgrd
        
        
    End If
    Unload frmSelect
    

End Function


Function F_SelectFuncionario(sExpr As String) As funcionario.funcionario
    On Error GoTo trata_Erro
'    frmSelect.sSelect = "select cod_fun as Cod,nome_fun as nome,guerra_fun as Chamado,nome_car as cargo, " & _
'        "venda_car as Venda from tab_funcionarios as tf, tab_cargos as tc " & _
'        " where tc.cod_car=tf.cod_car " & _
'         " and status_fun=1"
'
'    frmSelect.cCriterio.Add " and nome_fun like '%" & sExpr & "%'"
'    frmSelect.cCriterio.Add " and cod_fun like '%" & sExpr & "%'"
'    frmSelect.sOrd = " order by nome_fun "
'    frmSelect.txtbxExpr.Text = expr
    frmEdiFun.tag = "x"
    frmEdiFun.txtbxBuscar.Text = sExpr
    
    frmEdiFun.Show 1
    
    If frmEdiFun.tag <> "" Then
        
        F_SelectFuncionario.cod_fun = frmEdiFun.ado!cod_fun
        
        F_SelectFuncionario.nome_fun = frmEdiFun.ado!nome_fun
        
        F_SelectFuncionario.guerra_fun = frmEdiFun.ado!guerra_fun
        F_SelectFuncionario.nome_car = frmEdiFun.ado!nome_car
        F_SelectFuncionario.venda_car = frmEdiFun.ado!venda_car
        
        
    End If
    Unload frmEdiFun
    Exit Function
trata_Erro:

End Function
Sub carregaFunc(obj As TextBox)
    Dim f As funcionario.funcionario
    
    f = funcionario.f_BuscaFun(obj.Text)
    If f.cod_fun <> "" Then
    
        obj.Text = f.nome_fun
        obj.tag = f.cod_fun
    End If
End Sub
Function f_BuscaFun(expr As String) As funcionario.funcionario
    frmEdiFun.tag = "x"
    frmEdiFun.txtbxBuscar.Text = expr
    frmEdiFun.Show 1
    If frmEdiFun.tag <> "" Then
        f_BuscaFun.cod_fun = frmEdiFun.ado!cod_fun
        f_BuscaFun.cod_Car = frmEdiFun.ado!cod_Car
        f_BuscaFun.nome_car = frmEdiFun.ado!nome_car
        f_BuscaFun.nome_fun = frmEdiFun.ado!nome_fun
        f_BuscaFun.guerra_fun = frmEdiFun.ado!guerra_fun
        f_BuscaFun.endereco_fun = frmEdiFun.ado!endereco_fun
        f_BuscaFun.cpf_fun = frmEdiFun.ado!cpf_fun
    End If
    Unload frmEdiFun
End Function
Function F_SelectFuncionarioCod(cod_fun As String) As funcionario.funcionario
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select cod_fun as Cod,nome_fun as nome,guerra_fun as Chamado,nome_car as cargo, " & _
        "venda_car as Venda from tab_funcionarios as tf, tab_cargos as tc " & _
        " where tc.cod_car=tf.cod_car and cod_fun='" & cod_fun & "'" & _
         " and status_fun=1"
    
    
    rPesq.Open strSQL, conn
    With rPesq
        If .RecordCount > 0 Then
            
            F_SelectFuncionarioCod.cod_fun = !cod
            
            F_SelectFuncionarioCod.nome_fun = !nome
            
            F_SelectFuncionarioCod.guerra_fun = !chamado
            
            F_SelectFuncionarioCod.nome_car = !nome
            
            F_SelectFuncionarioCod.venda_car = !venda
            
            
        End If
        .Close
    End With
        

End Function
Sub P_DefUsu(cod_usu As String)
    frmAddPermUsu.txtbxNomeUsu.tag = cod_usu
    frmAddPermUsu.txtbxNomeUsu.Text = cod_usu
    frmAddPermUsu.Show 1
End Sub


Function F_VeririficaPermissao(grupo As String, permissao As String, cod_usu As String) As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select count(distinct tp.cod_per) as total from tab_permissoes as tp,tab_permgru as tg, tab_permusu as tu "
    strSQL = strSQL & " where tg.cod_gru=tp.cod_gru "
    strSQL = strSQL & " and tp.cod_per=tu.cod_per"
    strSQL = strSQL & " and tu.cod_fun='" & cod_usu & "' "
    rPesq.Open strSQL, conn
    With rPesq
        If .RecordCount = 0 Then
            F_VeririficaPermissao = False
        ElseIf !total = 0 Then
            F_VeririficaPermissao = False
        Else
            F_VeririficaPermissao = True
        End If
        
        .Close
    End With
End Function
