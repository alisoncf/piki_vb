Attribute VB_Name = "Clientes"
Type cliente
    cod_cli As String
    tipo_cli As String
    nome_Cli As String
    endereco_cli As String
    email_cli As String
    cnpj_Cli As String
    cpf_cli As String
    ie_Cli As String
    telefone_cli As String
    desconto_cli As Currency
    credito_cli As Currency
    uf_cli As String
    bairro_cli As String
    cidade_cli As String
    cep_cli As String
    inscprodrur_cli As String
    complemento As String
    numero As String
    Logradouro As String
    c As Cidades.Cidade
    status_cli As String
    limite_cli As Currency
    formaPag As FormasPag.FormasPag
End Type



Type grupoCli
    cod_gru As String
    nome_gru As String
End Type

Sub P_histCli(cod_cli As String, Optional dataIni As String = "", Optional dataFim As String = "")
    On Error GoTo trata_Erro
    frmParcelasTer.txtbxNomeCli.tag = cod_cli
    If dataFim <> "" Then
        frmParcelasTer.txtAte.text = dataFim
    End If
    If dataIni <> "" Then
        frmParcelasTer.txtDe.text = dataIni
    End If
    frmParcelasTer.tag = cod_cli
    
    frmParcelasTer.Show 0, mdiFrm
Exit Sub
trata_Erro:
    P_MostraErro
End Sub







Function F_GravaCli(nome_Cli As String, datanasc_cli As String, cpf_cli As String, rg_cli As String, _
    endereco_cli As String, bairro_cli As String, cidade_cli As String, uf_cli As String, cep_cli As String, _
    obs_cli As String, mae_cli As String, pai_cli As String, mens_cli As String, limite_cli As String, _
    cod_set As String, localtrab_cli As String, endtrab_cli As String, _
    setortrab_cli As String, teltrab_cli As String, ceptrab_cli As String, cod_gru As String, cobranca_cli As String, cod_org As String, sexo_cli As String) As String
    Dim strSQL As String
    
    Dim sCod As String
    sCod = MaxKeySQL(TAB_CLIENTES, "cod_cli", "", "0000#")
    
    
    strSQL = "insert into " & TAB_CLIENTES & "(cod_cli,nome_cli,datanasc_cli,cpf_cli,rg_cli,"
    strSQL = strSQL & "endereco_Cli,bairro_cli,cidade_Cli,"
    strSQL = strSQL & "uf_cli,cep_cli,"
    strSQL = strSQL & "data_cli,obs_cli,mae_cli,pai_cli,mens_cli,limite_cli,"
    strSQL = strSQL & "status_cli,cod_set,"
    strSQL = strSQL & "localtrab_cli,endtrab_cli,setortrab_cli,teltrab_cli,ceptrab_cli,cobranca_cli,cod_gru,sexo_cli,cod_org) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_Cli & "',"
    strSQL = strSQL & F_GravaData(datanasc_cli) & ","
    strSQL = strSQL & "'" & cpf_cli & "',"
    strSQL = strSQL & "'" & rg_cli & "',"
    strSQL = strSQL & "'" & endereco_cli & "',"
    strSQL = strSQL & "'" & bairro_cli & "',"
    strSQL = strSQL & "'" & cidade_cli & "',"
    strSQL = strSQL & "'" & uf_cli & "',"
    strSQL = strSQL & "'" & cep_cli & "',"
    
    
    
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & obs_cli & "',"
    strSQL = strSQL & "'" & mae_cli & "',"
    strSQL = strSQL & "'" & pai_cli & "',"
    
    strSQL = strSQL & "'" & mens_cli & "',"
    
    strSQL = strSQL & F_GravaDec(limite_cli) & ","
    strSQL = strSQL & "'1',"
    strSQL = strSQL & "'" & cod_set & "',"
    strSQL = strSQL & "'" & localtrab_cli & "',"
    strSQL = strSQL & "'" & endtrab_cli & "',"
    strSQL = strSQL & "'" & setortrab_cli & "',"
    strSQL = strSQL & "'" & teltrab_cli & "',"
    strSQL = strSQL & "'" & ceptrab_cli & "',"
    strSQL = strSQL & "'" & cobranca_cli & "',"
    strSQL = strSQL & "'" & cod_gru & "',"
    strSQL = strSQL & "'" & sexo_cli & "',"
    strSQL = strSQL & "'" & cod_org & "')"
    conn.Execute strSQL
    F_GravaCli = sCod
    
End Function
Function F_MigraCli(cod_cli As String, nome_Cli As String, datanasc_cli As String, cpf_cli As String, rg_cli As String, _
    endereco_cli As String, bairro_cli As String, cidade_cli As String, uf_cli As String, cep_cli As String, _
    obs_cli As String, mae_cli As String, pai_cli As String, mens_cli As String, limite_cli As String, _
    cod_set As String, localtrab_cli As String, endtrab_cli As String, _
    setortrab_cli As String, teltrab_cli As String, ceptrab_cli As String) As String
    Dim strSQL As String
    
    Dim sCod As String
    sCod = cod_cli
    
    
    strSQL = "insert into " & TAB_CLIENTES & "(cod_cli,nome_cli,datanasc_cli,cpf_cli,rg_cli,"
    strSQL = strSQL & "endereco_Cli,bairro_cli,cidade_Cli,"
    strSQL = strSQL & "uf_cli,cep_cli,"
    strSQL = strSQL & "data_cli,obs_cli,mae_cli,pai_cli,mens_cli,limite_cli,"
    strSQL = strSQL & "status_cli,cod_set,"
    strSQL = strSQL & "localtrab_cli,endtrab_cli,setortrab_cli,teltrab_cli,ceptrab_cli) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_Cli & "',"
    strSQL = strSQL & F_GravaData(datanasc_cli) & ","
    strSQL = strSQL & "'" & cpf_cli & "',"
    strSQL = strSQL & "'" & rg_cli & "',"
    strSQL = strSQL & "'" & endereco_cli & "',"
    strSQL = strSQL & "'" & bairro_cli & "',"
    strSQL = strSQL & "'" & cidade_cli & "',"
    strSQL = strSQL & "'" & uf_cli & "',"
    strSQL = strSQL & "'" & cep_cli & "',"
    
    
    
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & obs_cli & "',"
    strSQL = strSQL & "'" & mae_cli & "',"
    strSQL = strSQL & "'" & pai_cli & "',"
    
    strSQL = strSQL & "'" & mens_cli & "',"
    
    strSQL = strSQL & F_GravaDec(limite_cli) & ","
    strSQL = strSQL & "'1',"
    strSQL = strSQL & "'" & cod_set & "',"
    strSQL = strSQL & "'" & localtrab_cli & "',"
    strSQL = strSQL & "'" & endtrab_cli & "',"
    strSQL = strSQL & "'" & setortrab_cli & "',"
    strSQL = strSQL & "'" & teltrab_cli & "',"
    strSQL = strSQL & "'" & ceptrab_cli & "')"
    On Error Resume Next
    conn.Execute strSQL
    F_MigraCli = sCod
    
End Function

Function F_AlteraCli(cod_cli, nome_Cli As String, datanasc_cli As String, cpf_cli As String, rg_cli As String, _
    endereco_cli As String, bairro_cli As String, cidade_cli As String, uf_cli As String, cep_cli As String, _
    obs_cli As String, mae_cli As String, pai_cli As String, mens_cli As String, _
    limite_cli As String, status_cli As String, cod_set As String, localtrab_cli As String, endtrab_cli As String, _
    setortrab_cli As String, teltrab_cli As String, ceptrab_cli As String, cod_gru As String, cobranca_cli As String, _
    cod_org As String, sexo_cli As String) As Boolean
    Dim strSQL As String
    
    
    
    On Error GoTo trata_Erro
    
    strSQL = "update " & TAB_CLIENTES & " set "
    strSQL = strSQL & "nome_cli='" & nome_Cli & "',"
    strSQL = strSQL & "datanasc_cli=" & F_GravaData(datanasc_cli) & ","
    strSQL = strSQL & "cpf_cli='" & cpf_cli & "',"
    strSQL = strSQL & "rg_cli='" & rg_cli & "',"
    strSQL = strSQL & "endereco_cli='" & endereco_cli & "',"
    strSQL = strSQL & "bairro_cli='" & bairro_cli & "',"
    strSQL = strSQL & "cidade_cli='" & cidade_cli & "',"
    strSQL = strSQL & "uf_cli='" & uf_cli & "',"
    strSQL = strSQL & "cep_cli='" & cep_cli & "',"
    strSQL = strSQL & "cod_set='" & cod_set & "',"
    
    strSQL = strSQL & "cod_gru='" & cod_gru & "',"
    
    strSQL = strSQL & "obs_cli='" & obs_cli & "',"
    strSQL = strSQL & "mae_cli='" & mae_cli & "',"
    strSQL = strSQL & "pai_cli='" & pai_cli & "',"
    
    strSQL = strSQL & "mens_cli='" & mens_cli & "',"
    
    strSQL = strSQL & "limite_cli=" & F_GravaDec(limite_cli) & ","
    strSQL = strSQL & "status_cli='" & status_cli & "', "
    strSQL = strSQL & "localtrab_cli='" & localtrab_cli & "',"
    strSQL = strSQL & "endtrab_cli='" & endtrab_cli & "',"
    strSQL = strSQL & "setortrab_cli='" & setortrab_cli & "',"
    strSQL = strSQL & "teltrab_cli='" & teltrab_cli & "',"
    strSQL = strSQL & "cobranca_cli='" & cobranca_cli & "',"
    strSQL = strSQL & "cod_org='" & cod_org & "',"
    strSQL = strSQL & "cod_org='" & cod_org & "',"
    strSQL = strSQL & "dataalt_cli=" & F_GravaData(Date) & ","
    strSQL = strSQL & "sexo_cli='" & sexo_cli & "',"
    strSQL = strSQL & "ceptrab_cli='" & ceptrab_cli & "' "
    
    strSQL = strSQL & " where cod_cli='" & cod_cli & "'"
    conn.Execute strSQL
    F_AlteraCli = True
    Exit Function
trata_Erro:
    F_AlteraCli = False
End Function


Function F_GravaTelefone(cod_cli As String, num_tel As String, tipo_tel As String, Optional tipo_ter As String = "C") As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    F_GravaTelefone = False
    strSQL = "insert into tab_telefones(cod_cli,num_tel,tipo_tel,tipo_ter)values("
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & "'" & num_tel & "',"
    strSQL = strSQL & "'" & tipo_tel & "',"
    strSQL = strSQL & "'" & tipo_ter & "')"
    conn.Execute strSQL
    F_GravaTelefone = True
    
    Exit Function
trata_Erro:
    F_GravaTelefone = False
End Function
Function F_GravaAutorizados(cod_cli As String, nome_aut As String, cod_tip As String) As String
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    sCod = MaxKeySQL("tab_autorizados", "seq_aut", "cod_cli='" & cod_cli & "'", "")
    
            strSQL = "insert into tab_autorizados(cod_cli,seq_aut,nome_aut,cod_tip)values("
            strSQL = strSQL & "'" & cod_cli & "',"
            strSQL = strSQL & sCod & ","
            
            strSQL = strSQL & "'" & nome_aut & "',"
            
            
            strSQL = strSQL & "'" & cod_tip & "')"
            conn.Execute strSQL
    F_GravaAutorizados = sCod
    
    Exit Function
trata_Erro:
    F_GravaAutorizados = ""
End Function


Function F_BuscaTel(cod_cli As String, tipo_ter As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select * from tab_telefones where cod_cli='" & cod_cli & "' and tipo_ter='" & tipo_ter & "'"
    rPesq.Open strSQL, conn_client
    With rPesq
        Do While Not .EOF
            F_BuscaTel = F_BuscaTel + !num_tel
            .MoveNext
            If Not .EOF Then
                F_BuscaTel = F_BuscaTel & " / "
            End If
        Loop
        .Close
    End With
    Exit Function
trata_Erro:
    
End Function

Function F_BuscaLimiteEmUsoCli(cod_cli As String) As Currency
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select sum(ti.reman_itd) as total from " & TAB_ITDOC & " as ti, " & TAB_DOC & " as td "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc and td.cod_ter='" & cod_cli & "' "
    strSQL = strSQL & " and td.tipo_doc='R'"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            If IsNull(!total) Then
                F_BuscaLimiteEmUsoCli = 0
            Else
            
                F_BuscaLimiteEmUsoCli = !total
            End If
        Else
            F_BuscaLimiteEmUsoCli = 0
        End If
        .Close
    End With
End Function
Function F_BuscaLimiteEmUsoRep(cod_Rep As String) As Currency
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select sum(reman_con*unit_con) as total from tab_consignacoes  "
    strSQL = strSQL & " where cod_rep='" & cod_Rep & "'    "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount > 0 Then
            If IsNull(!total) Then
                F_BuscaLimiteEmUsoRep = 0
            Else
            
                F_BuscaLimiteEmUsoRep = !total
            End If
        Else
            F_BuscaLimiteEmUsoRep = 0
        End If
        .Close
    End With
End Function


Function F_BuscaGrupoCli() As grupoCli
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "select cod_gru as cod,nome_gru as nome from tab_grupocli "
    
    
    
    frmSelect.sSelect = strSQL
    
    Set frmSelect.cCriterio = Nothing
    frmSelect.cCriterio.Add "where cod_gru>-1  "
    frmSelect.sOrd = " order by nome_gru"
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        F_BuscaGrupoCli.cod_gru = frmSelect.msflxgrd.text
        frmSelect.msflxgrd.col = 1
        F_BuscaGrupoCli.nome_gru = frmSelect.msflxgrd.text
    End If
    Unload frmSelect
Exit Function
trata_Erro:
    MostraErro
End Function
Sub P_ResumoDeb(cod As String, Optional tipo_ter As String = "C")
    P_LocalizaOperacao tipo_ter, frmResumoDeb.cmbxTipoTer
    
    frmResumoDeb.tag = cod
    frmResumoDeb.txtbxTerceiro.text = cod
    frmResumoDeb.txtbxTerceiro.tag = cod
    frmResumoDeb.Show 1
    
End Sub
Sub P_ClienteSpc(cod_cli As String, Optional optnShow As Integer = 1)
    On Error Resume Next
    frmClienteSpc.tag = cod_cli
    frmClienteSpc.Show optnShow
    
End Sub
Function F_BuscaCliente(Optional busca As String = "") As cliente
    On Error Resume Next
    
    frmSelectCli.txtbxNome.text = busca
    frmSelectCli.Show 1
    If frmSelectCli.tag <> "" Then
        F_BuscaCliente.cod_cli = frmSelectCli.adoR!cod_cli
        F_BuscaCliente.nome_Cli = frmSelectCli.adoR!nome_Cli
        F_BuscaCliente.cpf_cli = Format(frmSelectCli.adoR!cpf_cli)
        F_BuscaCliente.telefone_cli = TiraNull(frmSelectCli.adoR!telefone_cli)
        F_BuscaCliente.endereco_cli = frmSelectCli.adoR!Logradouro_cli & " " & frmSelectCli.adoR!NUMERO_cli & " " & frmSelectCli.adoR!COMPLEMENTO_cli
        F_BuscaCliente.ie_Cli = frmSelectCli.adoR!rg_cli
        F_BuscaCliente.cnpj_Cli = frmSelectCli.adoR!cpf
        F_BuscaCliente.limite_cli = frmSelectCli.adoR!limite_cli
        F_BuscaCliente.bairro_cli = frmSelectCli.adoR!bairro_cli
        F_BuscaCliente.cidade_cli = frmSelectCli.adoR!cidade_cli
        F_BuscaCliente.cep_cli = frmSelectCli.adoR!cep_cli
        F_BuscaCliente.uf_cli = frmSelectCli.adoR!uf_cli
        F_BuscaCliente.tipo_cli = frmSelectCli.adoR!tipo_cli
        F_BuscaCliente.inscprodrur_cli = frmSelectCli.adoR!inscprodrur_cli
        F_BuscaCliente.Logradouro = frmSelectCli.adoR!Logradouro_cli
        F_BuscaCliente.cnpj_Cli = F_SoNumero(frmSelectCli.adoR!cnpj_Cli)
        F_BuscaCliente.complemento = frmSelectCli.adoR!COMPLEMENTO_cli
        F_BuscaCliente.numero = frmSelectCli.adoR!NUMERO_cli
        F_BuscaCliente.c.cod_cid = frmSelectCli.adoR!cod_cid
        F_BuscaCliente.c.uf_cid = frmSelectCli.adoR!estado_cid
        F_BuscaCliente.c.nome_cid = frmSelectCli.adoR!nome_cid
        F_BuscaCliente.c.uf_cid = frmSelectCli.adoR!uf_cid
        F_BuscaCliente.desconto_cli = frmSelectCli.adoR!desconto_cli
        F_BuscaCliente.c.cod_cid = frmSelectCli.adoR!cod_cid
        F_BuscaCliente.c.cod_mun = frmSelectCli.adoR!cod_cid
        F_BuscaCliente.c.nome_cid = frmSelectCli.adoR!nome_cid
        F_BuscaCliente.email_cli = TiraNull(adoR!email_cli)
        F_BuscaCliente.c.uf_cid = frmSelectCli.adoR!uf_cid
        F_BuscaCliente.ie_Cli = F_SoNumero(frmSelectCli.adoR!ie_Cli)
        F_BuscaCliente.formaPag.cod_For = frmSelectCli.adoR!cod_For
        F_BuscaCliente.formaPag.nome_For = frmSelectCli.adoR!nome_For
        
    End If
    Unload frmSelectCli
End Function

Sub P_ParcelaSpc(cod_cli As String, seq_itd As String, valor As String)
    On Error Resume Next
    frmClienteSpc.tag = cod_cli
    frmClienteSpc.txtbxRazao.text = F_BuscaUmCampo("nome_cli", TAB_CLIENTES, "cod_cli='" & cod_cli & "'")
    frmClienteSpc.txtbxDoc.text = seq_itd
    frmClienteSpc.txtbxValor.text = valor
    frmClienteSpc.Show 0
    
End Sub
Sub P_GravaClienteSpc(cod_cli As String, data_spc As String, valor_spc As String, doc_spc As String, tipo_doc As String, tipo_ter As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_clientespc(seq_spc,cod_cli,data_spc,valor_spc,doc_spc,cod_usu,dataacao_spc,tipo_ter,tipo_doc)values("
    strSQL = strSQL & "'',"
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & F_GravaData(data_spc) & ","
    strSQL = strSQL & F_GravaDec(valor_spc) & ","
    strSQL = strSQL & "'" & doc_spc & "',"
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & tipo_ter & "',"
    strSQL = strSQL & "'" & tipo_doc & "')"
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    MostraErro
End Sub
Sub P_AlteraClienteSpc(seq_spc As String, data_spc As String, valor_spc As String, doc_spc As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_clientespc set "
    
    
    strSQL = strSQL & "data_spc=" & F_GravaData(data_spc) & ","
    strSQL = strSQL & "valor_spc=" & F_GravaDec(valor_spc) & ","
    strSQL = strSQL & "cod_spc='" & doc_spc & "',"
    strSQL = strSQL & "cod_usu='" & strCod_Usu & "' "
    strSQL = strSQL & "dataacao_spc=" & F_GravaDataTempo(Now) & " "
    strSQL = strSQL & " where seq_spc=" & seq_spc
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    MostraErro
    
End Sub
Sub P_BaixaClienteSpc(seq_spc As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_clientespc set "
    
    
    strSQL = strSQL & "baixa_spc=" & F_GravaData(Date)
    strSQL = strSQL & " where seq_spc=" & seq_spc
    conn.Execute strSQL
    Exit Sub
trata_Erro:
    MostraErro
    
End Sub
Sub P_ListaClientesNaoDevedoresSPC()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select tc.*"
End Sub
Sub P_ParcelaBaixadaSPC(seq_itd As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select ti.* from  " & TAB_ITDOC & "  as ti,  " & TAB_DOC & "  as td, tab_clientespc as tc where seq_itd='" & seq_itd & "' "
    strSQL = strSQL & " and tc.cod_cli=td.cod_ter "
    strSQL = strSQL & " and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and baixa_spc='0000-00-00' "
    rPesq.Open strSQL, conn_client
    With rPesq
        If .RecordCount > 0 Then
            MsgBox "O cliente encontra-se negativado, lembre-se de Baixá-lo no SPC", vbExclamation
        End If
        .Close
    End With
    Exit Sub
trata_Erro:
End Sub


Function F_SelectItensCompradosVendas(cod_cli As String, Optional dataI As String = "", _
    Optional dataF As String = "", Optional cod_mar As String = "", Optional cod_cat As String = "", _
    Optional cod_pro As String = "") As String
    Dim strSQL As String
    strSQL = "select tfo.nome_for,nome_mar,nome_cat,tc.nome_cli, ti.*,((ti.qtd_ite*unit_ite-desconto_ite)-(ti.devolvido_ite*unit_ite-desconto_ite)) as total, tv.num_ven,tv.data_ven,codfab_pro,"
    strSQL = strSQL & "tv.valor_ven,desconto_ven, tf.guerra_fun "
    strSQL = strSQL & " from " & TAB_CLIENTES & " as tc,"
    strSQL = strSQL & TAB_VENDAS & " as tv , tab_funcionarios as tf,tab_formaspag as tfo,  "
    strSQL = strSQL & TAB_ITVENDAS & " as ti inner join " & TAB_PRODUTOS & " as tp on tp.cod_pro=ti.cod_pro "
    strSQL = strSQL & " inner join " & TAB_MARCAS & " as tm on tm.cod_mar=tp.cod_mar "
    strSQL = strSQL & " inner join " & TAB_CATEGORIAS & " as tcat on tcat.cod_Cat=tp.cod_cat  "
    
    
    strSQL = strSQL & " where tf.cod_fun=tv.abrevenda_usu "
    strSQL = strSQL & " and ti.num_ven=tv.num_ven "
    strSQL = strSQL & " and tfo.cod_for=tv.cod_for "
    
    strSQL = strSQL & " and situacao_ven=4 "
    
    strSQL = strSQL & " and tc.cod_cli=tv.cod_cli "
    If abreVendaUsu <> "" Then
        strSQL = strSQL & " and abrevenda_usu='" & abreVendaUsu & "' "
    End If
    
    If cod_cli <> "" Then
        strSQL = strSQL & " and tv.cod_cli='" & cod_cli & "' "
    End If
    If cod_pro <> "" Then
        strSQL = strSQL & " and ti.cod_pro='" & cod_pro & "' "
    End If
    
    If dataI <> "" Then
        strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(dataF & " 23:59:59")
    End If
    F_SelectItensCompradosVendas = strSQL
End Function
Function F_BuscaClienteCPFCNPJ(cpfcnpj As String) As cliente
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo trata_Erro
    
    
    strSQL = "select cli.*,cid.nome_cid,cid.cod_cid from " & TAB_CLIENTES & " as cli left join " & TAB_CIDADES & " as cid "
    strSQL = strSQL & " on cli.cod_Cid=cid.cod_cid where (cpf_cli = '" & cpfcnpj & "'"
    strSQL = strSQL & " or cnpj_cli = '" & cpfcnpj & "%')"
    
    
    
    
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    
    With rPesq
        If rPesq.RecordCount > 0 Then
            F_BuscaClienteCPFCNPJ.cod_cli = rPesq!cod_cli

            F_BuscaClienteCPFCNPJ.status_cli = Format(rPesq!status_cli)
            F_BuscaClienteCPFCNPJ.nome_Cli = rPesq!nome_Cli
            F_BuscaClienteCPFCNPJ.credito_cli = rPesq!credito_cli
            F_BuscaClienteCPFCNPJ.email_cli = TiraNull(rPesq!email_cli)
            F_BuscaClienteCPFCNPJ.endereco_cli = rPesq!Logradouro_cli & " N°" & rPesq!NUMERO_cli & " " & rPesq!COMPLEMENTO_cli
            F_BuscaClienteCPFCNPJ.ie_Cli = rPesq!rg_cli
            If Not IsNull(rPesq!cpf_cli) Then
                F_BuscaClienteCPFCNPJ.cnpj_Cli = (rPesq!cpf_cli)
            End If
            If Not IsNull(rPesq!cnpj_Cli) Then
                F_BuscaClienteCPFCNPJ.cnpj_Cli = (rPesq!cnpj_Cli)
            End If
            F_BuscaClienteCPFCNPJ.limite_cli = rPesq!limite_cli
            F_BuscaClienteCPFCNPJ.bairro_cli = rPesq!bairro_cli
            F_BuscaClienteCPFCNPJ.cidade_cli = rPesq!nome_cid
            F_BuscaClienteCPFCNPJ.cep_cli = rPesq!cep_cli
            F_BuscaClienteCPFCNPJ.uf_cli = rPesq!uf_cli
            F_BuscaClienteCPFCNPJ.telefone_cli = !telefone_cli
            F_BuscaClienteCPFCNPJ.Logradouro = !Logradouro_cli
            F_BuscaClienteCPFCNPJ.numero = !NUMERO_cli
            F_BuscaClienteCPFCNPJ.complemento = !COMPLEMENTO_cli
            
            
        End If
        .Close
    End With
        
    Exit Function
trata_Erro:
    MsgBox Err.Description
End Function

Public Function F_BuscaClienteCodigo(cod_cli As String) As cliente
    'On Error Resume Next
    Dim sTel As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    If cod_cli = "" Then cod_cli = "0"
    
    
    strSQL = "select cli.*,cid.nome_cid,cid.cod_cid from " & TAB_CLIENTES & " as cli left join " & TAB_CIDADES & " as cid "
    strSQL = strSQL & " on cli.cod_Cid=cid.cod_cid where cod_cli=" & cod_cli & ""
            
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    
    
    If rPesq.RecordCount > 0 Then
        F_BuscaClienteCodigo.cod_cli = rPesq!cod_cli

        F_BuscaClienteCodigo.status_cli = Format(rPesq!status_cli)
        F_BuscaClienteCodigo.nome_Cli = rPesq!nome_Cli
        F_BuscaClienteCodigo.credito_cli = rPesq!credito_cli
        F_BuscaClienteCodigo.endereco_cli = rPesq!Logradouro_cli & " N°" & rPesq!NUMERO_cli & " " & rPesq!COMPLEMENTO_cli
        F_BuscaClienteCodigo.ie_Cli = rPesq!rg_cli
        F_BuscaClienteCodigo.email_cli = TiraNull(rPesq!email_cli)
        If Not IsNull(rPesq!cpf_cli) Then
            F_BuscaClienteCodigo.cnpj_Cli = (rPesq!cpf_cli)
        End If
        If Not IsNull(rPesq!cnpj_Cli) Then
            F_BuscaClienteCodigo.cnpj_Cli = (rPesq!cnpj_Cli)
        End If
        F_BuscaClienteCodigo.limite_cli = rPesq!limite_cli
        F_BuscaClienteCodigo.bairro_cli = rPesq!bairro_cli
        F_BuscaClienteCodigo.cidade_cli = rPesq!nome_cid
        F_BuscaClienteCodigo.cep_cli = rPesq!cep_cli
        F_BuscaClienteCodigo.uf_cli = rPesq!uf_cli
        F_BuscaClienteCodigo.telefone_cli = rPesq!telefone_cli
        
        
        
        
    End If
    
End Function
Public Function F_BuscaClienteVenda(num_ven As String) As cliente
    'On Error Resume Next
    Dim sTel As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    
    strSQL = "select cli.*,cid.nome_cid,cid.cod_cid from " & TAB_CLIENTES & " as cli left join " & TAB_CIDADES & " as cid "
    strSQL = strSQL & " on cli.cod_Cid=cid.cod_cid "
    strSQL = strSQL & " inner join " & TAB_VENDAS & " as tv on cli.cod_cli=tv.cod_cli "
    strSQL = strSQL & " where tv.num_Ven ='" & num_ven & "'"
            
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    
    
    If rPesq.RecordCount > 0 Then
        F_BuscaClienteVenda.cod_cli = rPesq!cod_cli

        F_BuscaClienteVenda.status_cli = Format(rPesq!status_cli)
        F_BuscaClienteVenda.nome_Cli = rPesq!nome_Cli
        F_BuscaClienteVenda.credito_cli = rPesq!credito_cli
        F_BuscaClienteVenda.endereco_cli = rPesq!Logradouro_cli & " N°" & rPesq!NUMERO_cli & " " & rPesq!COMPLEMENTO_cli
        F_BuscaClienteVenda.ie_Cli = rPesq!rg_cli
        F_BuscaClienteVenda.email_cli = TiraNull(rPesq!email_cli)
        
        If Not IsNull(rPesq!cpf_cli) Then
            F_BuscaClienteVenda.cnpj_Cli = (rPesq!cpf_cli)
        End If
        If Not IsNull(rPesq!cnpj_Cli) Then
            F_BuscaClienteVenda.cnpj_Cli = (rPesq!cnpj_Cli)
        End If
        F_BuscaClienteVenda.limite_cli = rPesq!limite_cli
        F_BuscaClienteVenda.bairro_cli = rPesq!bairro_cli
        F_BuscaClienteVenda.cidade_cli = rPesq!nome_cid
        F_BuscaClienteVenda.cep_cli = rPesq!cep_cli
        F_BuscaClienteVenda.uf_cli = rPesq!uf_cli
        F_BuscaClienteVenda.telefone_cli = rPesq!telefone_cli
        
        
        
        
    End If
    
End Function


