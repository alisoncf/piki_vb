Attribute VB_Name = "Clientes"
Type tCliente
    cod_cli As String
    tipo_cli As String
    nome_Cli As String
    endereco_cli As String
    cnpj_Cli As String
    ie_Cli As String
    telefone_cli As String
    limite_cli As Currency
    uf_cli As String
    bairro_cli As String
    cidade_cli As String
    cep_cli As String
    inscprodrur_cli As String
End Type

Type Cliente
    cod_cli As String
    nome_Cli As String
    endereco_cli As String
    nasc_cli As String
    cnpj_Cli As String
    ie_Cli As String
    telefone_cli As String
    limite_cli As Currency
    uf_cli As String
    bairro_cli As String
    cidade_cli As String
    cep_cli As String
    
    credito_cli As Currency
    status_cli As String
    cod_gru As String
    nome_gru As String
    
    localtrab_cli As String
    endtrab_cli As String
    
    teltrab_cli As String
    ceptrab_cli As String
    datanasc_cli As String
    
End Type

Type grupoCli
    cod_gru As String
    nome_gru As String
End Type

Sub P_histCli(cod_cli As String)
    frmParcelasTer.txtbxNomeCli.Tag = cod_cli
    
    frmParcelasTer.Tag = cod_cli
    
    frmParcelasTer.Show 0
End Sub







Function F_GravaCli(nome_Cli As String, datanasc_cli As String, cpf_cli As String, rg_cli As String, _
    endereco_cli As String, bairro_cli As String, cidade_cli As String, uf_cli As String, cep_cli As String, _
    obs_cli As String, mae_cli As String, pai_cli As String, mens_cli As String, limite_cli As String, _
    cod_set As String, localtrab_cli As String, endtrab_cli As String, _
    setortrab_cli As String, teltrab_cli As String, ceptrab_cli As String, cod_gru As String, cobranca_cli As String, cod_org As String, sexo_cli As String) As String
    Dim strSQL As String
    
    Dim sCod As String
    sCod = MaxKeySQL("tab_clientes", "cod_cli", "", "0000#")
    
    
    strSQL = "insert into tab_clientes(cod_cli,nome_cli,datanasc_cli,cpf_cli,rg_cli,"
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
    banco.Execute strSQL
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
    
    
    strSQL = "insert into tab_clientes(cod_cli,nome_cli,datanasc_cli,cpf_cli,rg_cli,"
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
    banco.Execute strSQL
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
    
    strSQL = "update tab_clientes set "
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
    banco.Execute strSQL
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
    banco.Execute strSQL
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
            banco.Execute strSQL
    F_GravaAutorizados = sCod
    
    Exit Function
trata_Erro:
    F_GravaAutorizados = ""
End Function


Function F_BuscaTel(cod_cli As String, tipo_ter As String) As String
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro
    strSQL = "select * from tab_telefones where cod_cli='" & cod_cli & "' and tipo_ter='" & tipo_ter & "'"
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
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
    Dim rPesq As Recordset
    strSQL = "select sum(ti.reman_itd) as total from tab_itdoc as ti, tab_doc as td "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc and td.cod_ter='" & cod_cli & "' "
    strSQL = strSQL & " and td.tipo_doc='R'"
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
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
    Dim rPesq As Recordset
    strSQL = "select sum(reman_con*unit_con) as total from tab_consignacoes  "
    strSQL = strSQL & " where cod_rep='" & cod_Rep & "'    "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
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
    
    If frmSelect.Tag <> "" Then
        frmSelect.msflxGrd.Col = 0
        F_BuscaGrupoCli.cod_gru = frmSelect.msflxGrd.Text
        frmSelect.msflxGrd.Col = 1
        F_BuscaGrupoCli.nome_gru = frmSelect.msflxGrd.Text
    End If
    Unload frmSelect
Exit Function
trata_Erro:
    P_MostraErro
End Function
Sub P_ResumoDeb(cod As String, Optional tipo_ter As String = "C")
    P_LocalizaOperacao tipo_ter, frmResumoDeb.cmbxTipoTer
    
    frmResumoDeb.Tag = cod
    frmResumoDeb.txtbxTerceiro.Text = cod
    frmResumoDeb.txtbxTerceiro.Tag = cod
    frmResumoDeb.Show 1
    
End Sub
Sub P_ClienteSpc(cod_cli As String, Optional optnShow As Integer = 1)
    On Error Resume Next
    frmClienteSpc.Tag = cod_cli
    frmClienteSpc.Show optnShow
    
End Sub
Function F_BuscaCliente(tipo_ter As String, busca As String) As tCliente
    On Error Resume Next
    frmSelectCli.txtbxNome.Tag = tipo_ter
    frmSelectCli.txtbxNome.Text = busca
    frmSelectCli.Show 1
    If frmSelectCli.Tag <> "" Then
        F_BuscaCliente.cod_cli = frmSelectCli.dtBusca.Recordset!cod
        F_BuscaCliente.nome_Cli = frmSelectCli.dtBusca.Recordset!nome
        F_BuscaCliente.telefone_cli = frmSelectCli.dtBusca.Recordset!telefone
        F_BuscaCliente.endereco_cli = frmSelectCli.dtBusca.Recordset!Endereco
        F_BuscaCliente.ie_Cli = frmSelectCli.dtBusca.Recordset!rg_cli
        F_BuscaCliente.cnpj_Cli = frmSelectCli.dtBusca.Recordset!CPF
        F_BuscaCliente.limite_cli = frmSelectCli.dtBusca.Recordset!limite_cli
        F_BuscaCliente.bairro_cli = frmSelectCli.dtBusca.Recordset!bairro_cli
        F_BuscaCliente.cidade_cli = frmSelectCli.dtBusca.Recordset!cidade_cli
        F_BuscaCliente.cep_cli = frmSelectCli.dtBusca.Recordset!cep_cli
        F_BuscaCliente.uf_cli = frmSelectCli.dtBusca.Recordset!uf_cli
        F_BuscaCliente.tipo_cli = frmSelectCli.dtBusca.Recordset!tipo_cli
        F_BuscaCliente.inscprodrur_cli = frmSelectCli.dtBusca.Recordset!inscprodrur_cli
    End If
    Unload frmSelectCli
End Function

Sub P_ParcelaSpc(cod_cli As String, seq_itd As String, valor As String)
    On Error Resume Next
    frmClienteSpc.Tag = cod_cli
    frmClienteSpc.txtbxRazao.Text = F_BuscaUmCampo("nome_cli", "tab_clientes", "cod_cli='" & cod_cli & "'")
    frmClienteSpc.txtbxDoc.Text = seq_itd
    frmClienteSpc.txtbxValor.Text = valor
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
    banco.Execute strSQL
    Exit Sub
trata_Erro:
    P_MostraErro
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
    banco.Execute strSQL
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub
Sub P_BaixaClienteSpc(seq_spc As String)
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_clientespc set "
    
    
    strSQL = strSQL & "baixa_spc=" & F_GravaData(Date)
    strSQL = strSQL & " where seq_spc=" & seq_spc
    banco.Execute strSQL
    Exit Sub
trata_Erro:
    P_MostraErro
    
End Sub
Sub P_ListaClientesNaoDevedoresSPC()
    Dim strSQL As String
    Dim rPesq As Recordset
    strSQL = "select tc.*"
End Sub
Sub P_ParcelaBaixadaSPC(seq_itd As String)
    Dim strSQL As String
    Dim rPesq As Recordset
    On Error GoTo trata_Erro
    strSQL = "select ti.* from tab_itdoc as ti, tab_doc as td, tab_clientespc as tc where seq_itd='" & seq_itd & "' "
    strSQL = strSQL & " and tc.cod_cli=td.cod_ter "
    strSQL = strSQL & " and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and baixa_spc='0000-00-00' "
    Set rPesq = banco.OpenRecordset(strSQL, dbOpenSnapshot)
    With rPesq
        If .RecordCount > 0 Then
            MsgBox "O cliente encontra-se negativado, lembre-se de Baixá-lo no SPC", vbExclamation
        End If
        .Close
    End With
    Exit Sub
trata_Erro:
End Sub


