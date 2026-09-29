Attribute VB_Name = "Clientes"
Dim Tel() As telefone
Enum StatusCliente
    liberado = 1
    bloqueado = 2
End Enum
Type telefone
    num_tel As String
    tipo_tel As String
    cod_cli As String
End Type

Type grupoCli
    cod_gru As String
    nome_gru As String
    

End Type

Type Endereco
    cep As String
    Logradouro As String
    complemento As String
    bairro As String
    localidade As String
    UF As String
    ibge As String
End Type


Type cliente
    
    cod_cli As String
    nome_cli As String
    endereco_cli As String
    nasc_cli As String
    cnpj_Cli As String
    Logradouro As String
    numero As String
    complemento As String
    ie_Cli As String
    sexo_cli As String
    cpf_cli As String
    telefone_cli As String
    limite_cli As Currency
    uf_cli As String
    bairro_cli As String
    cidade_cli As String
    cep_cli As String
    c As Cidade
    t() As telefone
    credito_cli As Currency
    status_cli As String
    cod_gru As String
    nome_gru As String
    org As orgao.orgao
    set As setor
    localtrab_cli As String
    endtrab_cli As String
    setortrab As setor
    teltrab_cli As String
    ceptrab_cli As String
    datanasc_cli As String
    dataalt_cli As String
    email_cli As String
    mae_cli As String
    pai_cli As String
    
End Type


Function F_BuscaCliente(busca As String, Optional buscar_por_doc As Boolean = False) As cliente
    On Error GoTo trata_Erro
    Dim adoR As New ADODB.Recordset
    Dim sTel As String
    Dim frm As New frmSelectCli
    frm.txtbxNome.tag = "C"
    frm.txtbxNome.Text = busca
    frm.BuscarPorDoc = buscar_por_doc
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    
    frm.Show 1
    If frm.tag <> "" Then

        Set adoR = frm.adoR
        With adoR
            F_BuscaCliente.cod_cli = !cod_cli
            If IsNull(adoR!cod_cli) Then
                F_BuscaCliente.cod_cli = !cod
            End If
            If CStr(adoR!cod_cli) = "" Then
                F_BuscaCliente.cod_cli = !cod
            End If
            
            
            F_BuscaCliente.nasc_cli = Format(!datanasc_cli)
            F_BuscaCliente.status_cli = !status_cli
            F_BuscaCliente.nome_cli = !nome_cli
            F_BuscaCliente.credito_cli = !credito_cli
            F_BuscaCliente.endereco_cli = !Logradouro_cli & " N° " & !numero_cli & " " & !complemento_cli
            F_BuscaCliente.ie_Cli = !rg_cli
            F_BuscaCliente.cnpj_Cli = !cpf_cli
            F_BuscaCliente.cpf_cli = !cpf_cli
            F_BuscaCliente.limite_cli = !limite_cli
            F_BuscaCliente.bairro_cli = !nome_set
            F_BuscaCliente.cidade_cli = !nome_cid
            F_BuscaCliente.cep_cli = !cep_cli
            F_BuscaCliente.uf_cli = !uf_cid
            F_BuscaCliente.set.Cidade = !nome_cid
            F_BuscaCliente.set.cobrador_set = !cobrador_set
            F_BuscaCliente.set.cod = !cod_set
            F_BuscaCliente.set.nome = !nome_set
            F_BuscaCliente.set.tipo = !tipo_set
            F_BuscaCliente.set.UF = !uf_cid
            F_BuscaCliente.localtrab_cli = !localtrab_cli
            F_BuscaCliente.endtrab_cli = !endtrab_cli
            F_BuscaCliente.setortrab.cod = !setortrab_cli
            F_BuscaCliente.setortrab = F_BuscaSetorPorCodigo(!setortrab_cli)
            F_BuscaCliente.dataalt_cli = TiraNull(!dataalt_cli)
            
            
            F_BuscaCliente.org = F_BuscaOrgaoPorCod(!cod_org)
            
            
            
            
            
            
            
            
            
            strSQL = "select * from tab_telefones where cod_cli='" & !cod_cli & "' and tipo_ter='C'"
        End With
        
        abreObjetoAdo rPesq, strSQL
        With rPesq
            ReDim F_BuscaCliente.t(.RecordCount) As telefone
            Do While Not .EOF
                
                F_BuscaCliente.t(.AbsolutePosition).cod_cli = !cod_cli
                F_BuscaCliente.t(.AbsolutePosition).num_tel = !num_tel
                F_BuscaCliente.t(.AbsolutePosition).tipo_tel = !tipo_tel
                
                sTel = sTel & !num_tel
                
                .MoveNext
                If Not .EOF Then
                    sTel = sTel & " "
                End If
            Loop
        End With
        F_BuscaCliente.telefone_cli = sTel
        
    End If
    Unload frm
    Exit Function
trata_Erro:
    MsgBox "Erro no procedimento de buscar clientes - " & Err.Description & Chr(13) & Err.Source, vbCritical
End Function
Public Function F_BuscaClienteCodigo(cod_cli As String) As cliente
    'On Error Resume Next
    Dim sTel As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    
    strSQL = "select tc.*,tc.cod_cli,cod_org,datanasc_cli,status_cli,nome_cli,credito_cli,logradouro_cli,numero_cli,complemento_cli,endereco_cli,rg_cli,cpf_cli,bairro_cli,cep_cli,limite_cli,ts.cod_set,tipo_set,"
        strSQL = strSQL & " cod_cli as cod, nome_cid,datanasc_cli,uf_cid,ts.cod_cid,cobrador_set,ts.nome_set,nome_cli as nome,cpf_cli as cnpj,endereco_cli as endereco,  "
                strSQL = strSQL & "cpf_cli as cpf from tab_clientes as tc,tab_setores as ts, "
                strSQL = strSQL & " tab_cidades as ti "
                strSQL = strSQL & " where tc.cod_set=ts.cod_set and ti.cod_cid=ts.cod_cid"
                strSQL = strSQL & " and cod_cli='" & cod_cli & "' "
            
    FechaObjeto rPesq
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    
    If rPesq.RecordCount > 0 Then
        F_BuscaClienteCodigo.cod_cli = rPesq!cod_cli
        If Not IsNull(rPesq!datanasc_cli) Then
            F_BuscaClienteCodigo.datanasc_cli = rPesq!datanasc_cli
        End If
        F_BuscaClienteCodigo.status_cli = rPesq!status_cli
        F_BuscaClienteCodigo.nome_cli = rPesq!nome_cli
        F_BuscaClienteCodigo.sexo_cli = rPesq!sexo_cli
        F_BuscaClienteCodigo.credito_cli = rPesq!credito_cli
        F_BuscaClienteCodigo.endereco_cli = rPesq!Logradouro_cli & " N°" & rPesq!numero_cli & " " & rPesq!complemento_cli
        F_BuscaClienteCodigo.ie_Cli = TiraNull(rPesq!rg_cli)
        F_BuscaClienteCodigo.cnpj_Cli = TiraNull(rPesq!cpf_cli)
        F_BuscaClienteCodigo.limite_cli = rPesq!limite_cli
        F_BuscaClienteCodigo.bairro_cli = rPesq!nome_set
        F_BuscaClienteCodigo.cidade_cli = rPesq!nome_cid
        F_BuscaClienteCodigo.cep_cli = TiraNull(rPesq!cep_cli)
        F_BuscaClienteCodigo.uf_cli = TiraNull(rPesq!uf_cid)
        F_BuscaClienteCodigo.mae_cli = TiraNull(rPesq!mae_cli)
        F_BuscaClienteCodigo.pai_cli = TiraNull(rPesq!pai_cli)
        F_BuscaClienteCodigo.set.Cidade = rPesq!nome_cid
        F_BuscaClienteCodigo.set.cobrador_set = rPesq!cobrador_set
        F_BuscaClienteCodigo.set.cod = rPesq!cod_set
        F_BuscaClienteCodigo.set.nome = rPesq!nome_set
        F_BuscaClienteCodigo.set.tipo = rPesq!tipo_set
        F_BuscaClienteCodigo.set.UF = rPesq!uf_cid
        F_BuscaClienteCodigo.nasc_cli = TiraNull(rPesq!datanasc_cli)
        F_BuscaClienteCodigo.dataalt_cli = TiraNull(rPesq!dataalt_cli)
        If rPesq!cod_org <> "" Then
            F_BuscaClienteCodigo.org = F_BuscaOrgaoPorCod(TiraNull(rPesq!cod_org))
        End If
        
        strSQL = "select * from tab_telefones where cod_cli='" & rPesq!cod_cli & "' and tipo_ter='C'"
        FechaObjeto rPesq
        
        abreObjetoAdo rPesq, strSQL
        With rPesq
            ReDim F_BuscaClienteCodigo.t(.RecordCount) As telefone
            Do While Not .EOF
                
                F_BuscaClienteCodigo.t(.AbsolutePosition).cod_cli = !cod_cli
                F_BuscaClienteCodigo.t(.AbsolutePosition).num_tel = !num_tel
                F_BuscaClienteCodigo.t(.AbsolutePosition).tipo_tel = !tipo_tel
                
                sTel = sTel & !num_tel
                
                .MoveNext
                If Not .EOF Then
                    sTel = sTel & "/"
                End If
            Loop
            .Close
        End With
        F_BuscaClienteCodigo.telefone_cli = sTel
        
    End If
    
End Function


Function F_GravaCli(nome_cli As String, datanasc_cli As String, cpf_cli As String, rg_cli As String, _
    endereco_cli As String, Logradouro As String, numero As String, complemento As String, bairro_cli As String, cidade_cli As String, uf_cli As String, cep_cli As String, _
    obs_cli As String, mae_cli As String, pai_cli As String, mens_cli As String, limite_cli As String, _
    cod_set As String, localtrab_cli As String, endtrab_cli As String, _
    setortrab_cli As String, teltrab_cli As String, ceptrab_cli As String, cod_gru As String, cobranca_cli As String, cod_org As String, cod_cid As String, sexo_cli As String, Optional email As String = "", Optional codigo As String = "") As String
    Dim strSQL As String
    
    Dim sCod As String
    If cobranca_cli = "" Then
        cobranca_cli = "1"
    End If
    If codigo = "" Then
        sCod = MaxKeySQL("tab_clientes", "cod_cli", "", "00000#")
    Else
        sCod = codigo
    End If
    
    strSQL = "insert into tab_clientes(cod_cli,nome_cli,datanasc_cli,cpf_cli,rg_cli,"
    strSQL = strSQL & "logradouro_cli,numero_cli,complemento_cli,"
    strSQL = strSQL & "endereco_Cli,bairro_cli,cidade_Cli,"
    strSQL = strSQL & "uf_cli,cep_cli,"
    strSQL = strSQL & "data_cli,obs_cli,mae_cli,pai_cli,mens_cli,limite_cli,"
    strSQL = strSQL & "status_cli,cod_set,"
    strSQL = strSQL & "localtrab_cli,endtrab_cli,setortrab_cli,teltrab_cli,ceptrab_cli,cobranca_cli,cod_gru,sexo_cli,cod_org,cod_cid,email_cli) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_cli & "',"
    strSQL = strSQL & F_GravaData(datanasc_cli) & ","
    strSQL = strSQL & "'" & F_SoNumero(cpf_cli) & "',"
    strSQL = strSQL & "'" & rg_cli & "',"
    strSQL = strSQL & "'" & Logradouro & "',"
    strSQL = strSQL & "'" & numero & "',"
    strSQL = strSQL & "'" & complemento & "',"
    
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
    strSQL = strSQL & "'" & cod_org & "',"
    strSQL = strSQL & "'" & cod_cid & "',"
    strSQL = strSQL & "'" & email & "')"
    Conn.Execute strSQL
    F_GravaCli = sCod
    
End Function
Function F_AlteraCli(cod_cli As String, nome_cli As String, datanasc_cli As String, cpf_cli As String, referencia_cli As String, _
     data_cli As String, obs_cli As String, status_cli As String, rg_cli As String, mae_cli As String, pai_cli As String, limite_cli As String, _
    endcom_cli As String, usucad_cli As String, tipo_cli As String, credito_cli As String, cod_set As String, cod_gru As String, cobranca_cli As String, localtrab_cli As String, _
    endtrab_cli As String, setortrab_cli As String, teltrab_cli As String, ceptrab_cli As String, cod_org As String, dataalt_cli As String, sexo_cli As String, _
    email_cli As String, creditosuspenso_cli As String, ie_Cli As String, cod_cid As String, _
    telefone_cli As String, Logradouro_cli As String, numero_cli As String, complemento_cli As String, _
    indiedest_cli As String, especial_cli As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro:
    
    strSQL = "update " & TAB_CLIENTES & " set "
    strSQL = strSQL & " nome_cli ='" & nome_cli & "',"
    strSQL = strSQL & "datanasc_cli =" & F_GravaData(datanasc_cli) & ","
    strSQL = strSQL & "cpf_cli ='" & F_SoNumero(cpf_cli) & "',"
    strSQL = strSQL & "cep_cli ='" & cep_cli & "',"
    strSQL = strSQL & "referencia_cli ='" & referencia_cli & "',"
    strSQL = strSQL & "especial_cli ='" & especial_cli & "',"
    strSQL = strSQL & "data_cli =" & F_GravaData(data_cli) & ","
    strSQL = strSQL & "obs_cli ='" & obs_cli & "',"
    strSQL = strSQL & "status_cli ='" & status_cli & "',"
    strSQL = strSQL & "rg_cli ='" & rg_cli & "',"
    strSQL = strSQL & "mae_cli ='" & mae_cli & "',"
    strSQL = strSQL & "pai_cli ='" & pai_cli & "',"
    
    If limite_cli <> "" Then
        strSQL = strSQL & "limite_cli =" & F_GravaDec(limite_cli) & ","
    End If
    
    strSQL = strSQL & "endcom_cli ='" & endcom_cli & "',"
    
    strSQL = strSQL & "usucad_cli ='" & usucad_cli & "',"
    
    strSQL = strSQL & "tipo_cli ='" & tipo_cli & "',"
    If credito_cli <> "" Then
        strSQL = strSQL & "credito_cli =" & F_GravaDec(credito_cli) & ","
    End If
    strSQL = strSQL & "cod_set ='" & cod_set & "',"
    strSQL = strSQL & "cod_gru ='" & cod_gru & "',"
    strSQL = strSQL & "cobranca_cli ='" & cobranca_cli & "',"
    strSQL = strSQL & "localtrab_cli='" & localtrab_cli & "',"
    strSQL = strSQL & "endtrab_cli='" & endtrab_cli & "', "
    strSQL = strSQL & "setortrab_cli='" & setortrab_cli & "',"
    strSQL = strSQL & "teltrab_cli='" & teltrab_cli & "',"
    strSQL = strSQL & "ceptrab_cli='" & ceptrab_cli & "',"
    strSQL = strSQL & "cod_org='" & cod_org & "',"
    strSQL = strSQL & "dataalt_cli=" & F_GravaData(dataalt_cli) & " ,"
    strSQL = strSQL & "sexo_cli='" & sexo_cli & "',"
    strSQL = strSQL & "email_cli='" & email_cli & "',"
    strSQL = strSQL & "creditosuspenso_cli='" & creditosuspenso_cli & "',"
    strSQL = strSQL & "ie_Cli='" & ie_Cli & "',"
    strSQL = strSQL & "cod_cid='" & cod_cid & "',"
    strSQL = strSQL & "telefone_cli='" & telefone_cli & "',"
    strSQL = strSQL & "Logradouro_cli='" & Logradouro_cli & "',"
    strSQL = strSQL & "numero_cli='" & numero_cli & "',"
    strSQL = strSQL & "complemento_cli='" & complemento_cli & "',"
    strSQL = strSQL & "indiedest_cli='" & indiedest_cli & "'   "
    strSQL = strSQL & " where cod_cli='" & cod_cli & "'"

    Conn.Execute strSQL
    F_AlteraCli = True
    Exit Function
trata_Erro:
    F_AlteraCli = False

End Function
Function F_MigraCli(cod_cli As String, nome_cli As String, datanasc_cli As String, cpf_cli As String, rg_cli As String, _
    endereco_cli As String, bairro_cli As String, cidade_cli As String, uf_cli As String, cep_cli As String, _
    obs_cli As String, mae_cli As String, pai_cli As String, mens_cli As String, limite_cli As String, _
    cod_set As String, localtrab_cli As String, endtrab_cli As String, _
    setortrab_cli As String, teltrab_cli As String, ceptrab_cli As String, Optional email As String = "") As String
    Dim strSQL As String
    
    Dim sCod As String
    sCod = cod_cli
    
    
    strSQL = "insert into tab_clientes(cod_cli,nome_cli,datanasc_cli,cpf_cli,rg_cli,"
    strSQL = strSQL & "endereco_Cli,bairro_cli,cidade_Cli,"
    strSQL = strSQL & "uf_cli,cep_cli,"
    strSQL = strSQL & "data_cli,obs_cli,mae_cli,pai_cli,mens_cli,limite_cli,"
    strSQL = strSQL & "status_cli,cod_set,"
    strSQL = strSQL & "localtrab_cli,endtrab_cli,setortrab_cli,teltrab_cli,ceptrab_cli,email_cli) values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_cli & "',"
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
    strSQL = strSQL & "'" & email & "')"
    On Error Resume Next
    Conn.Execute strSQL
    F_MigraCli = sCod
    
End Function

Function F_AlteraStatusCli(cod_cli As String, status_cli As String, Optional creditosuspenso_cli As String = "N") As Boolean
    Dim strSQL As String
    
    If cod_org = "" Then
        cod_org = "00000"
    End If
    
    On Error GoTo trata_Erro
    
    strSQL = "update tab_clientes set "
    
    
     
    strSQL = strSQL & "status_cli='" & status_cli & "' "
    If creditosuspenso_cli <> "" Then
        strSQL = strSQL & ",creditosuspenso_cli='" & creditosuspenso_cli & "'  "
    End If
    strSQL = strSQL & " where cod_cli='" & cod_cli & "'"
    Conn.Execute strSQL
    F_AlteraStatusCli = True
    Exit Function
trata_Erro:
    F_AlteraStatusCli = False
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
    Conn.Execute strSQL
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
            Conn.Execute strSQL
    F_GravaAutorizados = sCod
    
    Exit Function
trata_Erro:
    F_GravaAutorizados = ""
End Function


Function F_BuscaTel(cod_cli As String, tipo_ter As String, Optional strDiv As String = " / ") As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    strSQL = "select * from tab_telefones where cod_cli='" & cod_cli & "' and tipo_ter='" & tipo_ter & "'"
    abreObjetoAdo rPesq, strSQL
    With rPesq
        Do While Not .EOF
            F_BuscaTel = F_BuscaTel + !num_tel
            .MoveNext
            If Not .EOF Then
                F_BuscaTel = F_BuscaTel & strDiv
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
    strSQL = "select sum(ti.reman_itd) as total from tab_itdoc as ti, tab_doc as td "
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
    abreObjetoAdo rPesq, strSQL
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
        F_BuscaGrupoCli.cod_gru = frmSelect.msflxgrd.Text
        frmSelect.msflxgrd.col = 1
        F_BuscaGrupoCli.nome_gru = frmSelect.msflxgrd.Text
    End If
    Unload frmSelect
Exit Function
trata_Erro:
    P_MostraErro
End Function
Function F_SelectItensCompradosVendas(cod_cli As String, Optional dataI As String = "", _
    Optional dataF As String = "") As String
    Dim strSQL As String
    strSQL = "select tfo.nome_for,tc.nome_cli, ti.*,tv.num_ven,tv.data_ven,"
    strSQL = strSQL & "tv.valor_ven,desconto_ven, tf.guerra_fun "
    strSQL = strSQL & " from tab_clientes as tc,"
    strSQL = strSQL & " tab_vendas as tv, tab_funcionarios as tf,tab_formaspag as tfo,  "
    strSQL = strSQL & " tab_itvendas as ti "
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
    
    If dataI <> "" Then
        strSQL = strSQL & " and data_ven >=" & F_GravaDataTempo(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_ven <=" & F_GravaDataTempo(dataF & " 23:59:59")
    End If
    F_SelectItensCompradosVendas = strSQL
End Function
Sub P_ResumoDeb(cod As String, Optional tipo_ter As String = "C")
    P_LocalizaOperacao tipo_ter, frmResumoDeb.cmbxTipoTer
    
    frmResumoDeb.tag = cod
    frmResumoDeb.txtbxTerceiro.Text = cod
    frmResumoDeb.txtbxTerceiro.tag = cod
    frmResumoDeb.Show 1
    
End Sub
Sub P_ClienteSpc(cod_cli As String, Optional optnShow As Integer = 1)
    On Error Resume Next
    frmClienteSpc.tag = cod_cli
    frmClienteSpc.Show optnShow
    
End Sub
Sub P_ParcelaSpc(cod_cli As String, seq_itd As String, valor As String)
    On Error Resume Next
    frmClienteSpc.tag = cod_cli
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
    Conn.Execute strSQL
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
    Conn.Execute strSQL
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
    Conn.Execute strSQL
    Exit Sub
trata_Erro:
    P_MostraErro
    
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
    strSQL = "select ti.* from tab_itdoc as ti, tab_doc as td, tab_clientespc as tc where seq_itd='" & seq_itd & "' "
    strSQL = strSQL & " and tc.cod_cli=td.cod_ter "
    strSQL = strSQL & " and td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and baixa_spc='0000-00-00' "
    abreObjetoAdo rPesq, strSQL
    With rPesq
        If .RecordCount > 0 Then
            MsgBox "O cliente encontra-se negativado, lembre-se de Baixá-lo no SPC", vbExclamation
        End If
        .Close
    End With
    Exit Sub
trata_Erro:
End Sub
Function F_GravaGrupoCliente(nome As String) As Boolean
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    sCod = MaxKeySQL("tab_grupocli", "cod_gru", "", "00#")
    strSQL = "insert into tab_grupocli(cod_gru,nome_gru)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome & "')"
    Conn.Execute strSQL
    Exit Function
trata_Erro:

End Function
Function TotalDebitoCliente(cod_ter As String, tipo_ter As String) As Currency
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    strSQL = "select sum(reman_itd) as total from TAB_DOC AS td inner join tab_itdoc as ti on td.seq_doc=ti.seq_doc where cod_ter='" & cod_ter & "' "
    strSQL = strSQL & " and tipo_ter='" & tipo_ter & "' "
    strSQL = strSQL & " and tipo_doc='R' "
    abreObjetoAdo rPesq, strSQL
    If rPesq.RecordCount > 0 Then
        TotalDebitoCliente = rPesq!total
    End If
    
    rPesq.Close
    
End Function
Sub BuscaClientePorTelefone(telefone As String)


    Dim strSQL As String
    Dim sTel As String
    Dim rPesq As New ADODB.Recordset
    Dim i As Integer
    
    For i = 1 To Len(telefone)
         sTel = sTel & Mid(telefone, i, 1) & "%"
    Next
    
    
    
    
    
    strSQL = "select tc.nome_cli,tc.cpf_cli, num_tel from tab_clientes as tc inner join tab_telefones as tf "
    strSQL = strSQL & " on tc.cod_cli=tf.cod_cli and tipo_ter='C'"
    strSQL = strSQL & " and num_tel like '%" & sTel & "%' "
    
    abreObjetoAdo rPesq, strSQL

    With rPesq
        Do While Not .EOF
            
            .MoveNext
        Loop
        .Close
    End With



End Sub


Sub BloquearTodos(data As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select distinct tc.cod_cli from tab_clientes as tc inner join tab_doc as td on td.cod_ter=tc.cod_cli"
    strSQL = strSQL & " inner join tab_itdoc as ti on td.seq_doc=ti.seq_doc where tipo_ter='C' "
    strSQL = strSQL & " and ti.vencini_itd <=" & F_GravaData(data)
    strSQL = strSQL & " and reman_itd>0 and tipo_doc='R' and status_cli=1"
    
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            If MsgBox("Confirma o bloqueio de " & .RecordCount & " clientes?", vbYesNo + vbQuestion) = vbYes Then
                Do While Not .EOF
                    F_AlteraStatusCli !cod_cli, "2"
                      
                    .MoveNext
                Loop
            End If
        End If
        .Close
    End With
    
    
End Sub
Sub GerarArquivoBloqueados()
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    strSQL = "select * from tab_clientes where status_cli=2 "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        Do While Not .EOF
            .MoveNext
        Loop
        .Close
    End With
    
    
    
    


End Sub
Sub AbreAnotacoesCli(cod_cli As String)
    frmAnotacoesCliente.tag = cod_cli
    frmAnotacoesCliente.Show 1
    
End Sub
Sub AdicionaAnot(cod_cli As String, anotacao As String)
    Dim strSQL As String
    On Error GoTo trata_Erro:
    strSQL = "insert into tab_clientesanot(cod_cli,anotacao,datahora,cod_usu,excluido)values("
    strSQL = strSQL & cod_cli & ","
    strSQL = strSQL & "'" & anotacao & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & "'N') "
    Conn.Execute strSQL
trata_Erro:
End Sub
Sub RemoveAnot(id As String)
    Dim strSQL As String
    On Error GoTo trata_Erro:
    strSQL = "delete from tab_clientesanot "
    strSQL = strSQL & " where id=" & id
    
    Conn.Execute strSQL
trata_Erro:
End Sub

Sub MarcaExcluidoAnot(id As String)
    Dim strSQL As String
    On Error GoTo trata_Erro:
    strSQL = "update tab_clientesanot set excluido='S' "
    strSQL = strSQL & " where id=" & id
    
    Conn.Execute strSQL
trata_Erro:
End Sub

Sub AbrirContasReceberCliente(cod_cli As String)
    On Error GoTo trata_Erro
    Dim fr As New frmEditaMov
    fr.txtbxTerceiro.tag = cod_cli
    fr.txtbxdataFin.Text = ""
    fr.txtbxDataIni.Text = ""
    fr.cmbxTipoTer.ListIndex = 0
    fr.txtbxTerceiro.Text = cod_cli
    fr.Show 1
    Exit Sub
trata_Erro:
    P_MostraErro
End Sub
