Attribute VB_Name = "Representante"
Type representante
    cod_Rep As String
    nome_rep As String
    limite_rep As Currency
    endereco_rep As String
    bairro_rep As String
    cidade_rep As String
    uf_rep As String
    cep_rep As String
    nasc_rep As String
    natural_rep As String
    rg_rep As String
    cpf_rep As String
    sexo_rep As String
    conjunge_rep As String
    pai_rep As String
    mae_rep As String
    endpais_rep As String
    datacad_rep As String
    status_rep As String
    obs_rep As String
    t() As telefone
    telefone_rep As String
    tipocob As String
    set As setor
    ultconf As String
End Type
Function F_GravaRepresentante(nome_rep As String, limite_rep As String, endereco_rep As String, bairro_rep As String, _
    cidade_rep As String, uf_rep As String, cep_rep As String, nasc_rep As String, natural_rep As String, _
    rg_rep As String, cpf_rep As String, sexo_rep As String, conjuge_rep As String, pai_rep As String, _
    mae_rep As String, endpais_rep As String, obs_rep As String, tipocob As String, Optional avalista_nome As String, _
    Optional avalista_cpf As String, Optional avalista_rg As String, Optional avalista_endereco As String, _
    Optional avalista_telefone As String, Optional avalista_anotacoes As String) As String
    
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    If avalista_telefone = "" Then
        avalista_telefone = "0"
    End If
    sCod = MaxKeySQL("tab_representantes", "cod_rep", "", "0000#")
    strSQL = "insert into tab_Representantes(cod_rep,nome_rep,limite_rep,endereco_rep,"
    strSQL = strSQL & "cod_set,cidade_rep,uf_rep,cep_rep,nasc_rep,natural_rep,rg_rep,cpf_rep,sexo_rep,"
    strSQL = strSQL & "conjuge_rep,pai_rep,mae_rep,endpais_rep,obs_rep,cod_tip, avalista_nome,"
    strSQL = strSQL & "avalista_cpf, avalista_rg,avalista_endereco, avalista_telefones, avalista_anotacoes  )values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_rep & "',"
    strSQL = strSQL & F_GravaDec(limite_rep) & ","
    strSQL = strSQL & "'" & endereco_rep & "',"
    strSQL = strSQL & "'" & bairro_rep & "',"
    strSQL = strSQL & "'" & cidade_rep & "',"
    strSQL = strSQL & "'" & uf_rep & "',"
    strSQL = strSQL & "'" & cep_rep & "',"
    strSQL = strSQL & F_GravaData(nasc_rep) & ","
    strSQL = strSQL & "'" & natural_rep & "',"
    strSQL = strSQL & "'" & rg_rep & "',"
    strSQL = strSQL & "'" & cpf_rep & "',"
    strSQL = strSQL & "'" & sexo_rep & "',"
    strSQL = strSQL & "'" & conjuge_rep & "',"
    strSQL = strSQL & "'" & pai_rep & "',"
    strSQL = strSQL & "'" & mae_rep & "',"
    strSQL = strSQL & "'" & endpais_rep & "',"
    strSQL = strSQL & "'" & obs_rep & "',"
    strSQL = strSQL & "'" & tipocob & "',"
    strSQL = strSQL & "'" & avalista_nome & "',"
    strSQL = strSQL & "'" & avalista_cpf & "',"
    strSQL = strSQL & "'" & avalista_rg & "',"
    strSQL = strSQL & "'" & avalista_endereco & "',"
    strSQL = strSQL & "'" & avalista_telefone & "',"
    strSQL = strSQL & "'" & avalista_anotacoes & "')"
    
    
    
    Conn.Execute strSQL
    F_GravaRepresentante = sCod
    Exit Function
trata_Erro:
    P_MostraErro
End Function

Function ExisteRegistroAberto(cod_Rep As String, mostramsg As Boolean) As Boolean
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    Dim sMsg As String
    strSQL = "select count(*) as total from tab_doc as d inner join tab_itdoc as i on d.seq_doc=i.seq_doc "
    strSQL = strSQL & " where d.cod_ter='" & cod_Rep & "' "
    strSQL = strSQL & " and tipo_ter='R' and reman_itd>0 "

    abreObjetoAdo rPesq, strSQL
    If rPesq!total > 0 Then
        ExisteRegistroAberto = True
    End If
    sMsg = "Parcelas em aberto: " & rPesq!total & Chr(13)
    
    strSQL = "select count(*) as total from tab_consignacoes  "
    strSQL = strSQL & " where cod_rep='" & cod_Rep & "' "
    strSQL = strSQL & " and reman_con>0 "

    abreObjetoAdo rPesq, strSQL
    
    sMsg = sMsg & "Produtos em estoque: " & rPesq!total & Chr(13)
    If rPesq!total > 0 Then
        ExisteRegistroAberto = True
    End If
    FechaObjeto rPesq
    
    If mostramsg Then
        MsgBox sMsg, vbInformation
        
    End If

End Function



Function F_MigraRepresentante(cod_Rep As String, nome_rep As String, limite_rep As String, endereco_rep As String, bairro_rep As String, _
    cidade_rep As String, uf_rep As String, cep_rep As String, nasc_rep As String, natural_rep As String, _
    rg_rep As String, cpf_rep As String, sexo_rep As String, conjuge_rep As String, pai_rep As String, _
    mae_rep As String, endpais_rep As String, obs_rep As String, tipocob As String) As String
    
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    sCod = Format(cod_Rep, "0000#")
    strSQL = "insert into tab_Representantes(cod_rep,nome_rep,limite_rep,endereco_rep,"
    strSQL = strSQL & "bairro_rep,cidade_rep,uf_rep,cep_rep,nasc_rep,natural_rep,rg_rep,cpf_rep,sexo_rep,"
    strSQL = strSQL & "conjuge_rep,pai_rep,mae_rep,endpais_rep,obs_rep,cod_tip)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_rep & "',"
    strSQL = strSQL & F_GravaDec(limite_rep) & ","
    strSQL = strSQL & "'" & endereco_rep & "',"
    strSQL = strSQL & "'" & bairro_rep & "',"
    strSQL = strSQL & "'" & cidade_rep & "',"
    strSQL = strSQL & "'" & uf_rep & "',"
    strSQL = strSQL & "'" & cep_rep & "',"
    strSQL = strSQL & F_GravaData(nasc_rep) & ","
    strSQL = strSQL & "'" & natural_rep & "',"
    strSQL = strSQL & "'" & rg_rep & "',"
    strSQL = strSQL & "'" & cpf_rep & "',"
    strSQL = strSQL & "'" & sexo_rep & "',"
    strSQL = strSQL & "'" & conjuge_rep & "',"
    strSQL = strSQL & "'" & pai_rep & "',"
    strSQL = strSQL & "'" & mae_rep & "',"
    strSQL = strSQL & "'" & endpais_rep & "',"
    strSQL = strSQL & "'" & obs_rep & "',"
    strSQL = strSQL & "'" & tipocob & "')"
    Conn.Execute strSQL
    F_MigraRepresentante = sCod
    Exit Function
trata_Erro:
    P_MostraErro
End Function

Sub AlteraRepresentanteBloqueio(cod_Rep As String, bloqimpr As String, bloqdev As String)
    Dim strSQL As String
    strSQL = "update tab_representantes set "
    If bloqimpr <> "" Then
        strSQL = strSQL & " bloqimpr='" & bloqimpr & "',"
    Else
        strSQL = strSQL & " bloqimpr=bloqimpr,"
    End If
    If bloqdev <> "" Then
        strSQL = strSQL & " bloqdev='" & bloqdev & "' "
    Else
        strSQL = strSQL & " bloqdev=bloqdev"
    End If
    strSQL = strSQL & " where cod_rep='" & cod_Rep & "' "
    Conn.Execute strSQL
End Sub


Function F_AlteraRepresentante(cod_Rep, nome_rep As String, limite_rep As String, endereco_rep As String, bairro_rep As String, _
    cidade_rep As String, uf_rep As String, cep_rep As String, nasc_rep As String, natural_rep As String, _
    rg_rep As String, cpf_rep As String, sexo_rep As String, conjuge_rep As String, pai_rep As String, _
    mae_rep As String, endpais_rep As String, obs_rep As String, status_rep As String, tipocob As String, Optional avalista_nome As String, _
    Optional avalista_cpf As String, Optional avalista_rg As String, Optional avalista_endereco As String, _
    Optional avalista_telefone As String, Optional avalista_anotacoes As String, Optional motivo_bloqueio As String = "") As Boolean
    
    Dim strSQL As String
    
    On Error GoTo trata_Erro
    strSQL = "update tab_representantes set "
    strSQL = strSQL & "nome_rep='" & nome_rep & "',"
    strSQL = strSQL & "limite_rep=" & F_GravaDec(limite_rep) & ","
    strSQL = strSQL & "endereco_rep='" & endereco_rep & "',"
    strSQL = strSQL & "cod_set='" & bairro_rep & "',"
    strSQL = strSQL & "cidade_rep='" & cidade_rep & "',"
    strSQL = strSQL & "uf_rep='" & uf_rep & "',"
    strSQL = strSQL & "cep_rep='" & cep_rep & "',"
    strSQL = strSQL & "nasc_rep=" & F_GravaData(nasc_rep) & ","
    strSQL = strSQL & "natural_rep='" & natural_rep & "',"
    strSQL = strSQL & "rg_rep='" & rg_rep & "',"
    strSQL = strSQL & "cpf_rep='" & cpf_rep & "',"
    strSQL = strSQL & "sexo_rep='" & sexo_rep & "',"
    strSQL = strSQL & "conjuge_rep='" & conjuge_rep & "',"
    strSQL = strSQL & "pai_rep='" & pai_rep & "',"
    strSQL = strSQL & "mae_rep='" & mae_rep & "',"
    strSQL = strSQL & "endpais_rep='" & endpais_rep & "',"
    strSQL = strSQL & "obs_rep= '" & obs_rep & "', "
    strSQL = strSQL & "status_rep='" & status_rep & "',"
    strSQL = strSQL & "avalista_nome='" & avalista_nome & "',"
    strSQL = strSQL & "avalista_cpf='" & avalista_cpf & "',"
    strSQL = strSQL & "avalista_rg='" & avalista_rg & "',"
    strSQL = strSQL & "avalista_endereco='" & avalista_endereco & "',"
    strSQL = strSQL & "avalista_telefones='" & avalista_telefone & "',"
    strSQL = strSQL & "avalista_anotacoes='" & avalista_anotacoes & "',"
    If motivo_bloqueio <> "" Then
        strSQL = strSQL & "motivo_bloqueio='" & motivo_bloqueio & "',"
    End If
    strSQL = strSQL & "cod_tip='" & tipocob & "'"
    
    
    strSQL = strSQL & " where cod_rep='" & cod_Rep & "' "
    Conn.Execute strSQL
    F_AlteraRepresentante = True
    Exit Function
trata_Erro:
    P_MostraErro
End Function

Function F_BuscaRep(busca As String) As representante.representante
    On Error Resume Next
    Dim sTel As String
    frmSelectCli.txtbxNOme.tag = "R"
    frmSelectCli.txtbxNOme.Text = busca
    
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    frmSelectCli.Show 1
    If frmSelectCli.tag <> "" Then
        F_BuscaRep.cod_Rep = frmSelectCli.adoR!cod_Rep
        F_BuscaRep.status_rep = frmSelectCli.adoR!status_rep
        F_BuscaRep.nome_rep = frmSelectCli.adoR!nome_rep
        F_BuscaRep.ultconf = frmSelectCli.adoR!ultimaconferencia
        F_BuscaRep.endereco_rep = frmSelectCli.adoR!endereco_rep
        F_BuscaRep.rg_rep = frmSelectCli.adoR!rg_rep
        F_BuscaRep.cpf_rep = frmSelectCli.adoR!cpf_rep
        F_BuscaRep.limite_rep = frmSelectCli.adoR!limite_rep
        F_BuscaRep.bairro_rep = frmSelectCli.adoR!bairro_rep
        F_BuscaRep.cidade_rep = frmSelectCli.adoR!cidade_rep
        F_BuscaRep.cep_rep = frmSelectCli.adoR!cep_rep
        F_BuscaRep.uf_rep = frmSelectCli.adoR!uf_rep
        F_BuscaRep.tipocob = frmSelectCli.adoR!cod_tip
        strSQL = "select * from tab_telefones where cod_cli='" & frmSelectCli.adoR!cod_Rep & "' and tipo_ter='R'"
        abreObjetoAdo rPesq, strSQL
        With rPesq
            ReDim F_BuscaRep.t(.RecordCount) As telefone
            Do While Not .EOF
                
                F_BuscaRep.t(.AbsolutePosition).cod_cli = !cod_cli
                F_BuscaRep.t(.AbsolutePosition).num_tel = !num_tel
                F_BuscaRep.t(.AbsolutePosition).tipo_tel = !tipo_tel
                
                sTel = sTel & !num_tel
                
                .MoveNext
                If Not .EOF Then
                    sTel = sTel & "/"
                End If
            Loop
        End With
        F_BuscaRep.telefone_rep = sTel
        
    End If
    Unload frmSelectCli
End Function
Sub P_HistProdCons(cod_Rep As String, cod_pro As String)
    frmHistProdConsignado.txtbxRepr.tag = cod_Rep
    frmHistProdConsignado.txtbxRepr.Text = cod_Rep
    frmHistProdConsignado.txtbxProduto.tag = cod_pro
    frmHistProdConsignado.txtbxProduto.Text = cod_pro
    frmHistProdConsignado.Show 1
End Sub


Sub P_Acertos(cod_Rep As String, Optional cod_pro As String = "")
    frmAcertos.txtbxRepr.tag = cod_Rep
    frmAcertos.txtbxRepr.Text = cod_Rep
    frmAcertos.txtbxProduto.tag = cod_pro
    frmAcertos.txtbxProduto.Text = cod_pro
    frmAcertos.Show 1
End Sub
Public Function F_BuscaRepCodigo(cod_Rep As String) As representante.representante
    On Error Resume Next
    Dim sTel As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    On Error GoTo trata_Erro
    
    
    strSQL = "select tc.*,cod_rep as cod, nome_cid,uf_cid,ts.cod_cid,cobrador_set,ts.nome_set,limite_rep,nome_rep as nome,cpf_rep as cnpj,endereco_rep as endereco,  "
                strSQL = strSQL & "cpf_rep as cpf from tab_representantes as tc,tab_setores as ts, "
                strSQL = strSQL & " tab_cidades as ti "
                strSQL = strSQL & " where tc.cod_set=ts.cod_set and ti.cod_cid=ts.cod_cid"
                strSQL = strSQL & " and cod_rep='" & cod_Rep & "' "
            
            
    abreObjetoAdo rPesq, strSQL
    
    
    
    
    
    
    If rPesq.RecordCount > 0 Then
        F_BuscaRepCodigo.cod_Rep = rPesq!cod_Rep
        F_BuscaRepCodigo.status_rep = rPesq!status_rep
        F_BuscaRepCodigo.nome_rep = rPesq!nome_rep
        'F_BuscaRepCodigo.credito_rep = rPesq!credito_rep
        F_BuscaRepCodigo.endereco_rep = rPesq!endereco_rep
        F_BuscaRepCodigo.rg_rep = rPesq!rg_rep
        F_BuscaRepCodigo.cpf_rep = rPesq!cpf_rep
        F_BuscaRepCodigo.limite_rep = rPesq!limite_rep
        F_BuscaRepCodigo.bairro_rep = rPesq!bairro_rep
        F_BuscaRepCodigo.cidade_rep = rPesq!cidade_rep
        F_BuscaRepCodigo.cep_rep = rPesq!cep_rep
        F_BuscaRepCodigo.uf_rep = rPesq!uf_rep
        F_BuscaRepCodigo.set.Cidade = rPesq!nome_cid
        F_BuscaRepCodigo.set.cobrador_set = rPesq!cobrador_set
        F_BuscaRepCodigo.set.cod = rPesq!cod_set
        F_BuscaRepCodigo.set.nome = rPesq!nome_set
       ' F_BuscaRepCodigo.set.tipo = rPesq!tipo_set
        F_BuscaRepCodigo.set.UF = rPesq!uf_cid
        
        strSQL = "select * from tab_telefones where cod_rep='" & rPesq!cod_Rep & "' and tipo_ter='C'"
        abreObjetoAdo rPesq, strSQL
        With rPesq
            ReDim F_BuscaRepCodigo.t(.RecordCount) As telefone
            Do While Not .EOF
                
                F_BuscaRepCodigo.t(.AbsolutePosition).cod_cli = !cod_Rep
                F_BuscaRepCodigo.t(.AbsolutePosition).num_tel = !num_tel
                F_BuscaRepCodigo.t(.AbsolutePosition).tipo_tel = !tipo_tel
                
                sTel = sTel & !num_tel
                
                .MoveNext
                If Not .EOF Then
                    sTel = sTel & "/"
                End If
            Loop
            .Close
        End With
        F_BuscaRepCodigo.telefone_rep = sTel
        
    End If
    Exit Function
    P_MostraErro
trata_Erro:
End Function

Sub P_RelResumoAtv(cod_Rep As String, dataI As String, dataF As String)
    Dim sNomeArq As String
    Dim strSQL As String
    Dim sNomeRep As String
    Dim rPesq As New ADODB.Recordset
    Dim iDifD As Integer
    Dim bOk As Boolean
    If dataI = "" Or dataF = "" Then
        MsgBox "Você deve informar as datas corretamente!", vbExclamation
        Exit Sub
    End If
    iDifD = DateDiff("D", dataI, dataF)
    sNomeRep = F_BuscaUmCampo("nome_rep", "tab_representantes", "cod_rep='" & cod_Rep & "'")
    sNomeArq = "ResumoAtvDeb.html"
    sNomeArq = F_AbreHtml(sNomeArq)
    P_HtmlAbreTitlePlus 1, "Resumo de Atividades de Representante"
    
        P_HtmlAbreBody
    
        P_HtmlAbreH 1, 3, "Representante: " & cod_Rep & "-" & sNomeRep & " Período de " & dataI & " Até " & dataF, True
    
    For iCount = 0 To iDifD - 1
        bOk = False
        strSQL = "select nome_sub,nome_sec,tp.pv_pro,tp.pa_pro,tp.nome_pro,tp.ref_pro,tc.*,unit_con*reman_con as total,pv_pro*reman_con as atual"
        strSQL = strSQL & " from tab_produtos as tp, tab_consignacoes as tc,tab_subsec as tu,tab_secoes as ts "
        strSQL = strSQL & " where tc.cod_pro=tp.cod_pro "
        strSQL = strSQL & " and tu.cod_sub=tp.cod_sub "
        strSQL = strSQL & " and tu.cod_sec=ts.cod_sec"
        strSQL = strSQL & " and tc.cod_rep='" & cod_Rep & "' "
        strSQL = strSQL & " and left(data_con,10)=" & F_GravaData(DateAdd("D", iCount, dataI))
        strSQL = strSQL & " order by cod_pro"
        abreObjetoAdo rPesq, strSQL
        With rPesq
            
            If .RecordCount > 0 Then
                    
                
                P_HtmlAbreTable 1, 1, 700
                
                P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, "<b>Transferências em " & Format(!data_con, "dd/mm/yy") & "</b>", "center", True, , "7"
                P_HtmlFechaLinha
                
                P_HtmlAbreLinha
                
                P_HtmlAbreColuna 1, "Hora:"
                P_HtmlAbreColuna 1, "Seq: "
                P_HtmlAbreColuna 1, "Código"
                P_HtmlAbreColuna 1, "Produto"
                P_HtmlAbreColuna 1, "Qtd"
                'P_HtmlAbreColuna 1, "Reman"
                P_HtmlAbreColuna 1, "Usuário"
                P_HtmlFechaLinha
                
                Do While Not .EOF
                    P_HtmlAbreLinha
                
                    P_HtmlAbreColuna 1, Format(!data_con, "hh:MM")
                    P_HtmlAbreColuna 1, !seq_con
                    P_HtmlAbreColuna 1, !cod_pro
                    P_HtmlAbreColuna 1, !nome_pro, "left"
                    P_HtmlAbreColuna 1, !qtd_con
                    'P_HtmlAbreColuna 1, (!reman_con), "Right"
                    P_HtmlAbreColuna 1, !cod_usu
                    P_HtmlFechaLinha
                    
                    bOk = True
                    .MoveNext
                
                    
                Loop
                P_HtmlFechaTable 1
            End If
        .Close
        End With
        
        'devolução
    
    
        strSQL = "select td.*,ti.seq_con,ti.seq_dev,ti.qtd_dev,ti.seq_ite,tp.nome_pro,ti.cod_pro "
        strSQL = strSQL & "from tab_itdevconsign as ti, tab_devconsign as td,"
        strSQL = strSQL & "tab_produtos as tp "
        strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
        strSQL = strSQL & " and td.seq_dev=ti.seq_dev "
        strSQL = strSQL & " and td.cod_rep='" & cod_Rep & "' "
        strSQL = strSQL & " and left(data_dev,10)=" & F_GravaData(DateAdd("D", iCount, dataI))
        strSQL = strSQL & " order by ti.cod_pro"
        abreObjetoAdo rPesq, strSQL
        With rPesq
            
            If .RecordCount > 0 Then
                P_HtmlAbreTable 1, 1, 700
                
                P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, "<b>Devoluções em " & Format(!data_dev, "dd/mm/yy") & "</b>", "center", True, , 6
                P_HtmlFechaLinha
                
                P_HtmlAbreLinha
                P_HtmlAbreColuna 1, "Hora"
                P_HtmlAbreColuna 1, "Seq: "
                P_HtmlAbreColuna 1, "Código"
                P_HtmlAbreColuna 1, "Produto"
                P_HtmlAbreColuna 1, "Qtd"
                P_HtmlAbreColuna 1, "Usuário"
                P_HtmlFechaLinha
                
                Do While Not .EOF
                    P_HtmlAbreLinha
                
                    P_HtmlAbreColuna 1, Format(!data_dev, "hh:MM")
                    P_HtmlAbreColuna 1, !seq_con
                    P_HtmlAbreColuna 1, !cod_pro
                    P_HtmlAbreColuna 1, !nome_pro, "Left"
                    P_HtmlAbreColuna 1, !qtd_dev
                    P_HtmlAbreColuna 1, !cod_usu
                    P_HtmlFechaLinha
                    
                    bOk = True
                    .MoveNext
                
                    
                Loop
                P_HtmlFechaTable 1
            End If
                
            If bOk Then
            '    P_HtmlAbreH 1, 4, "Data: " & DateAdd("D", iCount, dataI), True
            End If
            
        .Close
    
    
        End With
    
        strSQL = "select ta.*,ti.seq_con,ti.*,tp.nome_pro "
        strSQL = strSQL & "from tab_itacertos as ti, tab_acertos as ta,"
        strSQL = strSQL & "tab_produtos as tp "
        strSQL = strSQL & " where tp.cod_pro=ti.cod_pro "
        strSQL = strSQL & " and ta.seq_ace=ti.seq_ace "
        strSQL = strSQL & " and status_ace=2"
        strSQL = strSQL & " and ta.cod_rep='" & cod_Rep & "' "
        strSQL = strSQL & " and (data_ace)=" & F_GravaData(DateAdd("D", iCount, dataI))
        strSQL = strSQL & " order by ti.cod_pro"
        abreObjetoAdo rPesq, strSQL
        With rPesq
            
            If .RecordCount > 0 Then
                P_HtmlAbreTable 1, 1, 700
                
                P_HtmlAbreLinha
                    P_HtmlAbreColuna 1, "<b>Acertos em " & Format(!data_ace, "dd/mm/yy") & "</b>", "center", True, , 6
                P_HtmlFechaLinha
                
                P_HtmlAbreLinha
                P_HtmlAbreColuna 1, "Hora"
                P_HtmlAbreColuna 1, "Seq: "
                P_HtmlAbreColuna 1, "Código"
                P_HtmlAbreColuna 1, "Produto"
                P_HtmlAbreColuna 1, "Qtd"
                P_HtmlAbreColuna 1, "Usuário"
                P_HtmlFechaLinha
                
                Do While Not .EOF
                    P_HtmlAbreLinha
                
                    P_HtmlAbreColuna 1, Format(!hora_ace, "hh:MM")
                    P_HtmlAbreColuna 1, !seq_ace
                    P_HtmlAbreColuna 1, !cod_pro
                    P_HtmlAbreColuna 1, !nome_pro, "Left"
                    P_HtmlAbreColuna 1, !qtd_ite
                    P_HtmlAbreColuna 1, !cod_usu
                    P_HtmlFechaLinha
                    
                    bOk = True
                    .MoveNext
                
                    
                Loop
                P_HtmlFechaTable 1
            End If
                
            If bOk Then
            '    P_HtmlAbreH 1, 4, "Data: " & DateAdd("D", iCount, dataI), True
            End If
            
        .Close
    
    
        End With
    
    
    
    
    
    
    
    Next
    P_HtmlFechaBody
    P_FechaHtml 1, True, sNomeArq
End Sub
Sub P_AlteraSenhaRep(cod_Rep As String, senha As String)
    Dim strSQL As String
    strSQL = "update tab_representantes set senha_rep='" & senha & "' "
    strSQL = strSQL & " where cod_rep='" & cod_Rep & "'"
    Conn.Execute strSQL
    
End Sub
Function F_SenhaRep(cod_Rep As String) As Boolean
    frmSenhaRepr.tag = cod_Rep
    frmSenhaRepr.txtbxSenha.Text = ""
    frmSenhaRepr.Show 1
    If frmSenhaRepr.tag = "x" Then
        F_SenhaRep = True
    Else
        F_SenhaRep = False
    End If
End Function
Function F_SelectProdutosPorRepresentante(Optional cod_pro As String = "", Optional cod_Rep As String = "", _
    Optional cod_sec As String = "", Optional cod_sub, Optional sOrderBy = "order by tp.cod_pro,tr.nome_rep") As String
    Dim strSQL As String
    strSQL = "select ts.nome_sub, tse.nome_sec,te.estoque_pro,"
    strSQL = strSQL & " tp.*,tr.nome_rep,tr.cod_rep"
    strSQL = strSQL & " from tab_produtos as tp,tab_subsec as ts, "
    strSQL = strSQL & " tab_secoes as tse,tab_estcons as te , tab_representantes as tr "
    strSQL = strSQL & " where tse.cod_sec=ts.cod_sec "
    strSQL = strSQL & " and ts.cod_sub=tp.cod_sub "
    strSQL = strSQL & " and tp.cod_pro=te.cod_pro "
    strSQL = strSQL & " and te.cod_rep=tr.cod_rep "
    strSQL = strSQL & " and estoque_pro>0"
    If cod_pro <> "" Then
        strSQL = strSQL & " and tp.cod_pro='" & cod_pro & "' "
    End If
    If cod_Rep <> "" Then
        strSQL = strSQL & " and tr.cod_rep='" & cod_Rep & "' "
    End If
    If cod_sub <> "" Then
        strSQL = strSQL & " and tp.cod_sub='" & cod_sub & "' "
    End If
    If cod_sec <> "" Then
        strSQL = strSQL & " and ts.cod_sec='" & cod_sec & "' "
    End If
    
    strSQL = strSQL & sOrderBy
    
    F_SelectProdutosPorRepresentante = strSQL
    
End Function
Sub P_RelParcelasPorRepr(cod_Rep As String, Optional dataI As String = "", Optional dataF As String = "")
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    'On Error Resume Next
    strSQL = doc.F_SelectDocsReceberPorReprDeAcerto(dataI, dataF, cod_Rep)
    abreObjetoAdo rPesq, strSQL
    P_rPesqToHtml "Relatório de Docs a Receber de Acertos por Representante", rPesq, dataI & " :.: " & dataF & " :.: " & cod_Rep & " - " & F_BuscaUmCampo("nome_rep", "tab_representantes", "cod_rep='" & cod_Rep & "'")
    
    

End Sub
Sub buscaLembrete(cod_Rep As String, Optional abreEmBranco = False)

    If F_BuscaUmCampo("lembrete", "tab_representantes", "cod_rep='" & cod_Rep & "'") <> "" Or abreEmBranco = True Then
        frmLembrete.tag = cod_Rep
        frmLembrete.id = "cod_rep"
        frmLembrete.campo = "lembrete"
        frmLembrete.tabela = "tab_representantes"
        frmLembrete.Show 1
    End If

End Sub

Sub buscaLembreteAcerto(cod_Rep As String, Optional abreEmBranco = False)

    If F_BuscaUmCampo("lembreteAcerto", "tab_representantes", "cod_rep='" & cod_Rep & "'") <> "" Or abreEmBranco = True Then
        frmLembrete.tag = cod_Rep
        frmLembrete.id = "cod_rep"
        frmLembrete.campo = "lembreteAcerto"
        frmLembrete.tabela = "tab_representantes"
        frmLembrete.Show 1
    End If

End Sub
Function RepresentanteBloqueado(cod_Rep As String) As Boolean

    If F_BuscaUmCampo("bloqimpr", "tab_representantes", "cod_rep='" & cod_Rep & "'") = "S" Then
        RepresentanteBloqueado = True
    End If
End Function

