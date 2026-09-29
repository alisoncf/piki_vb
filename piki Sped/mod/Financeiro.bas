Attribute VB_Name = "Financeiro"
Type cheque1
    num_che As String
    cod_bco As String
    cco_che As String
    age_che As String
    nome_che As String
    cnpj_che As String
    Venc_che As String
    Valor_che As Currency
    cpf_che As String
End Type
Type chePag1 '
    num_che As String
    cod_con As String
    Venc_che As String
    Valor_che As Currency
End Type
Type parcela
    vencimento As String
    valor As Currency
    boleto As String
End Type
Type localche
    cod_loc As String
    nome_loc As String
End Type
Type ContaCorrente
    cod_con As String
    nome_con As String

End Type
Type financeiro
    seq_fin As String
    cod_con As String
    data_fin As String
    historico_fin As String
    cod_usu As String
    saldo_fin As Currency
    valor_fin As Currency
    vinculado_fin As String
    
End Type

Type PlanoContas
    cod_pla As String
    tipo_pla As String
    nome_pla As String
End Type


Function F_SelectContasPagar(Optional dataI As String = "", Optional dataF As String = "", _
    Optional num_doc As String = "", Optional seq_doc As String = "", Optional seq_itd As String = "", _
    Optional cod_For As String = "", Optional cod_pla As String = "", Optional historico_doc As String = "", _
    Optional emaberto As Boolean = True, Optional sOrderBy As String = "order by venc_itd,nome_for") As String
    Dim strSQL As String
    
    strSQL = "select tf.nome_for,tf.cnpj_for,td.*,ti.*,tp.nome_pla,concat(td.parc_doc,'/',ti.parc_itd) as parc,"
    strSQL = strSQL & " tf.nome_for as nome,tf.cnpj_for as cnpj "
    strSQL = strSQL & " from  " & TAB_DOC & " 1 as td, " & TAB_PLANOCONTAS & " as tp, " & TAB_FORNECEDORES & " as tf, "
    strSQL = strSQL & "  " & TAB_ITDOC & " 1 as ti "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.cod_pla=tp.cod_pla "
    strSQL = strSQL & " and td.cod_ter=tf.cod_for "
    strSQL = strSQL & " and td.tipo_doc='P' "
    If dataI <> "" Then
        strSQL = strSQL & " and venc_itd>=" & F_GravaData(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and venc_itd<=" & F_GravaData(dataF)
    End If
    If num_doc <> "" Then
        strSQL = strSQL & " and num_doc='" & num_doc & "' "
    End If
    If seq_doc <> "" Then
        strSQL = strSQL & " and td.seq_doc='" & seq_doc & "' "
    End If
    If seq_itd <> "" Then
        strSQL = strSQL & " and ti.seq_itd='" & seq_itd & "' "
    End If
    If cod_For <> "" Then
        strSQL = strSQL & " and tf.cod_for='" & cod_For & "' "
    End If
    If cod_pla <> "" Then
        strSQL = strSQL & " and td.cod_pla='" & cod_pla & "' "
    End If
    If historico_doc <> "" Then
        strSQL = strSQL & " and td.historico_doc='" & historico_doc & "' "
    End If
    If emaberto Then
        strSQL = strSQL & " and ti.reman_itd>0"
    End If
    If sOrderBy <> "" Then
        strSQL = strSQL & sOrderBy
    End If
    F_SelectContasPagar = strSQL
End Function

Function F_SelectContasReceber(Optional dataI As String = "", Optional dataF As String = "", _
    Optional num_doc As String = "", Optional seq_doc As String = "", Optional seq_itd As String = "", _
    Optional cod_cli As String = "", Optional cod_pla As String = "", Optional historico_doc As String = "", _
    Optional emaberto As Boolean = True, Optional sOrderBy As String = "order by venc_itd,nome_cli") As String
    Dim strSQL As String
    
    strSQL = "select tc.nome_cli,tc.cpf_cli,td.*,ti.*,tp.nome_pla,concat(td.parcelas_doc,'/',ti.parcela_itd) as parc,"
    strSQL = strSQL & " tc.nome_cli as nome,tc.cpf_cli as cpf  "
    strSQL = strSQL & " from  " & TAB_DOC & "  as td, " & TAB_PLANOCONTAS & " as tp, " & TAB_CLIENTES & " as tc, "
    strSQL = strSQL & "  " & TAB_ITDOC & "  as ti "
    strSQL = strSQL & " where td.seq_doc=ti.seq_doc "
    strSQL = strSQL & " and td.cod_pla=tp.cod_pla "
    strSQL = strSQL & " and td.cod_ter=tc.cod_cli "
    strSQL = strSQL & " and td.tipo_doc='R' "
    If dataI <> "" Then
        strSQL = strSQL & " and venc_itd>=" & F_GravaData(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and venc_itd<=" & F_GravaData(dataF)
    End If
    If num_doc <> "" Then
        strSQL = strSQL & " and num_doc='" & num_doc & "' "
    End If
    If seq_doc <> "" Then
        strSQL = strSQL & " and td.seq_doc='" & seq_doc & "' "
    End If
    If seq_itd <> "" Then
        strSQL = strSQL & " and ti.seq_itd='" & seq_itd & "' "
    End If
    If cod_cli <> "" Then
        strSQL = strSQL & " and tc.cod_cli='" & cod_cli & "' "
    End If
    If cod_pla <> "" Then
        strSQL = strSQL & " and td.cod_pla='" & cod_pla & "' "
    End If
    If historico_doc <> "" Then
        strSQL = strSQL & " and td.historico_doc='" & historico_doc & "' "
    End If
    If emaberto Then
        strSQL = strSQL & " and ti.reman_itd>0"
    End If
    If sOrderBy <> "" Then
        strSQL = strSQL & sOrderBy
    End If
    F_SelectContasReceber = strSQL
End Function


Function f_BuscaLocalChe(expr As String) As localche
    frmSelect.sSelect = "select cod_loc as cod,nome_loc as nome from " & TAB_LOCALCHE
    frmSelect.cCriterio.Add " where nome_loc like '" & expr & "%'"
    frmSelect.sOrd = " order by nome_loc "
    frmSelect.txtbxExpr.text = expr
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        
        f_BuscaLocalChe.cod_loc = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.col = 1
        
        f_BuscaLocalChe.nome_loc = frmSelect.msflxgrd
        
        
        
    End If
    Unload frmSelect
End Function
Function F_BuscaContaCorrente(expr As String) As ContaCorrente
    On Error Resume Next
    frmEdiCon.tag = "x"
    frmEdiCon.txtBusca.text = expr
    frmEdiCon.Show 1

    If frmEdiCon.tag <> "" Then
        F_BuscaContaCorrente.cod_con = frmEdiCon.ado!cod_con
        F_BuscaContaCorrente.nome_con = frmEdiCon.ado!nome_con
    End If
    Unload frmEdiCon
    
End Function


Function F_GravaFinanceiro(tipo_fin As String, data_fin As String, _
    valor_fin As String, historico_fin As String, _
    cod_con As String, vinculado_fin As String, cod_pla As String, Optional seq_fin As String = "", Optional seq_tra As String) As String
    
    Dim strSQL As String
    Dim sSeq  As String
    On Error GoTo trata_Erro
    'operacao_fin F=Financeiro, A=Acerto
    If cod_pla = "000" Or cod_pla = "" Then
        If tipo_fin = "C" Then
            cod_pla = "10001"
        Else
            cod_pla = "20001"
        End If
    End If
    If tipo_fin = "R" Or tipo_cai = "+" Then
        tipo_fin = "C"
    End If
    If tipo_fin = "P" Or tipo_cai = "-" Then
        tipo_fin = "D"
    End If
    strSQL = "insert into " & TAB_FINANCEIRO & "(seq_fin,data_fin,valor_fin,saldo_fin,historico_fin,tipo_fin,"
    strSQL = strSQL & "cod_con,vinculado_fin,cod_pla,seq_tra,cod_usu)Values("
    If seq_fin = "" Then
        sSeq = ChaveSQL(TAB_FINANCEIRO, "seq_fin", "000000#")
    Else
        sSeq = seq_fin
    End If
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(data_fin) & ","
    strSQL = strSQL & F_GravaDec(valor_fin) & ","
    If tipo_fin = "C" Then
        strSQL = strSQL & " saldo_fIN + " & F_GravaDec(valor_fin) & ","
    Else
        strSQL = strSQL & " saldo_FIN + " & F_GravaDec(valor_fin) & ","
    End If
    strSQL = strSQL & "'" & Left(historico_fin, 100) & "',"
    strSQL = strSQL & "'" & tipo_fin & "', "
    
    
    strSQL = strSQL & "'" & cod_con & "',"
    strSQL = strSQL & "'" & vinculado_fin & "',"
    strSQL = strSQL & "'" & cod_pla & "',"
    strSQL = strSQL & "'" & seq_tra & "', "
    strSQL = strSQL & "'" & strCod_Usu & "') "
    conn.Execute strSQL
    F_GravaFinanceiro = sSeq
    
trata_Erro:
    
    
End Function
Function F_AlteraFin(seq_fin As String, historico_fin As String, data_fin As String, _
    valor_fin As String, cod_con As String, cod_pla As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update " & TAB_FINANCEIRO & " set "
    strSQL = strSQL & " historico_Fin='" & historico_fin & "',"
    strSQL = strSQL & " cod_pla='" & cod_pla & "',"
    strSQL = strSQL & " data_fin=" & F_GravaData(data_fin) & ","
    strSQL = strSQL & "valor_fin=" & F_GravaDec(valor_fin) & ","
    
    strSQL = strSQL & "cod_con='" & cod_con & "'  "
    strSQL = strSQL & " where seq_fin='" & seq_fin & "' "
    conn.Execute strSQL
    F_AlteraFin = True
Exit Function
trata_Erro:
    F_AlteraFin = False
End Function
Sub P_Class(valor_cla As Currency, sTipo As String, seq As String, operacao As String, historico As String)
    'On Error Resume Next
    'operação F = lançamento financeiro, D = documento
    frmClass.cmndbtnConfirma.tag = sTipo
    frmClass.lblTotal.caption = FormatVal(valor_cla)
    frmClass.tag = seq
    frmClass.operacao = operacao
    frmClass.lblHistorico.caption = historico
    frmClass.Show 1

    If frmClass.tag <> "" Then
'        MsgBox "Lançamento Classificado!", vbInformation

    Else
        MsgBox "O Lançamento não foi classificado!", vbExclamation
    End If
    
    Unload frmClass
End Sub

Sub P_PreencheDepto(obj As Object)

End Sub


Function F_GravaPagFun(seq_fin As String, cod_fun As String, cod_tip As String) As Boolean
    Dim strSQL As String
    strSQL = "delete from tab_pagfun where seq_fin='" & seq_fin & "'"
    conn.Execute strSQL
    strSQL = "insert into tab_pagfun(seq_fin,cod_fun)values("
    strSQL = strSQL & "'" & seq_fin & "',"
    strSQL = strSQL & "'" & cod_fun & "',"
    strSQL = strSQL & "'" & cod_tip & "') "
    conn.Execute strSQL
End Function
    
Function F_BuscaPlanoContas(tipo_pla As String, Optional expr As String = "") As PlanoContas
    Dim strSQL As String
    If tipo_pla = "P" Then
        tipo_pla = "D"
    ElseIf tipo_pla = "R" Then
        tipo_pla = "C"
    End If
    frmEdiPlanoContas.tag = "x"
    frmEdiPlanoContas.txtDesc.text = expr
    If tipo_pla = "C" Then
        frmEdiPlanoContas.optnCred.Value = True
        
    Else
        frmEdiPlanoContas.optnDeb.Value = True
    End If
    frmEdiPlanoContas.optnDeb.Enabled = False
    frmEdiPlanoContas.optnCred.Enabled = False
    
    frmEdiPlanoContas.Show 1
    With frmEdiPlanoContas
        If .tag <> "" Then
            F_BuscaPlanoContas.cod_pla = frmEdiPlanoContas.ado!cod_pla
            F_BuscaPlanoContas.nome_pla = frmEdiPlanoContas.ado!nome_pla
            F_BuscaPlanoContas.tipo_pla = frmEdiPlanoContas.ado!tipo_pla
            
        End If
        
    End With
    Unload frmEdiPlanoContas
    
End Function

Function F_BuscaChaveBaixa() As String
    F_BuscaChaveBaixa = ChaveSQL("tab_baixait", "seq_cab", "000000#")
End Function
Function F_GravaValesCai(seq_cai As String, cod_fun As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into tab_valescai(seq_cai,cod_fun)values("
    strSQL = strSQL & "'" & seq_cai & "',"
    strSQL = strSQL & "'" & cod_fun & "')"
    conn.Execute strSQL
    F_GravaValesCai = True
    Exit Function
trata_Erro:
    
    
End Function

Function F_SelectContaCorrente() As ContaCorrente
    frmContaCorrente.Show 1
    If frmContaCorrente.tag <> "" Then
        F_SelectContaCorrente.cod_con = Left(frmContaCorrente.lstConta.text, 3)
        F_SelectContaCorrente.nome_con = Mid(frmContaCorrente.lstConta.text, 4, Len(frmContaCorrente.lstConta.text))
        
    End If
    Unload frmContaCorrente
End Function
Sub P_ImprimeBaixaCarnesPorCaixa(cod_cai As String, data_ini As String, data_fin As String, cod_usu As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select "
    
    strSQL = strSQL & " tb.cod_cob as Cod,"
    strSQL = strSQL & " tc.nome_cob as Cobranca,"
    strSQL = strSQL & " sum(tx.valor_cai) as total from "
    
    strSQL = strSQL & TAB_BAIXAITD & " As tb, " & TAB_BAIXAFIN & " As tf, tab_Contacob As tc, tab_caixa As tx"
    
    strSQL = strSQL & " where"
    strSQL = strSQL & " tx.seq_cai = tf.seq_fin"
    strSQL = strSQL & " and tc.cod_cob=tb.cod_cob"
    
    
    strSQL = strSQL & " and data_bai<=" & F_GravaData(data_fin)
    strSQL = strSQL & " and data_bai>=" & F_GravaData(data_ini)
    strSQL = strSQL & " and tf.seq_cab=tb.seq_cab"
    If cod_cai <> "" Then
        strSQL = strSQL & " and cod_cai='" & cod_cai & "' "
    End If
    If cod_usu <> "" Then
        strSQL = strSQL & " and tb.cod_usu='" & cod_usu & "' "
    End If
    strSQL = strSQL & " group by tb.cod_cob"
    
    
    rPesq.Open strSQL, conn
    
    P_rPesqToHtml "Relatório de Baixas por Caixa", rPesq, "Caixa" & strCod_Cai, "De: " & data_ini & " Até: " & data_fin, _
     "2"
    
End Sub
Function F_SelectBaixaCarnesPorCaixa(cod_cai As String, data_ini As String, data_fin As String, cod_usu As String) As String
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select "
    
    strSQL = strSQL & " tb.cod_cob as Cod,"
    strSQL = strSQL & " tc.nome_cob as Cobranca,"
    strSQL = strSQL & " sum(tx.valor_cai) as total from "
    
    strSQL = strSQL & TAB_BAIXAITD & " As tb, " & TAB_BAIXAFIN & " As tf, tab_Contacob As tc, tab_caixa As tx"
    
    strSQL = strSQL & " where"
    strSQL = strSQL & " if(tf.esp_bai='FIN',tx.num_ven=tf.seq_fin,tx.seq_cai=tf.seq_fin)"
    strSQL = strSQL & " and tc.cod_cob=tb.cod_cob"
    
    
    strSQL = strSQL & " and data_bai<=" & F_GravaData(data_fin)
    strSQL = strSQL & " and data_bai>=" & F_GravaData(data_ini)
    strSQL = strSQL & " and tf.seq_cab=tb.seq_cab"
    If cod_cai <> "" Then
        strSQL = strSQL & " and cod_cai='" & Left(cod_cai, 2) & "' "
    End If
    If cod_usu <> "" Then
        strSQL = strSQL & " and tb.cod_usu='" & Left(cod_usu, 3) & "' "
    End If
    strSQL = strSQL & " group by tb.cod_cob"
    
    F_SelectBaixaCarnesPorCaixa = strSQL
    
End Function

Sub P_ImprimeValeCai(seq_cai As String)
    Dim strSQL As String
    Dim sNomeArq As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select tc.*,tf.nome_fun from tab_caixa as tc, tab_funcionarios as tf, tab_valescai as tv "
    strSQL = strSQL & " where tv.seq_cai=tc.seq_cai "
    strSQL = strSQL & " and tv.cod_fun=tf.cod_fun "
    strSQL = strSQL & " and tv.seq_cai='" & seq_cai & "' "
    rPesq.Open strSQL, conn
    
    With rPesq
        If .RecordCount > 0 Then
            sNomeArq = html.F_AbreHtml("vale" & seq_cai)
            P_HtmlAbreHead 1
            P_HtmlAbreTitle 1, "Vale " & seq_cai, True
            P_HtmlFechaHead
            
            P_HtmlAbreBody 1
            
            
            P_HtmlAbreH 1, 3, "VALE/ADIANTAMENTO DE FUNCIONÁRIO", True
            
            P_HtmlAbreTable 1, 1, 600
            P_HtmlAbreLinha 1
            P_HtmlAbreColuna 1, "", , False
            P_HtmlAbreP "Data: " & !data_cai, 10
            'P_HtmlBR
            P_HtmlAbreP "Histórico: " & !historico_cai, 10
            'P_HtmlBR
            P_HtmlAbreP "Valor: " & FormatVal(!valor_Cai), 10
            P_HtmlAbreP "Usuário: " & !cod_usu, 10
            
            P_HtmlFechaColuna
            P_HtmlFechaLinha
            P_HtmlFechaTable
            
            
            
            P_HtmlAbreH 1, 3, !nome_fun, True
            
            
            
            P_HtmlFechaBody 1
            P_FechaHtml 1, True, sNomeArq
            
            
            
        Else
            .Close
            MsgBox "Não encontrado!", vbExclamation
            
            Exit Sub
        End If
    End With
    
End Sub

Function F_SelectFinanceiro(Optional dataI As String = "", Optional dataF As String = "", _
    Optional cod_con As String = "", Optional cod_pla As String = "", Optional historico As String = "", _
    Optional valor As String = "", Optional tipo As String = "", Optional order As String = "") As String
    
    If order = "" Then
        order = "tf.seq_fin"
    End If
    
    strSQL = "select   tf.*,if(tipo_fin='C',valor_fin,0) as cred,fun.guerra_fun as usuario,if(tipo_fin='D',valor_fin,0) as deb, tp.nome_pla, tc.nome_con"
    strSQL = strSQL & " From   " & TAB_FINANCEIRO & " as tf left join " & TAB_PLANOCONTAS & " as tp "
    strSQL = strSQL & "  on tf.cod_pla=tp.cod_pla"
    
    strSQL = strSQL & " left join tab_contas as tc"
    
    strSQL = strSQL & " on tc.cod_con=tf.cod_con"
    strSQL = strSQL & " left join " & TAB_FUNCIONARIOS & " as fun on tf.cod_usu=fun.cod_fun "
    strSQL = strSQL & " where valor_fin>0 "
    If tipo <> "" Then
        strSQL = strSQL & " and tipo_fin='" & tipo & "' "
    End If
    If dataI <> "" Then
        strSQL = strSQL & " and data_fin>=" & F_PesqData(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and data_fin<=" & F_PesqData(dataF)
    End If

    If cod_con <> "" Then
        strSQL = strSQL & " and tf.cod_con='" & cod_con & "' "
    End If
    
    If cod_pla <> "" Then
        strSQL = strSQL & " and tf.cod_pla='" & cod_pla & "' "
    End If
    
    If historico <> "" Then
        strSQL = strSQL & " and historico_fin like '%" & historico & "%' "
    End If
    If valor <> "" Then
        strSQL = strSQL & " and valor_fin=" & F_GravaDec(valor)
    End If
    
    strSQL = strSQL & " order by " & order

    F_SelectFinanceiro = strSQL
End Function



Function F_GravaVinc(seq_itd As String, data_vin As String, valor_vin As String, acres_vin As String, _
    decres_vin As String, venc_vin As String, parc_vin As String, seq_doc As String, Optional seq_vin As String = "") As String
    Dim strSQL As String
    Dim sSeq_vin As String
    On Error GoTo trata_Erro:
    If seq_vin = "" Then
        sSeq_vin = MaxKeySQL("tab_vinc", "seq_vin", "", "000000#")
    Else
        sSeq_vin = seq_vin
    End If
    
    strSQL = "insert into tab_vinc(seq_vin,seq_itd,data_vin,"
    strSQL = strSQL & "valor_vin,acres_vin,decres_vin,venc_vin,"
    strSQL = strSQL & "parc_vin,seq_doc)values("
    strSQL = strSQL & "'" & sSeq_vin & "',"
    strSQL = strSQL & "'" & seq_itd & "',"
    strSQL = strSQL & F_GravaData(data_vin) & ","
    strSQL = strSQL & F_GravaDec(valor_vin) & ","
    strSQL = strSQL & F_GravaDec(acres_vin) & ","
    strSQL = strSQL & F_GravaDec(decres_vin) & ","
    strSQL = strSQL & F_GravaData(venc_vin) & ","
    strSQL = strSQL & parc_vin & ","
    strSQL = strSQL & "'" & seq_doc & "')"
    conn.Execute strSQL

    F_GravaVinc = sSeq_vin
Exit Function
    
trata_Erro:
MsgBox "erro"

End Function
