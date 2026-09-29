Attribute VB_Name = "Financeiro"
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
Type ContaCorrente
    cod_con As String
    nome_con As String

End Type
Type PlanoContas
    cod_pla As String
    tipo_pla As String
    nome_pla As String
End Type
Type localche
    cod_loc As String
    nome_loc As String
End Type
Function F_GravaFinanceiro(tipo_fin As String, data_fin As String, _
    valor_fin As String, historico_fin As String, _
    cod_con As String, vinculado_fin As String, cod_pla As String) As String
    
    Dim ret As New ADODB.Recordset
    Dim strSQL As String
    Dim sSeq  As String
    'operacao_fin F=Financeiro, A=Acerto
    If tipo_fin = "R" Or tipo_cai = "+" Then
        tipo_fin = "C"
    End If
    If tipo_fin = "P" Or tipo_cai = "-" Then
        tipo_fin = "D"
    End If
    strSQL = "insert into " & TAB_FINANCEIRO & "(seq_fin,data_fin,valor_fin,saldo_fin,historico_fin,tipo_fin,"
    strSQL = strSQL & "cod_con,vinculado_fin,cod_pla)Values("
    sSeq = ChaveSQL("tab_financeiro", "seq_fin", "000000#")
    strSQL = strSQL & "'" & sSeq & "',"
    strSQL = strSQL & F_GravaData(data_fin) & ","
    strSQL = strSQL & F_GravaDec(valor_fin) & ","
    If tipo_fin = "C" Then
        strSQL = strSQL & " saldo_fIN + " & F_GravaDec(valor_fin) & ","
    Else
        strSQL = strSQL & " saldo_FIN + " & F_GravaDec(valor_fin) & ","
    End If
    strSQL = strSQL & "'" & historico_fin & "',"
    strSQL = strSQL & "'" & tipo_fin & "', "
    
    
    strSQL = strSQL & "'" & cod_con & "',"
    strSQL = strSQL & "'" & vinculado_fin & "',"
    strSQL = strSQL & "'" & cod_pla & "') "
    
    F_GravaFinanceiro = sSeq
    Set ret = Conn.Execute(strSQL)
    F_GravaFinanceiro = sSeq
End Function

Function F_AlteraFin(seq_fin As String, historico_fin As String, data_fin As String, _
    valor_fin As String, cod_con As String, cod_pla As String) As Boolean
    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update tab_financeiro set "
    strSQL = strSQL & " historico_Fin='" & historico_fin & "',"
    strSQL = strSQL & " cod_pla='" & cod_pla & "',"
    strSQL = strSQL & " data_fin=" & F_GravaData(data_fin) & ","
    strSQL = strSQL & "valor_fin=" & F_GravaDec(valor_fin) & ","
    
    strSQL = strSQL & "cod_con='" & cod_con & "'  "
    strSQL = strSQL & " where seq_fin='" & seq_fin & "' "
    Conn.Execute strSQL
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
Function F_BuscaContaCorrente(expr As String) As ContaCorrente
    frmSelect.sSelect = "select cod_con as cod,nome_con as nome from tab_contas "
    frmSelect.cCriterio.Add " where nome_con like '" & expr & "%'"
    frmSelect.sOrd = " order by nome_con "
    frmSelect.txtbxExpr.text = expr
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.col = 0
        
        F_BuscaContaCorrente.cod_con = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.col = 1
        
        F_BuscaContaCorrente.nome_con = frmSelect.msflxgrd
        
        
        
    End If
    Unload frmSelect
    
End Function

Function F_GravaPagFun(seq_fin As String, cod_fun As String, cod_tip As String) As Boolean
    Dim strSQL As String
    strSQL = "delete from tab_pagfun where seq_fin='" & seq_fin & "'"
    Conn.Execute strSQL
    strSQL = "insert into tab_pagfun(seq_fin,cod_fun)values("
    strSQL = strSQL & "'" & seq_fin & "',"
    strSQL = strSQL & "'" & cod_fun & "',"
    strSQL = strSQL & "'" & cod_tip & "') "
    Conn.Execute strSQL
End Function
    
Function F_BuscaPlanoContas(tipo_pla As String, Optional expr As String = "") As PlanoContas
    Dim strSQL As String
    If tipo_pla = "P" Then
        tipo_pla = "D"
    ElseIf tipo_pla = "R" Then
        tipo_pla = "C"
    End If
    
    If expr = "NÃO CLASSIFICADO" Then
        expr = ""
    End If
    
    
    If tipo_pla = "C" Then
        frmEdiPlanoContas.optnCred.Value = True
    Else
        frmEdiPlanoContas.optnCred.Value = False
    End If
    frmEdiPlanoContas.optnCred.Enabled = False
    frmEdiPlanoContas.optnDeb.Enabled = False
    frmEdiPlanoContas.txtbxExpr.text = expr
    frmEdiPlanoContas.tipo_pla = tipo_pla
    frmEdiPlanoContas.tag = "x"
    frmEdiPlanoContas.Show 1
    If frmEdiPlanoContas.tag <> "" Then
        F_BuscaPlanoContas.cod_pla = frmEdiPlanoContas.ado!cod_pla
        F_BuscaPlanoContas.nome_pla = frmEdiPlanoContas.ado!nome_pla
        F_BuscaPlanoContas.tipo_pla = frmEdiPlanoContas.ado!tipo_pla
    End If
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
    Conn.Execute strSQL
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
Function f_BuscaLocalChe(expr As String) As localche
    frmSelect.sSelect = "select cod_loc as cod,nome_loc as nome from tab_localche "
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
Sub P_ImprimeBaixaCarnesPorCaixa(cod_cai As String, data_ini As String, data_fin As String, cod_usu As String)
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "select "
    
    strSQL = strSQL & " tb.cod_cob as Cod,"
    strSQL = strSQL & " tc.nome_cob as Cobranca,"
    strSQL = strSQL & " sum(tx.valor_cai) as total from "
    
    strSQL = strSQL & " tab_baixaitd2 As tb, tab_baixafin As tf, tab_Contacob As tc, tab_caixa As tx"
    
    strSQL = strSQL & " where"
    strSQL = strSQL & " tx.seq_cai = tf.seq_fin"
    strSQL = strSQL & " and tc.cod_cob=tb.cod_cob"
    strSQL = strSQL & " and excluido_cai=false "
    
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
    
    
    abreObjetoAdo rPesq, strSQL
    
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
    
    strSQL = strSQL & " tab_baixaitd2 As tb, tab_baixafin As tf, tab_Contacob As tc, tab_caixa As tx"
    
    strSQL = strSQL & " where"
    strSQL = strSQL & " if(tf.esp_bai='FIN',tx.num_ven=tf.seq_fin,tx.seq_cai=tf.seq_fin)"
    strSQL = strSQL & " and tc.cod_cob=tb.cod_cob"
    strSQL = strSQL & " and excluido_cai=false "
    
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
    strSQL = strSQL & " and excluido_cai=false "
    abreObjetoAdo rPesq, strSQL
    
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
    Optional valor As String = "", Optional tipo As String = "") As String
    strSQL = "select   tf.*, tp.nome_pla, tc.nome_con"
    strSQL = strSQL & " From   tab_financeiro as tf left join tab_planocontas as tp "
    strSQL = strSQL & "  on tf.cod_pla=tp.cod_pla"
    
    strSQL = strSQL & " left join tab_contas as tc"
    
    strSQL = strSQL & " on tc.cod_con=tf.cod_con"
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
        strSQL = strSQL & " and tp.historico_fin='" & historico & "' "
    End If
    If valor <> "" Then
        strSQL = strSQL & " and valor_fin=" & F_GravaDec(valor)
    End If
    
    F_SelectFinanceiro = strSQL
End Function
Function F_SelectCai(seq_cai As String) As String
    Dim strSQL As String
    strSQL = "select tc.*,tp.nome_pla from " & TAB_CAIXA & " as tc left join " & TAB_PLANOCONTAS & " as tp "
    strSQL = strSQL & " on tc.cod_pla=tc.cod_pla and tp.tipo_pla=tc.tipo_cai  "
    strSQL = strSQL & " where 1 > 0 "
    If seq_cai <> "" Then
        strSQL = strSQL & " and tc.seq_cai='" & seq_cai & "' "
    End If
    strSQL = strSQL & " and excluido_cai=false "
    F_SelectCai = strSQL
    
    End Function
Function GravaFechaCaixa(cod_cai As String, data As String, dinheiro_total As String, dinheiro_info As String, _
    cheque_total As String, cheque_info As String, cartao_total As String, cartao_info As String) As Boolean

    
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into fecha_caixa(caixa,data,dinheiro_total,dinheiro_info,cheque_total,cheque_info,cartao_total,carta_info,"
    strSQL = strSQL & "quem_fechou,quando)values("
    strSQL = strSQL & "'" & cod_cai & "',"
    strSQL = strSQL & F_GravaData(data) & ","
    strSQL = strSQL & F_GravaDec(dinheiro_total) & ","
    strSQL = strSQL & F_GravaDec(dinheiro_info) & ","
    strSQL = strSQL & F_GravaDec(cheque_total) & ","
    strSQL = strSQL & F_GravaDec(cheque_info) & ","
    strSQL = strSQL & F_GravaDec(cartao_total) & ","
    strSQL = strSQL & F_GravaDec(cartao_info) & ","
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & "')"
    Conn.Execute strSQL
    
    GravaFechaCaixa = True
trata_Erro:
End Function

