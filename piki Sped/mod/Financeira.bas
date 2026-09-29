Attribute VB_Name = "Financeira"
Type Financeira
    cod_fin As String
    nome_fin As String
    sigla_fin As String
    tipo_fin As String
    tel_fin As String
    operacao_fin As String
    CNPJ As String
    tBand As String
End Type

Type pagamentocartao
    cod_fin As String
    parcelas As String
    nome_fin As String
    autorizacao As String
    tipo_fin As String
    valor As String
    operacao As String
End Type
Type parcela_cartao
    num_ven As String
    cod_fin As String
    valor_fin As String
    reman_fin As String
    data_fin As String
    autorizacao_fin As String
    parcelas_fin As String
    parcela_fin As String
    vencimento_fin As String
    operacao_fin As String
    cod_usu As String
    tipo_fin As String
    
End Type
Function F_BuscaFinanceira(expr As String) As Financeira.Financeira
    frmEdiFinan.txtExpr.Text = expr
    
    frmEdiFinan.tag = "x"
    frmEdiFinan.Show 1
    
    
    
    
    If frmEdiFinan.tag <> "" Then
        F_BuscaFinanceira.nome_fin = frmEdiFinan.ado!nome_fin
        F_BuscaFinanceira.cod_fin = frmEdiFinan.ado!cod_fin
        
        F_BuscaFinanceira.operacao_fin = frmEdiFinan.ado!operacao_fin
        F_BuscaFinanceira.sigla_fin = frmEdiFinan.ado!sigla_fin
        F_BuscaFinanceira.tel_fin = TiraNull(frmEdiFinan.ado!tel_fin)
        F_BuscaFinanceira.tipo_fin = frmEdiFinan.ado!tipo_fin
        F_BuscaFinanceira.CNPJ = Format(frmEdiFinan.ado!CNPJ, "0#############")
        F_BuscaFinanceira.tBand = Format(frmEdiFinan.ado!tBand)
        
    End If
    Unload frmEdiFinan
    
End Function

    
Function F_VendaCartao(num_ven As String, valor_ven As String) As Boolean
    frmVendaCartao.lblNumVen.caption = num_ven
    frmVendaCartao.lblValor.caption = FormatVal(valor_ven)
    frmVendaCartao.Show 1
    If frmVendaCartao.tag <> "" Then
        With frmVendaCartao
            If F_GravaVenFin(num_ven, .txtbxFinanceira.tag, CCur(valor_ven), Date, _
                .txtbxAutorizacao.Text, .txtbxParc.Text, "V", "1", Date) Then
                F_VendaCartao = True
            Else
                F_VendaCartao = False
            End If
        End With
    Else
        F_VendaCartao = False
    End If

End Function


Function F_GravaVenFin(num_ven As String, cod_fin As String, valor_fin As Currency, _
    data_fin As String, autorizacao_fin As String, parcelas_fin As String, operacao_fin As String, parcela_fin As String, vencimento_fin As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    F_GravaVenFin = True
    If parcelas_fin = "" Then
        parcelas_fin = "1"
    End If
    strSQL = "insert into " & TAB_VENFIN & "(num_ven,cod_fin,valor_fin,data_fin,cod_usu,parcelas_Fin,"
    strSQL = strSQL & "autorizacao_fin,reman_fin,hora_fin,operacao_fin,parcela_fin,vencimento_fin)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & "'" & cod_fin & "',"
    strSQL = strSQL & F_GravaDec(valor_fin) & ","
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & strCod_Usu & "' ,"
    strSQL = strSQL & Int(parcelas_fin) & ","
    strSQL = strSQL & "'" & autorizacao_fin & "', "
    strSQL = strSQL & F_GravaDec(valor_fin) & ", "
    strSQL = strSQL & "'" & Format(Time, "hh:mm") & "', "
    strSQL = strSQL & "'" & operacao_fin & "', "
    strSQL = strSQL & parcela_fin & ","
    strSQL = strSQL & F_GravaData(vencimento_fin) & ")"
    conn.Execute strSQL
    Exit Function
trata_Erro:
    MsgBox Err.Description, vbCritical
    F_GravaVenFin = False
End Function


Function F_GravaBaixaVenFin(num_ven As String, valor_bai As String, esp_bai As String, seq_fin As String) As String
    Dim sSeq As String
    'On Error goto trata_erro
    sSeq = MaxKeySQL("baixavenfin", "seq_bai", "num_ven='" & num_ven & "'", "")
    strSQL = "insert into baixavenfin(num_ven,seq_bai,valor_bai,esp_bai,data_bai,seq_fin,cod_usu)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & F_GravaDec(valor_bai) & ","
    
    strSQL = strSQL & "'" & esp_bai & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & seq_fin & "',"
    strSQL = strSQL & "'" & strCod_Usu & "') "
    conn.Execute strSQL
    F_GravaBaixaVenFin = sSeq
trata_Erro:
End Function



Function F_DadosCartao(Optional cod_fin As String = "", Optional nome_fin As String = "", _
    Optional parc As String = "", Optional aut As String = "", Optional valor As String = "0,00") As pagamentocartao
    On Error Resume Next
    frmDadosCartao.txtbxFinanceira.tag = cod_fin
    frmDadosCartao.txtbxFinanceira.Text = nome_fin
    frmDadosCartao.txtbxParc.Text = parc
    frmDadosCartao.txtbxAutorizacao.Text = aut
    frmDadosCartao.txtbxValor.Text = valor
    
    frmDadosCartao.Show 1
    If frmDadosCartao.tag <> "" Then
        F_DadosCartao.operacao = frmDadosCartao.lblOperacao.caption
        F_DadosCartao.tipo_fin = frmDadosCartao.lblTipo.caption
        F_DadosCartao.cod_fin = frmDadosCartao.txtbxFinanceira.tag
        F_DadosCartao.nome_fin = frmDadosCartao.txtbxFinanceira.Text
        F_DadosCartao.parcelas = frmDadosCartao.txtbxParc.Text
        F_DadosCartao.autorizacao = frmDadosCartao.txtbxAutorizacao.Text
        F_DadosCartao.valor = frmDadosCartao.txtbxValor.Text
        
    End If
End Function

Function BuscaParcelaPorAutorizacao(autorizacao As String, parcela As String, valor As String) As Financeira.parcela_cartao
    Dim rPesq As New ADODB.Recordset
    Dim strSQL As String
    Dim p As parcela_cartao
    If valor = "" Then Exit Function
    strSQL = "select * from " & TAB_VENFIN & " where autorizacao_fin='" & autorizacao & "' "
    strSQL = strSQL & " and valor_fin>=" & F_GravaDec(CCur(valor) - 0.2)
    strSQL = strSQL & " and valor_fin<=" & F_GravaDec(CCur(valor) + 0.2)
    strSQL = strSQL & " and parcela_fin=" & parcela
    abreObjetoAdo rPesq, strSQL
    
    
    With rPesq
        If .RecordCount = 0 Then
            
            .Close
            Exit Function
        End If
        p.autorizacao_fin = !autorizacao_fin
        p.cod_fin = !cod_fin
        p.cod_usu = !cod_usu
        p.data_fin = !data_fin
        p.num_ven = !num_ven
        p.operacao_fin = !operacao_fin
        p.parcela_fin = !parcela_fin
        p.parcelas_fin = !parcelas_fin
        p.reman_fin = !reman_fin
        p.tipo_fin = !tipo_fin
        p.valor_fin = !valor_fin
        p.vencimento_fin = !vencimento_fin
    
        
        BuscaParcelaPorAutorizacao = p
    End With
    
End Function
Function Baixar_parcela_fin(num_ven As String, parcela_fin As String, valor As String, conta As String) As Boolean
    On Error GoTo trata_Erro:
    Dim strSQL As String
    Dim sSeq_Fin As String
    strSQL = "update " & TAB_VENFIN & "    set reman_fin=reman_fin-" & F_GravaDec(valor)
    strSQL = strSQL & " where num_ven='" & num_ven & "' "
    strSQL = strSQL & " and parcela_fin=" & parcela_fin
    conn.Execute strSQL
    sSeq_Fin = F_GravaFinanceiro("C", Date, valor, _
                "Compensação Financeira Automática " & num_ven & "/" & parcela_fin, conta, _
                 "N", "001")
                      
    
    Baixar_parcela_fin = True
    Exit Function
trata_Erro:
    Baixar_parcela_fin = False
End Function
