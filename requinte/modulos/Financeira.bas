Attribute VB_Name = "Financeira"
Type financeira
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
    CNPJ As String
    tBand As String
End Type

Function F_BuscaFinanceira(expr As String) As financeira.financeira
    frmEdiFinan.txtExpr.Text = expr
    
    frmEdiFinan.tag = "x"
    frmEdiFinan.Show 1
    
    
    
    
    If frmEdiFinan.tag <> "" Then
        F_BuscaFinanceira.nome_fin = frmEdiFinan.ado!nome_fin
        F_BuscaFinanceira.cod_fin = frmEdiFinan.ado!cod_fin
        
        F_BuscaFinanceira.operacao_fin = frmEdiFinan.ado!operacao_fin
        F_BuscaFinanceira.sigla_fin = frmEdiFinan.ado!sigla_fin
        F_BuscaFinanceira.tel_fin = frmEdiFinan.ado!tel_fin
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
                .txtbxAutorizacao.Text, .txtbxParc.Text, "V", "") Then
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
    data_fin As String, autorizacao_fin As String, parcelas_fin As String, operacao_fin As String, CNPJ As String, Optional msg As String = "") As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    F_GravaVenFin = True
    If parcelas_fin = "" Then
        parcelas_fin = "1"
    End If
    strSQL = "insert into  " & TAB_VENFIN & " (num_ven,cod_fin,valor_fin,data_fin,cod_usu,parcelas_Fin,"
    strSQL = strSQL & "autorizacao_fin,reman_fin,hora_fin,cnpj_fin,operacao_fin,mensagem)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & "'" & cod_fin & "',"
    strSQL = strSQL & F_GravaDec(valor_fin) & ","
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & strCod_Usu & "' ,"
    strSQL = strSQL & Int(parcelas_fin) & ","
    strSQL = strSQL & "'" & autorizacao_fin & "', "
    strSQL = strSQL & F_GravaDec(valor_fin) & ", "
    strSQL = strSQL & "'" & Time & "', "
    strSQL = strSQL & "'" & Format(CNPJ, "0000000000000#") & "', "
    strSQL = strSQL & "'" & operacao_fin & "', "
    strSQL = strSQL & "'" & F_AjustaTexto(msg, False) & "') "
    Conn.Execute strSQL
    Exit Function
trata_Erro:
    MsgBox Err.Description, vbCritical
    F_GravaVenFin = False
End Function


Function F_GravaBaixaVenFin(num_ven As String, valor_bai As String, esp_bai As String, seq_fin As String) As String
    Dim sSEQ As String
    'On Error goto trata_erro
    sSEQ = MaxKeySQL("tab_baixavenfin", "seq_bai", "num_ven='" & num_ven & "'", "")
    strSQL = "insert into tab_baixavenfin(num_ven,seq_bai,valor_bai,esp_bai,data_bai,seq_fin,cod_usu)values("
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & sSEQ & ","
    strSQL = strSQL & F_GravaDec(valor_bai) & ","
    
    strSQL = strSQL & "'" & esp_bai & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & seq_fin & "',"
    strSQL = strSQL & "'" & strCod_Usu & "') "
    Conn.Execute strSQL
    F_GravaBaixaVenFin = sSEQ
trata_Erro:
End Function



Function F_DadosCartao(Optional cod_fin As String = "", Optional nome_fin As String = "", _
    Optional parc As String = "1", Optional aut As String = "", Optional valor As String = "0,00") As pagamentocartao
    frmDadosCartao.txtbxFinanceira.tag = cod_fin
    frmDadosCartao.txtbxFinanceira.Text = nome_fin
    
    If parc = "" Then
        parc = "1"
    End If
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
        F_DadosCartao.CNPJ = frmDadosCartao.lblCnpj.caption
    End If
End Function
