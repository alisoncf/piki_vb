Attribute VB_Name = "CAIXA"
Type esp_cai
    sigla As String
    descricao As String

End Type
Type caixa_logado
    codigo As String
    status  As String
    turno As String
End Type




Function AbreCaixa(cod_cai As String, cod_usu As String, valor_inicial As String) As String
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into abrecaixa(cod_cai,data,hora,cod_usu,valor_inicial)values("
    strSQL = strSQL & "'" & cod_cai & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & Format(Time, "HH:MM") & "',"
    strSQL = strSQL & "'" & cod_usu & "',"
    strSQL = strSQL & F_GravaDec(valor_inicial) & ")"
    Conn.Execute strSQL
    
    AbreCaixa = MaxKeySQL("abrecaixa", "id", "cod_cai='" & cod_cai & "'", "")
    
    
    Exit Function
trata_Erro:
    
End Function

Function FechaCaixa(cod_cai As String, cod_usu As String, valor_final As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into abrecaixa(cod_cai,data_fehcamento,hora_fechamento,fecha_usu,valor_final)values("
    strSQL = strSQL & "'" & cod_cai & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & Format(Time, "HH:MM") & "',"
    strSQL = strSQL & "'" & cod_usu & "',"
    strSQL = strSQL & F_GravaDec(valor_final) & ")"
    Conn.Execute strSQL
    
    FechaCaixa = True
    
    
    Exit Function
trata_Erro:
    
End Function

Function BuscaCaixaLogado(cod_cai As String) As caixa_logado
    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    
    
    
    strSQL = "select * from " & TAB_CAIXAS & " where cod_Cai='" & cod_cai & "' "
    abreObjetoAdo rPesq, strSQL
    
    With rPesq
        If .RecordCount > 0 Then
            BuscaCaixaLogado.codigo = !cod_cai
            BuscaCaixaLogado.status = !status_cai
            BuscaCaixaLogado.turno = !cod_tur
        End If
        .Close
    End With
    
End Function

Function GravaEspCai(sigla As String, descricao As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "insert into tab_espcai(sigla_cai,descricao_cai)values("
    strSQL = strSQL & "'" & sigla & "',"
    strSQL = strSQL & "'" & descricao & "')"
    Conn.Execute strSQL
    GravaEspCai = True
    Exit Function
trata_Erro:
    P_MostraErro
    GravaEspCai = False
End Function

Function AlteraEspCai(sigla As String, descricao As String) As Boolean
    On Error GoTo trata_Erro
    Dim strSQL As String
    strSQL = "update tab_espcai set  "
    strSQL = strSQL & "descricao='" & descricao & "' "
    
    strSQL = strSQL & " where sigla='" & sigla & "' "
    
    Conn.Execute strSQL
    AlteraEspCai = True
    Exit Function
trata_Erro:
    P_MostraErro
    AlteraEspCai = False
End Function


Function BuscaEspCai(desc As String) As esp_cai
    frmEdiEspCai.tag = "x"
    frmEdiEspCai.txtBusca.text = desc
    frmEdiEspCai.Show 1
    
    If frmEdiEspCai.tag <> "" Then
        BuscaEspCai.descricao = frmEdiEspCai.ado!descricao_cai
        BuscaEspCai.sigla = frmEdiEspCai.ado!sigla_cai
    End If
    Unload frmEdiEspCai
End Function

Public Sub carrega_caixa()
    strCod_Cai = F_BuscaUmCampo("cod_cai", "tab_caixasIP", "ip_cai='" & strHost_Usu & "'")
    
   
    'If strCod_Cai = "" Then
    '    strCod_Cai = F_BuscaUmCampo("cod_cai", "tab_caixasIP", "ip_cai='" & strIP_Usu & "'")
    'End If
   
    strSerie_Nfe = F_BuscaUmCampo("serie_nfe", "tab_caixas", "cod_Cai='" & strCod_Cai & "'")
    If strSerie_Nfe = "" Then
        strSerie_Nfe = "0"
    End If
   strCod_Tur = F_BuscaUmCampo("Cod_tur", "tab_caixas", "cod_Cai='" & strCod_Cai & "'")

End Sub
Function isLancamentoDuplicado(seq_cai As String) As Boolean




    Dim strSQL As String
    Dim rPesq As New ADODB.Recordset
    strSQL = "SELECT tc.*"
    strSQL = strSQL & " FROM tab_caixa tc"
    strSQL = strSQL & " Where tc.seq_cai ='" & seq_cai & "' "
    strSQL = strSQL & " AND EXISTS ("
    strSQL = strSQL & "  SELECT 1"
    strSQL = strSQL & "    FROM tab_caixa tc2"
    strSQL = strSQL & "  Where tc2.num_ven = tc.num_ven"
    strSQL = strSQL & " and tc.num_ven <> '' and tc.num_ven is not null"
    strSQL = strSQL & "      AND tc2.historico_cai = tc.historico_cai"
    strSQL = strSQL & "      AND tc2.valor_cai = tc.valor_cai"
    strSQL = strSQL & "      AND tc2.seq_cai <> tc.seq_cai"
    strSQL = strSQL & "  );"
    
    abreObjetoAdo rPesq, strSQL
    
    If rPesq.RecordCount > 0 Then
        isLancamentoDuplicado = True
    End If
End Function
Function excluir_lancamento(seq_cai As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "update " & TAB_CAIXA & " set excluido_cai=true "
    strSQL = strSQL & "  where seq_cai='" & seq_cai & "' "
    Conn.Execute strSQL
    excluir_lancamento = True
trata_Erro:
    
End Function
