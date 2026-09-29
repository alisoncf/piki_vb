Attribute VB_Name = "Caixa"
Type caixa_logado
    codigo As String
    status  As String
    turno As String
End Type

Type esp_cai
    sigla As String
    descricao As String

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
    conn.Execute strSQL
    
    AbreCaixa = MaxKeySQL("abrecaixa", "id", "cod_cai='" & cod_cai & "'", "")
    
    
    Exit Function
trata_Erro:
    
End Function
Sub P_PreencheEspCai(cmbx As Object)
    Dim sWh As String
    
    
    P_Preenche cmbx, "tab_espcai", "concat(sigla_cai,'-',descricao_cai)", "", "", True
    
End Sub
Function FechaCaixa(cod_cai As String, cod_usu As String, valor_final As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into abrecaixa(cod_cai,data_fehcamento,hora_fechamento,fecha_usu,valor_final)values("
    strSQL = strSQL & "'" & cod_cai & "',"
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & Format(Time, "HH:MM") & "',"
    strSQL = strSQL & "'" & cod_usu & "',"
    strSQL = strSQL & F_GravaDec(valor_final) & ")"
    conn.Execute strSQL
    
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
    conn.Execute strSQL
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
    
    conn.Execute strSQL
    AlteraEspCai = True
    Exit Function
trata_Erro:
    P_MostraErro
    AlteraEspCai = False
End Function


Function BuscaEspCai(desc As String) As esp_cai
    frmEdiEspCai.tag = "x"
    frmEdiEspCai.txtBusca.Text = desc
    frmEdiEspCai.Show 1
    
    If frmEdiEspCai.tag <> "" Then
        BuscaEspCai.descricao = frmEdiEspCai.ado!descricao_cai
        BuscaEspCai.sigla = frmEdiEspCai.ado!sigla_cai
    End If
    Unload frmEdiEspCai
End Function
