Attribute VB_Name = "vales"
Type vale
    id As String
    cod_fun As String
    
End Type

Function gravaVale(cod_fun As String, historico As String, valor As String, data As String, referencia As String, _
    cod_usu As String, parcela As String, Optional num_ven As String) As Boolean
    Dim strSQL As String
    On Error GoTo trata_Erro
    strSQL = "insert into vales(cod_fun,historico,valor,data,referencia,cod_usu,num_ven,parcela)values("
    strSQL = strSQL & "'" & cod_fun & "',"
    strSQL = strSQL & "'" & historico & "',"
    strSQL = strSQL & F_GravaDec(valor) & ", "
    strSQL = strSQL & F_GravaData(data) & ", "
    strSQL = strSQL & F_GravaData(referencia) & ", "
    strSQL = strSQL & "'" & cod_usu & "',"
    strSQL = strSQL & "'" & num_ven & "',"
    strSQL = strSQL & "'" & parcela & "')"
    
    Conn.Execute strSQL
    gravaVale = True
trata_Erro:
    
    
End Function
Function LancaVale(cod_fun As String, valor As String, parcelas As String, data_base As String, _
    historico As String, Optional venda As String) As Boolean
    Dim i As Integer
    For i = 1 To Int(parcelas)
        If gravaVale(cod_fun, historico, CStr(CCur(valor) / Int(parcelas)), data_base, _
            DateAdd("M", i - 1, data_base), strCod_Usu, i & "/" & parcelas, venda) Then
            LancaVale = True
        Else
        
            LancaVale = False
        End If
    Next
End Function
Function GravaValeVenda(num_ven As String, valor As String) As String
    frmLancaVale.tag = num_ven
    frmLancaVale.txtDataRef.Text = Date
    frmLancaVale.txtHistorico.Text = "Venda " & num_ven & " dia " & Date & " valor: " & FormatVal(valor)
    frmLancaVale.txtHistorico.Locked = True
    frmLancaVale.txtValor.Text = FormatVal(valor)
    frmLancaVale.txtValor.Locked = True
    frmLancaVale.txtValorTotal.Text = FormatVal(valor)
    
    
    
    GravaValeVenda = frmLancaVale.tag
    Unload frmLancaVale


End Function
