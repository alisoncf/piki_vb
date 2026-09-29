Attribute VB_Name = "Cartoes"
Sub gravaCartao_Movimento(dia As String, coo As String, doc As String, autorizacao As String, valor As String, parcelas As String, conferido As String, financeir As String)
    Dim strSQL As String
    
    
    strSQL = "insert into cartao_movimento(dia,usuario,coo,doc,autorizacao,valor,parcelas,conferido,financeira)values("
    strSQL = strSQL & F_GravaData(dia) & ","
    strSQL = strSQL & strCod_Usu & ","
    strSQL = strSQL & "'" & coo & "',"
    strSQL = strSQL & "'" & doc & "',"
    strSQL = strSQL & "'" & autorizacao & "',"
    strSQL = strSQL & F_GravaDec(valor) & ","
    strSQL = strSQL & parcelas & ","
    strSQL = strSQL & "'" & financeir & "', "
    strSQL = strSQL & "'" & conferido & "') "
    
    Conn.Execute strSQL
End Sub
Sub AlteraCartao_Movimento(id As String, dia As String, coo As String, doc As String, autorizacao As String, valor As String, parcelas As String, conferido As String, financeira As String)
    Dim strSQL As String
    
    
    strSQL = "update cartao_movimento set "
    strSQL = strSQL & "dia=" & F_GravaData(dia) & ","
    strSQL = strSQL & "usuario=" & strCod_Usu & ","
    strSQL = strSQL & "coo='" & coo & "',"
    strSQL = strSQL & "doc='" & doc & "',"
    strSQL = strSQL & "autorizacao='" & autorizacao & "',"
    strSQL = strSQL & "valor=" & F_GravaDec(valor) & ","
    strSQL = strSQL & "parcelas=" & parcelas & ","
    strSQL = strSQL & "financeira=" & financeira & ","
    strSQL = strSQL & "conferido='" & conferido & "' "
    strSQL = strSQL & " where id=" & id
    
    Conn.Execute strSQL
End Sub
Function SelectCartao_Movimento(de As String, ate As String, Optional bandeira As String, _
Optional valor As String, Optional doc As String) As String
    Dim strSQL As String
    strSQL = "select tf.nome_fin,c.* from cartao_movimento as tc left join tab_financeiras as tf "
    strSQL = strSQL & " on tf.cod_fin=c.financeira "
    
    strSQL = strSQL & " where dia>=" & F_GravaData(de)
    
    If ate <> "" Then
        strSQL = strSQL & " and dia <=" & F_GravaData(ate)
    End If
    
    If valor <> "" Then
        strSQL = strSQL & " and valor=" & F_GravaDec(valor)
    End If
    If doc <> "" Then
        strSQL = strSQL & " and doc like '%" & doc & "%'"
    End If
    strSQL = strSQL & " order by dia,seq "
    SelectCartao_Movimento = strSQL
End Function


