Attribute VB_Name = "contagem"
Function GravaContagem(cod_pro As String, qtd_ant As String, qtd As String) As Boolean
    Dim strSQL As String
    
    
    On Error GoTo trata_erro
    
    
    strSQL = "insert into contagem(cod_pro,qtd_ant,qtd,data,hora,cod_usu)values("
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(qtd_ant) & ","
    strSQL = strSQL & F_GravaDec(qtd) & ","
    strSQL = strSQL & F_GravaData(Date) & ","
    strSQL = strSQL & "'" & Time & "',"
    strSQL = strSQL & "'" & strCod_Usu & "')"
    Conn.Execute strSQL
    
    
    
    GravaContagem = True
    Exit Function
trata_erro:
    GravaContagem = False
    
    
End Function
