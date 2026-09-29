Attribute VB_Name = "convenio"
Public Type convenio
    cod_con As String
    cod_cli As String
    data_con As String
    status_con As String
    nome_con As String
End Type


Function F_GravaConv(cod_cli As String, desconto_con As String) As String
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    
    
    strSQL = "insert into tab_convenios(cod_cli,data_con,desconto_con)values("
    
    
    strSQL = strSQL & "'" & cod_cli & "',"
    strSQL = strSQL & F_GravaData(Date) & ", "
    strSQL = strSQL & F_GravaDec(desconto_con) & ") "
    Conn.Execute strSQL
    F_GravaConv = sCod
    
    
    Exit Function
trata_Erro:
    
    
End Function



Function F_AlteraConv(cod_cli As String, desconto_con As String, status_con As String) As Boolean
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    
    
    strSQL = "update tab_convenios set "

    strSQL = strSQL & "desconto_con=" & F_GravaDec(desconto_con) & ","

    strSQL = strSQL & " status_Con='" & status_con & "'"
    strSQL = strSQL & " where cod_cli='" & cod_con & "' "
    Conn.Execute strSQL
    F_AlteraConv = True
    
    
    Exit Function
    
trata_Erro:
    F_AlteraConv = False
    
End Function
