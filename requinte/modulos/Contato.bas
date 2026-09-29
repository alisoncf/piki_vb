Attribute VB_Name = "Contato"
Type contato
    id As String
    nome As String
    telefone As String
End Type

Sub gravaContato(nome As String, telefone As String)
    Dim strSQL As String
    strSQL = "insert into contato(nome,telefone)values("
    strSQL = strSQL & "'" & nome & "',"
    strSQL = strSQL & "'" & telefone & "')"
    Conn.Execute strSQL
    
End Sub
Sub alteraContato(id As String, nome As String, telefone As String)
    Dim strSQL As String
    strSQL = "update contato set "
    strSQL = strSQL & "nome='" & nome & "',"
    strSQL = strSQL & "telefone='" & telefone & "' "
    strSQL = strSQL & " where id=" & id
    Conn.Execute strSQL

End Sub
Function BuscaContato(expr As String) As String
    Dim strSQL As String
    strSQL = "select * from contato where nome like '%" & expr & "%' or telefone like '%" & expr & "%' "
    BuscaContato = strSQL
End Function
