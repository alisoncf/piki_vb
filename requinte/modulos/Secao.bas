Attribute VB_Name = "Secao"


Public Type secao
    cod_sec As String
    Nome_Sec As String
End Type
Function F_GravaSec(Nome_Sec As String) As String
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    If F_BuscaUmCampo("nome_sec", "tab_secoes", "nome_sec='" & Nome_Sec & "'") <> "" Then
        MsgBox "Já existe uma seção com este nome!", vbExclamation
        Exit Function
    End If
    sCod = MaxKeySQL("tab_secoes", "cod_sec", "", "00#")
    strSQL = "insert into tab_secoes(cod_sec,nome_sec)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & Nome_Sec & "') "
    Conn.Execute strSQL
    F_GravaSec = sCod
    Exit Function
trata_Erro:
    P_MostraErro "F_GravaSec"
End Function
Sub P_GravaSecComCodigo(cod_sec As String, Nome_Sec As String)
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    
    
    strSQL = "insert into tab_secoes(cod_sec,nome_sec)values("
    strSQL = strSQL & "'" & cod_sec & "',"
    strSQL = strSQL & "'" & Nome_Sec & "') "
    Conn.Execute strSQL
    
    Exit Sub
trata_Erro:
    P_MostraErro "F_GravaSec"
End Sub


Function F_AlteraSec(cod_sec As String, Nome_Sec As String) As Boolean
    Dim strSQL As String
    F_AlteraSec = False
    On Error GoTo trata_Erro
    If F_BuscaUmCampo("nome_sec", "tab_secoes", "nome_sec='" & Nome_Sec & _
        "'and cod_sec<>'" & cod_sec & "'") <> "" Then
        MsgBox "Já existe uma seção com este nome!", vbExclamation
        Exit Function
    End If
    sCod = MaxKeySQL("tab_secoes", "cod_sec", "", "00#")
    strSQL = "update tab_secoes set "
    strSQL = strSQL & "nome_sec='" & Nome_Sec & "' "
    strSQL = strSQL & " where cod_sec='" & cod_sec & "'"
    Conn.Execute strSQL
    F_AlteraSec = True
    Exit Function
trata_Erro:
    P_MostraErro "F_GravaSec"
End Function
Sub P_EditaSecao()
    frmEdiUmCampo.caption = "Edição de Seções"
    frmEdiUmCampo.operacao = secao_
    frmEdiUmCampo.Show 0, mdiFrm
    
End Sub
Function F_SelectSecao(Optional expr As String = "", Optional sOrder As String = "nome_sec") As String
    Dim strSQL As String
    strSQL = "select nome_sec as campo, cod_sec as cod from tab_secoes "
    strSQL = strSQL & " where nome_sec like '" & expr & "%'  order by " & sOrder
    F_SelectSecao = strSQL
    
End Function

Sub P_PreencheSecao(obj As Object)
If IsNull(obj) = False Then
    P_Preenche obj, "tab_secoes", "nome_sec"
End If
End Sub
Function F_BuscaSec(expr As String) As secao.secao
    
    frmEdiSec.txtExpr.Text = expr
    frmEdiSec.tag = "x"
    frmEdiSec.Show 1
    On Error Resume Next
    If frmEdiSec.tag <> "" Then
        
        
        F_BuscaSec.cod_sec = frmEdiSec.ado!cod_sec
                
        
        
        F_BuscaSec.Nome_Sec = frmEdiSec.ado!Nome_Sec
        
    End If
    Unload frmEdiSec
    
End Function
