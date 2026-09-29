Attribute VB_Name = "Principio"


Public Type principio
    cod_pri As String
    nome_pri As String
End Type
Function F_GravaPri(nome_pri As String) As String
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_Erro
    If F_BuscaUmCampo("nome_pri", "tab_principio", "nome_pri='" & Nome_Sec & "'") <> "" Then
        MsgBox "Já existe uma seção com este nome!", vbExclamation
        Exit Function
    End If
    sCod = MaxKeySQL("tab_principios", "cod_pri", "", "0000#")
    strSQL = "insert into tab_principios(cod_pri,nome_pri)values("
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_pri & "') "
    Conn.Execute strSQL
    F_GravaPri = sCod
    Exit Function
trata_Erro:
    P_MostraErro "F_GravaPri"
End Function
Function F_AlteraPri(cod_pri As String, nome_pri As String) As Boolean
    Dim strSQL As String
    F_AlteraPri = False
    
    On Error GoTo trata_Erro
    If F_BuscaUmCampo("nome_pri", "tab_principios", "nome_pri='" & nome_pri & _
        "'and cod_pri<>'" & cod_pri & "'") <> "" Then
        MsgBox "Já existe um princípio com este nome!", vbExclamation
        Exit Function
    End If
    sCod = MaxKeySQL("tab_principios", "cod_pri", "", "00#")
    strSQL = "update tab_principios set "
    strSQL = strSQL & "nome_pri='" & nome_pri & "' "
    strSQL = strSQL & " where cod_pri='" & cod_pri & "'"
    Conn.Execute strSQL
    F_AlteraPri = True
    Exit Function
trata_Erro:
    P_MostraErro "F_alteraPri"
End Function
Sub P_EditaPrincipio()
    frmEdiUmCampo.caption = "Edição de Princípios"
    frmEdiUmCampo.operacao = "2"
    frmEdiUmCampo.Show 0
    
End Sub
Function F_SelectPrincipio(Optional expr As String = "", Optional sOrder As String = "nome_pri") As String
    Dim strSQL As String
    strSQL = "select nome_pri as campo, cod_pri as cod from tab_principios"
    strSQL = strSQL & " where nome_pri like '" & expr & "%'  order by " & sOrder
    F_SelectPrincipio = strSQL
    
End Function

Sub P_PreenchePrincipios(obj As Object)
    P_Preenche obj, "tab_principios", "nome_pri"
End Sub
Function F_BuscaPri(expr As String) As principio.principio
    frmSelect.sSelect = "select cod_pri as cod,nome_pri as nome from tab_principios "
    frmSelect.cCriterio.Add " where nome_pri like '" & expr & "%'"
    frmSelect.sOrd = " order by nome_pri "
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.Col = 0
        
        F_BuscaPri.cod_pri = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.Col = 1
        
        F_BuscaPri.nome_pri = frmSelect.msflxgrd
        
    End If
    Unload frmSelect
    
End Function

