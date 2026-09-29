Attribute VB_Name = "subsecAO"
Public Type subsec
    cod_sub As String
    nome_sub As String
    sec As secao.secao
    descmax_sub As Currency
    comiss_sub As Currency
    
End Type


Function F_GravaSub(cod_sec As String, nome_sub As String, desc As String, comiss As String) As String
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_erro
    
    If F_BuscaUmCampo("cod_sub", "tab_subsec", "nome_sub='" & nome_sub & "' and cod_sec='" & cod_sec & "'") <> "" Then
'        MsgBox "Já existe uma subseção com este nome!", vbexe
        Exit Function
    End If
    
    
    
    sCod = MaxKeySQL("tab_subsec", "cod_sub", "", "00#")
    strSQL = "insert into tab_subsec(cod_sec,cod_sub,nome_sub,descmax_sub,comiss_sub)values("
    strSQL = strSQL & "'" & cod_sec & "',"
    strSQL = strSQL & "'" & sCod & "',"
    strSQL = strSQL & "'" & nome_sub & "', "
    strSQL = strSQL & F_GravaDec(desc) & ","
    strSQL = strSQL & F_GravaDec(comiss) & ")"
    Conn.Execute strSQL
    F_GravaSub = sCod
    Exit Function
trata_erro:
    P_MostraErro
End Function
Sub P_GravaSubSecComCodigo(cod_sec As String, cod_sub As String, nome_sub As String)
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_erro
    
    
    
    
    strSQL = "insert into tab_subsec(cod_sec,cod_sub,nome_sub)values("
    strSQL = strSQL & "'" & cod_sec & "',"
    strSQL = strSQL & "'" & cod_sub & "',"
    strSQL = strSQL & "'" & nome_sub & "') "
    
    Conn.Execute strSQL
    
    Exit Sub
trata_erro:
    P_MostraErro
End Sub

Function F_AlteraSub(cod_sub As String, cod_sec As String, nome_sub As String, desc As String, comiss As String) As Boolean
    Dim strSQL As String
    Dim sCod As String
    On Error GoTo trata_erro
    F_AlteraSub = False
    
    If F_BuscaUmCampo("cod_sub", "tab_subsec", "nome_sub='" & nome_sub & "' and cod_sub<>'" & cod_sub & "'") <> "" Then
        MsgBox "Já existe uma subseção com este nome!", vbexe
        Exit Function
    End If
    strSQL = "update tab_subsec set "
    strSQL = strSQL & "cod_sec='" & cod_sec & "',"
    strSQL = strSQL & "comiss_sub=" & F_GravaDec(comiss) & ","
    strSQL = strSQL & "descmax_sub=" & F_GravaDec(desc) & ","
    strSQL = strSQL & "nome_sub='" & nome_sub & "'  "
    strSQL = strSQL & " Where cod_sec='" & cod_sec & "' "
    strSQL = strSQL & " and cod_sub='" & cod_sub & "' "
    Conn.Execute strSQL
    F_AlteraSub = True
    Exit Function
trata_erro:
    P_MostraErro
End Function
Sub P_PreencheSubSec(obj As Object, cod_sec As String)
    P_Preenche obj, "tab_subsec", "nome_sub", IIf(cod_sec <> "", "cod_sec='" & cod_sec & "'", "")
End Sub

Function F_SelectSubSec(Optional where As String = "", Optional sOrder As String = "nome_sub") As String
    Dim strSQL As String
    strSQL = "select distinct tu.*,ts.nome_sec from tab_secoes as ts,tab_subsec as tu "
    strSQL = strSQL & " where tu.cod_sec=ts.cod_sec "
    If where <> "" Then
        strSQL = strSQL & " and " & where
    End If
    If sOrder <> "" Then
        strSQL = strSQL & " order by " & sOrder
    End If
    F_SelectSubSec = strSQL
End Function
Function F_BuscaSubSec(cod_sec As String, expr As String) As subsec
    frmSelect.sSelect = "select cod_sub as cod,nome_sub as nome from tab_subsec "
    frmSelect.cCriterio.Add " where nome_sub like '" & expr & "%' and cod_sec='" & cod_sec & "' "
    frmSelect.sOrd = " order by nome_sub "
    frmSelect.txtbxExpr.Text = expr
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.Col = 0
        
        F_BuscaSubSec.cod_sub = frmSelect.msflxgrd
                
        frmSelect.msflxgrd.Col = 1
        
        F_BuscaSubSec.nome_sub = frmSelect.msflxgrd
        
    End If
    Unload frmSelect
    
End Function

