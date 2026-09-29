Attribute VB_Name = "Permissao"

Type permissao
    codigo As String
    nome As String
    
End Type
Type grupoPer
    codigo As String
    nome As String
End Type

    
Global Permissao_usu As permissao
    
Function F_GravaPermissao(nome_per As String, obs_per As String, cod_gru As String) As String
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_erro
    sSeq = MaxKeySQL("tab_permissoes", "cod_per", "", "")
    strSQL = "insert into tab_permissoes(cod_per,nome_per,obs_per,cod_gru)values("
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & nome_per & "',"
    strSQL = strSQL & "'" & obs_per & "',"
    strSQL = strSQL & "'" & cod_gru & "')"
    conn.Execute strSQL
    
    F_GravaPermissao = sSeq
    Exit Function
trata_erro:
    
    
End Function

    
Sub F_GravaPermissaoUsu(cod_usu As String, cod_per As String)
    Dim strSQL As String
    Dim sSeq As String
    On Error GoTo trata_erro
    
    strSQL = "insert into tab_permissoes(cod_per,cod_usu)values("
    strSQL = strSQL & cod_per & ","
    strSQL = strSQL & "'" & cod_usu & "')"
    
    conn.Execute strSQL
    
    
    Exit Sub
trata_erro:
    
    
End Sub
Function F_SelectPermissoes(sExpr As String, Optional cod_gru As String = "") As String
    Dim strSQL As String
    strSQL = "select tp.*,tg.nome_gru from tab_permissoes as tp,"
    strSQL = strSQL & " tab_permgru as tg "
    strSQL = strSQL & " where tg.cod_gru=tp.cod_gru "
    If cod_gru <> "" Then
        strSQL = strSQL & " and tp.cod_gru='" & cod_gru & "'  "
    End If
    If sExpr <> "" Then
        strSQL = strSQL & " and nome_per like '%" & sExpr & "%'  "
    End If
    F_SelectPermissoes = strSQL
    
End Function

Function F_SelectPermissoesUsu(cod_usu As String, Optional sExpr As String = "", Optional cod_gru As String = "") As String
    Dim strSQL As String
    strSQL = "select tp.*,tg.nome_gru from tab_permissoes as tp,tab_permusu as tu "
    strSQL = strSQL & " tab_permgru as tg "
    strSQL = strSQL & " where tg.cod_gru=tp.cod_gru "
    strSQL = strSQL & " and tp.cod_per=tu.cod_per "
    strSQL = strSQL & " and tu.cod_usu='" & cod_usu & "' "
    strSQL = strSQL & " and nome_per like %'" & sExpr & "'%"
    
    If cod_gru <> "" Then
        strSQL = strSQL & " and tp.cod_gru='" & cod_gru & "'  "
    End If
    F_SelectPermissoesUsu = strSQL
    
End Function
Function F_BuscaGrupoPer(sExpr As String) As grupoPer
    Dim strSQL As String
    If IsLoad("frmSelect") Then
        Unload frmSelect
    End If
    strSQL = "select cod_gru as cod,nome_gru as Grupo from tab_permgru"
    frmSelect.sOrd = "order by nome_gru"
    frmSelect.sSelect = strSQL
    frmSelect.cCriterio.Add " where cod_gru<>'x'"
    
    frmSelect.Show 1
    
    If frmSelect.tag <> "" Then
        frmSelect.msflxgrd.Col = 0
        F_BuscaGrupoPer.codigo = frmSelect.msflxgrd.Text
        frmSelect.msflxgrd.Col = 1
        F_BuscaGrupoPer.nome = frmSelect.msflxgrd.Text
    End If
    Unload frmSelect
End Function
