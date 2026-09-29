Attribute VB_Name = "Reposicao"


Function F_GeraNovaReposicao() As String
    Dim sNum As String
    On Error GoTo trata_erro
    Dim strSQL As String
    sNum = MaxKeySQL("tab_reposicao", "seq_rep", "", "000000#")
    strSQL = "insert into tab_reposicao(seq_rep,solicitante_rep,abertura_rep,status_rep)values("
    strSQL = strSQL & "'" & sNum & "',"
    strSQL = strSQL & "'" & strCod_Usu & "',"
    strSQL = strSQL & F_GravaDataTempo(Now) & ","
    
    strSQL = strSQL & "'1') "
    conn.Execute strSQL
    
    F_GeraNovaReposicao = sNum
    
    
    
    Exit Function
trata_erro:
    P_MostraErro "Erro na criação da reposição, tente novamente!"
End Function

Sub P_AlteraReposicao(seq_rep As String, status As String, Optional fechamento As Boolean = False, Optional recepcao As Boolean = False)
    Dim sNum As String
    On Error GoTo trata_erro
    Dim strSQL As String
    strSQL = "update tab_reposicao set "
    strSQL = strSQL & "solicitante_rep='" & strCod_Usu & "',"
    If fechamento Then
        strSQL = strSQL & "fechamento_rep=" & F_GravaDataTempo(Now) & ","
    End If
    If recepcao Then
        strSQL = strSQL & "atendente_rep='" & strCod_Usu & "',"
        strSQL = strSQL & "atendimento_rep=" & F_GravaDataTempo(Now) & ","
    End If
    strSQL = strSQL & "status_rep='" & status & "' "
    strSQL = strSQL & " where seq_rep='" & seq_rep & "' "
    
    
    conn.Execute strSQL
    
    
    
    
    
    Exit Sub
trata_erro:
    P_MostraErro "Erro na criação da reposição, tente novamente!"
End Sub
Function F_BuscaItensRep(seq_rep As String, Optional sOrd As String = "seq_ite", Optional sWhere As String = "") As String
    Dim strSQL As String
    strSQL = "select ts.nome_sec, tu.nome_sub, "
        
    strSQL = strSQL & "ti.qtd_ite*tp.pv_pro as total,"
    strSQL = strSQL & " tp.st_pro,tp.nome_pro,tp.ref_pro,tp.endereco_pro,tp.pv_pro,tp.pq_pro,tp.pv_pro*qtd_ite as total,ti.* "
    strSQL = strSQL & " from tab_produtos as tp, "
    strSQL = strSQL & " tab_subsec as tu,tab_secoes as ts, "
    strSQL = strSQL & " tab_itrep as ti "

    
    strSQL = strSQL & " where tp.cod_sub=tu.cod_sub "
    strSQL = strSQL & " and tu.cod_sec=ts.cod_sec "
    
    
    strSQL = strSQL & " and tp.cod_pro=ti.cod_pro "
    
    strSQL = strSQL & " and ti.seq_rep='" & seq_rep & "' "
    If sWhere <> "" Then
        strSQL = strSQL & " and " & sWhere
    End If
    If sOrd <> "" Then
        strSQL = strSQL & " order by " & sOrd
    End If
    F_BuscaItensRep = strSQL
End Function

Function F_GravaItemRep(seq_rep As String, cod_pro, qtd_ite) As String
    Dim sSeq As String
    sSeq = MaxKeySQL("tab_itrep", "seq_rep", "seq_rep='" & seq_rep & "'", "")
    
    If sSeq = "" Then sSeq = "1"
    
    strSQL = "insert into tab_itrep(seq_rep,seq_ite,cod_pro,qtd_ite)values("
    strSQL = strSQL & "'" & seq_rep & "',"
    strSQL = strSQL & sSeq & ","
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(qtd_ite) & ")"
    conn.Execute strSQL
    F_GravaItemRep = sSeq
    
End Function
Sub F_DeletaitemRep(seq_rep As String, seq_ite As String)
    Dim strSQL As String
    strSQL = "delete from tab_itrep where seq_rep='" & seq_rep & "' "
    strSQL = strSQL & " and seq_ite=" & seq_ite & " "
    conn.Execute strSQL
End Sub
