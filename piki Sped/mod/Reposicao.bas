Attribute VB_Name = "Reposicao"
Type reposicao
    id As String
    abertura As String
    fechamento As String
    abresusu As String
    fechausu As String
    status As String
    
End Type

Function gravaReposicao(abertura As String, abreusu As String) As String
    Dim strSQL As String
    strSQL = "insert into tab_reposicao(abertura,abreusu,status)values("
    strSQL = strSQL & F_GravaDataTempo(abertura) & ","
    strSQL = strSQL & "'" & abreusu & "',"
    strSQL = strSQL & "1)"
    conn.Execute strSQL
    gravaReposicao = MaxKeySQL("tab_reposicao", "id", "", "")
    
    
End Function
Sub gravaReposicaoIT(reposicao As String, cod_pro As String, qtd As String)
    Dim strSQL As String
    strSQL = "insert into tab_reposicaoit (reposicao,cod_pro,qtd)values("
    strSQL = strSQL & reposicao & ","
    strSQL = strSQL & "'" & cod_pro & "',"
    strSQL = strSQL & F_GravaDec(qtd) & ")"
    conn.Execute strSQL
    
    
    
End Sub
Function selectReposicao(Optional status As String = "1", Optional dataI As String = "", _
    Optional dataF As String = "", Optional cod_pro As String = "", _
    Optional cod_mar As String = "", Optional cod_cat As String = "", _
    Optional cod_usu As String = "") As String
    Dim tipodata As String
    Dim strSQL As String
    strSQL = "select tr.*,tf.guerra_fun from tab_reposicao as tr "
    strSQL = strSQL & " left join tab_funcionarios as tf"
    strSQL = strSQL & " on tf.cod_fun=tr.abreusu"
    strSQL = strSQL & "left join tab_reposicaoit as ti"
    strSQL = strSQL & " on tr.id=ti.reposicao"
    strSQL = strSQL & " left join " & TAB_PRODUTOS & " as tp "
    strSQL = strSQL & " on tp.cod_pro=ti.cod_pro"
    strSQL = strSQL & " left join tab_categorias as tc "
    strSQL = strSQL & " on tp.cod_cat=tc.cod_cat"
    strSQL = strSQL & " left join " & TAB_MARCAS & " as tm "
    strSQL = strSQL & " on tp.cod_mar=tm.cod_mar "
    strSQL = strSQL & " where status <> 0 "
    If status <> "" Then
        strSQL = strSQL & " and status=" & status
    End If
    
    If dataI <> "" Then
        strSQL = strSQL & " and abertura>=" & F_GravaData(dataI)
    End If
    If dataF <> "" Then
        strSQL = strSQL & " and abertura<=" & F_GravaData(dataF)
    End If
    If cod_pro <> "" Then
        strSQL = strSQL & " and ti.cod_pro='" & cod_pro & "' "
    End If
    If cod_mar <> "" Then
        strSQL = strSQL & " and tm.cod_mar='" & cod_mar & "' "
    End If
    If cod_cat <> "" Then
        strSQL = strSQL & " and tc.cod_cat='" & cod_cat & "' "
    End If
    If cod_usu <> "" Then
        strSQL = strSQL & " and tr.abreusu='" & cod_usu & "' "
    End If
    selectReposicao = strSQL
End Function
    
    

